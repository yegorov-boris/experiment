import * as fs from 'fs';
import { promisify } from 'util';
import { exec } from 'child_process';
import zipdir from 'zip-dir';
import { parseStringPromise, Builder } from 'xml2js';
import { find } from 'xml2js-xpath';
import { convert } from 'libreoffice-convert';
import rimraf from 'rimraf';
import mkdirp from 'mkdirp';
import emailRegex from 'email-regex';

const docx2pdf = promisify(convert);
const execP = promisify(exec);
const rimrafP = promisify(rimraf);
const fsP = fs.promises;

const safeLen = x => x?.length || 0;

function restoreOrder(doc, docOrdered) {
  const matches = doc['w:document']['w:body'][0]['w:p'];
  if (!safeLen(matches)) return;
  const matchesOrdered = docOrdered['w:document']['w:body'][0]['w:p'];
  matches.forEach((m, i) => {
    const rs = m['w:r'];
    if (!rs) return;
    const rsNew = [];
    rs.forEach((r, j) => {
      if (safeLen(r['w:t']) < 2) {
        rsNew.push(r);
        return
      }
      let k = 0;
      let rNew = {
        ...r,
        'w:t': [""]
      };
      delete rNew['w:br'];
      matchesOrdered[i]['w:r'][j]['$$'].forEach(v => {
        const name = v['#name'];
        if (name === 'w:br') {
          rNew['w:br'] = rNew['w:br'] || [];
          rNew['w:br'].push("")
        }
        if (name === 'w:t') {
          if (rNew['w:t'][0]) {
            rsNew.push({...rNew});
            rNew = {
              ...r,
              'w:t': [r['w:t'][k]]
            };
            delete rNew['w:br'];
          } else {
            rNew['w:t'] = [r['w:t'][k]]
          }
          k++
        }
      });
      rsNew.push({...rNew});
    });
    m['w:r'] = rsNew;
  });

  const body = doc['w:document']['w:body'][0];
  if (!safeLen(body['w:tbl']) || !safeLen(body['w:p'])) return;

  let p = 0;
  let t = 0;
  docOrdered['w:document']['w:body'][0]['$$'].forEach(tag => {
    const name = tag['#name'];
    if (name === 'w:p') {
      p++;
      return
    }

    body['w:p'].splice(p, 0, {'w:tbl': {...body['w:tbl'][t]}});
    t++
  });
  delete body['w:tbl'];
}

function matchesToText(matches) {
  let joined = '';
  const borders = [];
  let left = 0;
  matches.forEach((run, j) => {
    if (!run['w:t']) return;

    run['w:t'].forEach((t, k) => {
      let part = '';

      if (typeof t === 'string') {
        part = t
      } else if (typeof t['_'] === 'string') {
        part = t['_']
      } else if (t['$'] && t['$']['xml:space'] === 'preserve') {
        part = ' '
      } else {
        return
      }

      joined += part;
      const right = left + part.length;
      borders.push([j, k, left, right]);
      left = right;
    });
  });
  const lr = x => {
    const i = joined.indexOf(x);
    return [i, i + x.length]
  };
  const emails = (joined.match(emailRegex()) || []).map(lr);
  const parts = joined.split(' ');
  const urls = parts
    .filter(u => {
      try {
        new URL(u);
        return true
      } catch {
        if (u.slice(0, 4) !== 'www.') return false;
        try {
          new URL(`http://${u}`);
          console.log(joined);
          return true
        } catch {
          return false
        }
      }
    })
    .map(lr);

  return {
    joined,
    borders,
    escape: emails.concat(urls)
  }
}

function applyTranslation(paragraph, text, translatedText, filterPunctuation) {
  const matches = find(paragraph, "//w:r");
  const punctuation = /[.,\/#!$%\^&\*;:{}=\-_`~()"']/g;
  const {joined, borders, escape} = matchesToText(matches);
  let filteredJoined, filteredText, rel;

  if (filterPunctuation) {
    filteredJoined = joined.replace(punctuation, '');
    let pj = 0;
    rel = filteredJoined.split('').map((letter, pf) => {
      if (pf === 0) return 0;
      if (pf === filteredJoined.length - 1) return joined.length - 1;
      while (joined[pj] !== letter) pj++;
      return pj
    });
    filteredText = text.replace(punctuation, '');
  } else {
    filteredJoined = joined;
    filteredText = text;
    rel = filteredJoined.split('').map((_, pf) => pf);
  }

  const foundF = filteredJoined.indexOf(filteredText);
  if (foundF === -1) return false;

  const foundL = rel[foundF];
  const foundR = rel[foundF + filteredText.length - 1] + 1;

  for (const [el, er] of escape) {
    if (el <= foundL && er >= foundR) return false;
  }

  const found = joined.slice(foundL, foundR);

  for (const [j, k, l, r] of borders) {
    if (r <= foundL || l >= foundR) continue;

    if (joined.slice(l, r).includes(found)) {
      if (typeof matches[j]['w:t'][k] === 'string') {
        matches[j]['w:t'][k] = matches[j]['w:t'][k].replace(found, translatedText)
      } else {
        matches[j]['w:t'][k]['_'] = matches[j]['w:t'][k]['_'].replace(found, translatedText)
      }
      break
    }

    if (foundL < l && foundR > r) {
      if (typeof matches[j]['w:t'][k] === 'string') {
        matches[j]['w:t'][k] = ''
      } else {
        matches[j]['w:t'][k]['_'] = ''
      }
      continue
    }

    if (foundL >= l && foundL < r) {
      if (typeof matches[j]['w:t'][k] === 'string') {
        matches[j]['w:t'][k] = matches[j]['w:t'][k].slice(0, foundL - l) + translatedText
      } else {
        matches[j]['w:t'][k]['_'] = matches[j]['w:t'][k]['_'].slice(0, foundL - l) + translatedText
      }
      continue
    }

    if (typeof matches[j]['w:t'][k] === 'string') {
      matches[j]['w:t'][k] = matches[j]['w:t'][k].slice(foundR - l)
    } else {
      matches[j]['w:t'][k]['_'] = matches[j]['w:t'][k]['_'].slice(foundR - l)
    }
  }

  return true
}

function applyTranslations(translates, doc, filterPunctuation) {
  translates.forEach(({text, translatedText}) => {
    if (!translatedText) return;
    find(doc, "//w:p").forEach(paragraph => {
      if (paragraph['w:tbl']) {
        applyTranslations(translates, paragraph, filterPunctuation)
      } else {
        while (applyTranslation(paragraph, text, translatedText, filterPunctuation)) {}
      }
    });
  });
}

async function processDocx(fileType, id, data, filterPunctuation=true) {
  const dir = `./uploads/${id}/`;
  await mkdirp(dir);
  const pathDefault = `${dir}default.${fileType}`;
  let pathDocx = pathDefault;
  await fsP.writeFile(pathDefault, data.buffer);

  if (fileType === 'pdf') {
    pathDocx = `${dir}default.docx`;
    const tmp = await execP(`pdf2docx convert ${pathDefault} ${pathDocx}`);
    console.log(tmp);
  }

  const resps = [];

  for (const {translates} of data.Text) {
    const dirTo = `${dir}${Date.now()}`;
    await execP(`unzip ${pathDocx} -d ${dirTo}`);
    const path = `${dirTo}/word/document.xml`;
    const xmlData = await fsP.readFile(path);
    const doc = await parseStringPromise(xmlData.toString());
    const docOrdered = await parseStringPromise(xmlData.toString(), {
      preserveChildrenOrder: true,
      explicitChildren: true
    });
    restoreOrder(doc, docOrdered);
    applyTranslations(
      [...translates].sort((x, y) => safeLen(y?.text) - safeLen(x?.text)),
      doc,
      filterPunctuation
    );
    const builder = new Builder();
    const updatedXML = builder.buildObject(doc);
    await fsP.writeFile(path, updatedXML);
    const buffer = await zipdir(`${dirTo}`);

    if (fileType === 'pdf') {
      const bufferPdf = await docx2pdf(buffer, 'pdf', undefined);
      resps.push(bufferPdf)
    } else {
      resps.push(buffer)
    }
  }

  await rimrafP(dir);

  return resps
}

async function testDocx(path, tr, outName, filterPunctuation=true) {
  const dataDocx = await fsP.readFile(path);
  const data = {
    buffer: dataDocx,
    Text: tr
  };
  const resultDocx = await processDocx('docx', 1, data, filterPunctuation);

  await Promise.all(resultDocx.map((buffer, i) => fsP.writeFile(`${outName}${i}.docx`, buffer)));
}

async function testPdf(path, tr, outName, filterPunctuation=true) {
  const dataPdf = await fsP.readFile(path);
  const data = {
    buffer: dataPdf,
    Text: tr
  };
  const resultPdf = await processDocx('pdf', 2, data, filterPunctuation);
  await Promise.all(resultPdf.map((buffer, i) => fsP.writeFile(`${outName}${i}.pdf`, buffer)));
}

const trs = [
  [
    {
      translates: [
        {
          text: 'Mount Everest is Earth\'s highest mountain above sea level, located in the Mahalangur Himal',
          translatedText: 'Гора Эверест - самая высокая гора на Земле над уровнем моря расположенная в Махалангурских Гималах'
        },
      ]
    },
  ],
  [
    {
      translates: [
        {
          text: 'While not posing substantial technical climbing challenges on the standard route, Everest presents dangers such as altitude sickness weather and wind as well as significant hazards from avalanches and the Khumbu Icefal',
          translatedText: 'Не создавая серьезных технических проблем для поднятия на стандартном пути, Эверест несет в себе такие опасности, как высотная болезнь, погодные условия'
        },
        {
          text: 'There are two main climbing routes one approaching the summit from the southeast in Nepal known as the standard route and the other from the north in Tibet',
          translatedText: 'Есть два основных маршрута для восхождения: один приближается к вершине с юго-востока в Непале (известный как «стандартный путь»), а другой - с севера в Тибете.'
        },
        {
          text: 'Mount Everest is Earth\'s highest mountain above sea level located in the Mahalangur Himal',
          translatedText: 'Гора Эверест - самая высокая гора на Земле над уровнем моря расположенная в Махалангурских Гималах'
        },
        {
          text: 'As of 2019 over 300 people have died on Everest many of whose bodies remain on the mountain',
          translatedText: 'По состоянию на 2019 год на Эвересте погибло более 300 человек, тела многих из них остались на горе.'
        },
        {
          text: 'Mount Everest attracts many climbers, including highly experienced mountaineers',
          translatedText: 'Гора Эверест привлекает множество туристов, в том числе опытных альпинистов.'
        },
        {
          text: 'The China–Nepal border runs across its summit point.',
          translatedText: 'Граница между Китаем и Непалом проходит через его точку.'
        }
      ]
    }
  ],
  [
    {
      translates: [
        {
          text: 'testtest',
          translatedText: 'тест'
        }
      ]
    }
  ],
  [
    {
      translates: [
        {
          text: 'All together now',
          translatedText: 'Все вместе сейчас'
        }
      ]
    }
  ],
  [
    {
      translates: [
        {
          text: 'together',
          translatedText: 'вместе'
        }
      ]
    }
  ],
  [
    {
      translates: [
        {
          text: 'wuerzburg',
          translatedText: 'вюрцбург'
        },
        {
          text: 'Example',
          translatedText: 'Пример'
        }
      ]
    }
  ]
];

(async () => {
  try {
    // await testDocx('./meta327b.docx', trs[5], 'new');
    // await testDocx('./lyrics_punct.docx', trs[4], 'new');
    await testDocx('./Test_File.docx', trs[1], 'new');
    // await testDocx('./Test_File.docx', trs[1], 'old', false);
    console.log('finished')
  } catch (e) {
    console.error(e)
  }
})()
