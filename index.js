import * as fs from 'fs';
import { promisify } from 'util';
import { exec } from 'child_process';
import zipdir from 'zip-dir';
import { parseStringPromise, Builder } from 'xml2js';
import { find } from 'xml2js-xpath';
import { convert } from 'libreoffice-convert';
import { decode } from 'html-entities';

import mkdirp from 'mkdirp';

const docx2pdf = promisify(convert);
const execP = promisify(exec);
const fsP = fs.promises;

async function processDocx(fileType, id, data) {
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

  for (const {from, to, translates} of data.Text) {
    await execP(`unzip ${pathDocx} -d ${dir}to`);
    const path = `${dir}${to}/word/document.xml`;
    const xmlData = await fsP.readFile(path);
    const doc = await parseStringPromise(xmlData.toString());
    const matches = find(doc, "//w:r");

    [...translates].sort((x, y) => (y?.text?.length || 0) - (x?.text?.length || 0)).forEach(({text, translatedText}) => {
      if (!translatedText) return;

      let joinedText = '';
      const borders = [];
      let left = 0;

      matches.forEach((run, j) => {
        if (!run['w:t']) return;

        run['w:t'].forEach((t, k) => {
          let part = '';
          if (typeof t === 'string') part = t;
          if (typeof t['_'] === 'string') part = t['_'];
          if (!part) return;
          joinedText += part;
          const right = left + part.length;
          borders.push([j, k, left, right]);
          left = right;
        });
      });

      const foundL = joinedText.indexOf(text);
      if (foundL === -1) return;
      const foundR = foundL + text.length;

      for (const [j, k, l, r] of borders) {
        if (r <= foundL || l >= foundR) continue;

        if (joinedText.slice(l, r).includes(text)) {
          if (typeof matches[j]['w:t'][k] === 'string') {
            matches[j]['w:t'][k] = matches[j]['w:t'][k].replace(text, translatedText)
          } else {
            matches[j]['w:t'][k]['_'] = matches[j]['w:t'][k]['_'].replace(text, translatedText)
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
    });

    const builder = new Builder();
    const updatedXML = builder.buildObject(doc);
    await fsP.writeFile(path, updatedXML);
    const buffer = await zipdir(`${dir}${to}`);
    const pathTo = `${dir}${to}.docx`;
    await fsP.writeFile(pathTo, buffer);

    if (fileType === 'pdf') {
      const bufferPdf = await docx2pdf(buffer, 'pdf', undefined);
      await fsP.writeFile(`${dir}${to}.pdf`, bufferPdf);
      resps.push({from, to, buffer: bufferPdf})
    } else {
      resps.push({from, to, buffer})
    }
  }

  return resps
}

async function testDocx(path) {
  const dataDocx = await fsP.readFile(path);
  const resultDocx = await processDocx('docx', 1, {
    buffer: dataDocx,
    Text: [
      // {
      //   translates: [
      //
      //   ]
      // },
      {
        // from: 'en',
        // to: 'ru',
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
      },
      // {
      //   // from: 'en',
      //   // to: 'ru',
      //   translates: [
      //     {
      //       text: 'While not posing substantial technical climbing challenges on the standard route, Everest presents dangers such as altitude sickness weather and wind as well as significant hazards from avalanches and the Khumbu Icefal',
      //       translatedText: 'Не создавая серьезных технических проблем для поднятия на стандартном пути, Эверест несет в себе такие опасности, как высотная болезнь, погодные условия'
      //     },
      //     {
      //       text: 'There are two main climbing routes one approaching the summit from the southeast in Nepal known as the standard route and the other from the north in Tibet',
      //       translatedText: 'Есть два основных маршрута для восхождения: один приближается к вершине с юго-востока в Непале (известный как «стандартный путь»), а другой - с севера в Тибете.'
      //     },
      //     {
      //       text: 'Mount Everest is Earth\'s highest mountain above sea level located in the Mahalangur Himal',
      //       translatedText: 'Гора Эверест - самая высокая гора на Земле над уровнем моря расположенная в Махалангурских Гималах'
      //     },
      //     {
      //       text: 'As of 2019 over 300 people have died on Everest many of whose bodies remain on the mountain',
      //       translatedText: 'По состоянию на 2019 год на Эвересте погибло более 300 человек, тела многих из них остались на горе.'
      //     },
      //     {
      //       text: 'Mount Everest attracts many climbers, including highly experienced mountaineers',
      //       translatedText: 'Гора Эверест привлекает множество туристов, в том числе опытных альпинистов.'
      //     },
      //     {
      //       text: 'The China–Nepal border runs across its summit point.',
      //       translatedText: 'Граница между Китаем и Непалом проходит через его точку.'
      //     }
      //   ]
      // }
    ]
  });

  await Promise.all(resultDocx.map(({buffer}, i) => fsP.writeFile(`result${i}.docx`, buffer)));
}

async function testPdf(path) {
  const dataPdf = await fsP.readFile(path);
  const resultPdf = await processDocx('pdf', 2, {
    buffer: dataPdf,
    Text: [
      {
        from: 'en',
        to: 'ru',
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
            text: 'Mount Everest is Earthaposs highest mountain above sea level located in the Mahalangur Himal',
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
    ]
  });
  await Promise.all(resultPdf.map(({buffer}, i) => fsP.writeFile(`result${i}.pdf`, buffer)));
}

async function main() {
  try {
    // await testDocx('./english-test.docx');
    await testDocx('./Test_File.docx');
    await testPdf('./Test_File.pdf');

    // await testDocx('./article_about_giraffes_test_file.docx');
    // await testDocx('./CAT_article_with_two_columns_test-file.docx');
    // await testPdf('./Article_about_elefants_test_file.pdf');
    // await testPdf('./CAT_article_with_two_columns_test-file.pdf');
    console.log('finished')
  } catch (e) {
    console.error(e)
  }
}

main()
