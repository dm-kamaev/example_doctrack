'use strict';

const { randomUUID } = require('node:crypto');
const fs = require('node:fs/promises');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const inputFile = './empty.xlsx';
// const inputFile = './test_with_image.xlsx';
// const inputFile = './manual_result.xlsx';
const outputFile = './manual_result.xlsx';
const imageUrl = 'http://localhost:5001/image.png';
// const imageUrl = 'http://localhost:5001/image2.png';

appendImageUrlToExcelWithNewSheetRels(inputFile, outputFile, imageUrl).catch(console.error);

async function appendImageUrlToExcelWithNewSheetRels(inputPath, outputPath, imageUrl) {
  const data = await fs.readFile(inputPath);

  const zip = await JSZip.loadAsync(data);

  const parser = new xml2js.Parser();
  const builder = new xml2js.Builder();

  const hasNotDrawingForSheet1 = !Boolean(zip.file('xl/drawings/drawing1.xml'));
  console.log(zip.file('xl/drawings/drawing1.xml'));

  const { rId: imageUrlRId, filePath: drawingPath } = await appendBlankDrawing({ zip, parser, builder });
  await appendRelationshipBetweenDrawingAndImage({ zip, parser, builder, drawingPath, image: { rId: imageUrlRId, url: imageUrl } });

  // const hasNotDrawingForSheet1 = true;
  console.log({ hasNotDrawingForSheet1 });
  if (hasNotDrawingForSheet1) {
    const drawingRId = await appendDrawingsOnSheet({ zip, parser, builder });
    await appendRelationshipBetweenDrawingsAndSheet({ zip, parser, builder, drawing: { rId: drawingRId, path: drawingPath } });
  }

  const content = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(outputPath, content);

  console.log(`SUCCESS: outputPath ===> ${outputPath}`);
}

function generateRId() {
  return `R${randomUUID().replaceAll('-', '')}`;
}

async function appendBlankDrawing({ zip, parser, builder }) {
  const filePath = 'xl/drawings/drawing1.xml';
  let relsXml;

  if (!zip.file(filePath)) {
    relsXml = `<?xml version="1.0" encoding="utf-8"?>
    <xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    </xdr:wsDr>`;
  } else {
    relsXml = await zip.file(filePath).async('string');
  }

  const drawingResult = await parser.parseStringPromise(relsXml);

  const rId = generateRId();

   // New blank draw
   const newDrawing = {
      '$': { editAs: 'oneCell' },
      'xdr:from': {
          'xdr:col': '0',
          'xdr:colOff': '0',
          'xdr:row': '0',
          'xdr:rowOff': '0'
      },
      'xdr:to': {
          'xdr:col': '3',
          'xdr:colOff': '0',
          'xdr:row': '3',
          'xdr:rowOff': '0'
      },
      'xdr:pic': {
          'xdr:nvPicPr': {
              'xdr:cNvPr': {
                  '$': {
                      id: '1',
                      name: 'Picture 1'
                  }
              },
              'xdr:cNvPicPr': {}
          },
          'xdr:blipFill': {
              'a:blip': {
                  '$': {
                      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                      'r:link': rId // This ID must match the ID in the relationships file
                  }
              },
              'a:stretch': {
                  'a:fillRect': {}
              }
          },
          'xdr:spPr': {
              'a:xfrm': {
                  'a:off': { '$': { x: '0', y: '0' } },
                  'a:ext': { '$': { cx: '1000000', cy: '1000000' } }
              },
              'a:prstGeom': {
                  '$': { prst: 'rect' },
                  'a:avLst': {}
              }
          }
      },
      'xdr:clientData': {}
    };

    if (!drawingResult['xdr:wsDr']['xdr:twoCellAnchor']) {
      drawingResult['xdr:wsDr']['xdr:twoCellAnchor'] = [];
    }
    drawingResult['xdr:wsDr']['xdr:twoCellAnchor'].push(newDrawing);


  zip.file(filePath, builder.buildObject(drawingResult));

  return { rId, filePath };
}

async function appendRelationshipBetweenDrawingAndImage({ zip, parser, builder, image }) {
  const filePath = `xl/drawings/_rels/drawing1.xml.rels`;
  let relsXml;
  if (!zip.file(filePath)) {
    relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  } else {
    relsXml = await zip.file(filePath).async('string');
  }

  const relsResult = await parser.parseStringPromise(relsXml);

    // Append new relationship for the external image url
    const newRel = {
      $: {
          Id: image.rId,
          Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          Target: image.url,
          TargetMode: 'External'
      }
    };

    if (!relsResult.Relationships) {
      relsResult.Relationships = {
        $: {
           'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
        },
        Relationship: []
      };
    }

    if (!relsResult.Relationships.Relationship) {
        relsResult.Relationships.Relationship = [];
    }

    relsResult.Relationships.Relationship.push(newRel);

    zip.file(filePath, builder.buildObject(relsResult));
}

async function appendDrawingsOnSheet({ zip, parser, builder }) {
  const rId = generateRId();
  const filePath = 'xl/worksheets/sheet1.xml';

  // Check if the sheet file exists
  if (!zip.file(filePath)) {
    throw new Error('Sheet1.xml does not exist in the provided XLSX file.');
  }


  const sheetXml = await zip.file(filePath).async('string');
  const sheetResult = await parser.parseStringPromise(sheetXml);

  // Append new drawing elements on sheet
  const newDrawing = {
    $: {
      'r:id': rId
     }
  };

  if (!sheetResult.worksheet.drawing) {
    sheetResult.worksheet.drawing = [];
  }
  sheetResult.worksheet.drawing.push(newDrawing);

  zip.file(filePath, builder.buildObject(sheetResult));

  return rId;
}


async function appendRelationshipBetweenDrawingsAndSheet({ zip, parser, builder, drawing }) {
  const filePath = 'xl/worksheets/_rels/sheet1.xml.rels';

  let relsXml;
  if (!zip.file(filePath)) {
    relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  } else {
    relsXml = await zip.file(filePath).async('string');
  }


  const relsResult = await parser.parseStringPromise(relsXml);


  // Create relationship for drawing on sheet
  const newRel = {
    $: {
       Id: drawing.rId,
       Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
       Target: '/'+drawing.path
    }
  };

  if (!relsResult.Relationships) {
    relsResult.Relationships = {
      Relationships: {
          $: {
              'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
          },
          Relationship: []
      }
    };
  }

  if (!relsResult.Relationships.Relationship) {
    relsResult.Relationships.Relationship = [];
  }

  relsResult.Relationships.Relationship.push(newRel);

  zip.file(filePath, builder.buildObject(relsResult));
}


