const { randomUUID } = require('node:crypto');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs');
const path = require('path');

async function appendImageUrlToExcelWithNewSheetRels(inputFile, outputFile, imageUrl) {
  // Read the XLSX file as a binary buffer
  const data = fs.readFileSync(inputFile);

  // Load the XLSX file as a zip archive
  const zip = await JSZip.loadAsync(data);


  // Build the new drawing XML
  const parser = new xml2js.Parser();
  const builder = new xml2js.Builder();

  const { rId: rIdImageUrl, filePath } = await appendDrawingToImage(zip);
  await appendRelationshipToImage(zip, { rId: rIdImageUrl, url: imageUrl });

  const rIdDrawing = await appendDrawingOnSheet(zip);
  await appendRelationshipToDrawingOnSheet(zip, { rId: rIdDrawing, path: filePath });


  // // Save the modified zip archive to a new XLSX file
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(outputFile, content);
  console.log('Modified XLSX file with new sheet relationships and drawing saved.');
}

function generateRId() {
  return `R${randomUUID().replaceAll('-', '')}`;
}

function appendDrawingToImage(zip) {
  const rId = generateRId();

  const builder = new xml2js.Builder();
   // Create a new drawing XML structure
    const drawingObj = {
        'xdr:wsDr': {
            '$': {
                'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            },
            'xdr:twoCellAnchor': {
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
            }
        }
    };

  // TODO: We should resolving conflict of filepath
  const filePath = 'xl/drawings/drawing1.xml'
  zip.file(filePath, builder.buildObject(drawingObj));

  return { rId, filePath };
}

async function appendRelationshipToImage(zip, image) {
    // TODO: We should resolving conflict of filepath
    const filePath = 'xl/drawings/_rels/drawing1.xml.rels';
    let relsXml;
    if (!zip.file(filePath)) {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    } else {
      relsXml = await zip.file(filePath).async('string');
    }

    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();
    const relsResult = await parser.parseStringPromise(relsXml);

    // Append new relationship for the external image
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

    const updatedRelsXml = builder.buildObject(relsResult);
    zip.file(filePath, updatedRelsXml);
}

async function appendDrawingOnSheet(zip) {
  const rId = generateRId();
  // TODO: We should resolving conflict of filepath
  // TODO: We should check renaming sheets
  const filePath = 'xl/worksheets/sheet1.xml';
  const parser = new xml2js.Parser();
  const builder = new xml2js.Builder();

  // Check if the sheet file exists
  if (!zip.file(filePath)) {
    throw new Error('Sheet1.xml does not exist in the provided XLSX file.');
  }

  const sheetXml = await zip.file(filePath).async('string');
  const sheetResult = await parser.parseStringPromise(sheetXml);

  // Append new drawing element
  const newDrawing = {
    $: {
      'r:id': rId
     }
  };

  // Ensure the worksheet has a drawing element array to push to
  if (!sheetResult.worksheet.drawing) {
    sheetResult.worksheet.drawing = [];
  }
  sheetResult.worksheet.drawing.push(newDrawing);

  zip.file(filePath, builder.buildObject(sheetResult));

  return rId;
}


async function appendRelationshipToDrawingOnSheet(zip, drawing) {
  // TODO: We should resolving conflict of filepath
  const filePath = 'xl/worksheets/_rels/sheet1.xml.rels';

  let relsXml;
  if (!zip.file(filePath)) {
    relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  } else {
    relsXml = await zip.file(filePath).async('string');
  }

  const parser = new xml2js.Parser();
  const builder = new xml2js.Builder();
  const relsResult = await parser.parseStringPromise(relsXml);


  // Create new relationships XML for the first worksheet
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

// Usage
const inputFile = './test.xlsx';
const outputFile = './manual_result.xlsx';
const imageUrl = 'http://localhost:5001/image.png';

appendImageUrlToExcelWithNewSheetRels(inputFile, outputFile, imageUrl).catch(console.error);