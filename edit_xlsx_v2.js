const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs');
const path = require('path');

async function appendImageUrlToExcelWithNewDrawing(inputFile, outputFile, imageUrl) {
    // Read the XLSX file as a binary buffer
    const data = fs.readFileSync(inputFile);

    // Load the XLSX file as a zip archive
    const zip = await JSZip.loadAsync(data);

    // Get the workbook relationships file
    const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
    const workbookRelsXml = await zip.file(workbookRelsPath).async("string");

    // Parse the XML
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();
    const workbookRelsObj = await parser.parseStringPromise(workbookRelsXml);

    // Add a new relationship for the image
    const imageRelId = 'rId' + (workbookRelsObj.Relationships.Relationship.length + 1);
    const imageRel = {
        '$': {
            Id: imageRelId,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            Target: imageUrl,
            TargetMode: 'External'
        }
    };
    workbookRelsObj.Relationships.Relationship.push(imageRel);

    // Build the updated XML for workbook relationships
    const newWorkbookRelsXml = builder.buildObject(workbookRelsObj);
    zip.file(workbookRelsPath, newWorkbookRelsXml);

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
                                'r:embed': imageRelId
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

    // Build the new drawing XML
    const newDrawingXml = builder.buildObject(drawingObj);
    zip.file('xl/drawings/drawing1.xml', newDrawingXml);

    // Update the first worksheet's relationships to include the new drawing
    const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
    const sheetRelsXml = await zip.file(sheetRelsPath).async("string");
    const sheetRelsObj = await parser.parseStringPromise(sheetRelsXml);

    const drawingRel = {
        '$': {
            Id: 'rId' + (sheetRelsObj.Relationships.Relationship.length + 1),
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
            Target: '../drawings/drawing1.xml'
        }
    };
    sheetRelsObj.Relationships.Relationship.push(drawingRel);

    const newSheetRelsXml = builder.buildObject(sheetRelsObj);
    zip.file(sheetRelsPath, newSheetRelsXml);

    // Save the modified zip archive to a new XLSX file
    const content = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(outputFile, content);
    console.log('Modified XLSX file with new drawing and image URL saved.');
}

// Usage
const inputFile = './test.xlsx';
const outputFile = './output.xlsx';
const imageUrl = 'http://localhost:5001/image.png';

appendImageUrlToExcelWithNewDrawing(inputFile, outputFile, imageUrl).catch(console.error);