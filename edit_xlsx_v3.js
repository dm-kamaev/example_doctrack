const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs');
const path = require('path');

async function appendImageUrlToExcelWithNewSheetRels(inputFile, outputFile, imageUrl) {
    // Read the XLSX file as a binary buffer
    const data = fs.readFileSync(inputFile);

    // Load the XLSX file as a zip archive
    const zip = await JSZip.loadAsync(data);

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
                                'r:embed': 'rId1' // This ID must match the ID in the relationships file
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
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();
    const newDrawingXml = builder.buildObject(drawingObj);
    zip.file('xl/drawings/drawing1.xml', newDrawingXml);

    // Create new relationships XML for the first worksheet
    const sheetRelsObj = {
        'Relationships': {
            '$': {
                'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
            },
            'Relationship': [{
                '$': {
                    'Id': 'rId1',
                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                    'Target': '../drawings/drawing1.xml'
                }
            }]
        }
    };

    // Build the new relationships XML
    const newSheetRelsXml = builder.buildObject(sheetRelsObj);
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', newSheetRelsXml);

    // Save the modified zip archive to a new XLSX file
    const content = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(outputFile, content);
    console.log('Modified XLSX file with new sheet relationships and drawing saved.');
}

// Usage
const inputFile = './test.xlsx';
const outputFile = './output.xlsx';
const imageUrl = 'http://localhost:5001/image.png';

appendImageUrlToExcelWithNewSheetRels(inputFile, outputFile, imageUrl).catch(console.error);