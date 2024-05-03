const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs');
const path = require('path');

async function appendImageUrlToExcelWithDrawing(inputFile, outputFile, imageUrl) {
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

    // Assume we are adding to the first worksheet
    const drawingPath = 'xl/drawings/drawing1.xml';
    let drawingXml = await zip.file(drawingPath).async("string");
    const drawingObj = await parser.parseStringPromise(drawingXml);

    // Add the image reference to the drawing XML
    const drawingRel = {
        'xdr:pic': {
            'xdr:nvPicPr': {
                'xdr:cNvPr': {
                    '$': {
                        id: '2',
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
        }
    };

    if (!drawingObj['xdr:wsDr']) {
        drawingObj['xdr:wsDr'] = {};
    }
    drawingObj['xdr:wsDr']['xdr:pic'] = drawingRel;

    // Build the updated XML for drawing
    const newDrawingXml = builder.buildObject(drawingObj);
    zip.file(drawingPath, newDrawingXml);

    // Save the modified zip archive to a new XLSX file
    const content = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(outputFile, content);
    console.log('Modified XLSX file with image URL saved.');
}

// Usage
const inputFile = './test.xlsx';
const outputFile = './output.xlsx';
const imageUrl = 'http://localhost:5001/image.png';

appendImageUrlToExcelWithDrawing(inputFile, outputFile, imageUrl).catch(console.error);