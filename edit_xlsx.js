const JSZip = require('jszip');
const xml2js = require('xml2js');
const fs = require('fs');

async function appendImageToXlsx(inputFilePath, outputFilePath, imageUrl) {
    const fileContent = fs.readFileSync(inputFilePath);
    const zip = await JSZip.loadAsync(fileContent);

    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const parser = new xml2js.Parser();
    const workbookObj = await parser.parseStringPromise(workbookXml);
    console.log(workbookObj);

    // const drawingFileName = `xl/drawings/drawing1.xml`;

    // let drawingObj;
    // if (!zip.file(drawingFileName)) {
    //     console.log('No existing drawing file found, creating one...');
    //     // Create a basic drawing XML structure
    //     drawingObj = {
    //         'xdr:wsDr': {
    //             '$': {
    //                 'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    //                 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    //                 'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    //             },
    //             'xdr:twoCellAnchor': [{
    //                 'xdr:from': {
    //                     'xdr:col': 0,
    //                     'xdr:colOff': 0,
    //                     'xdr:row': 0,
    //                     'xdr:rowOff': 0
    //                 },
    //                 'xdr:to': {
    //                     'xdr:col': 1,
    //                     'xdr:colOff': 0,
    //                     'xdr:row': 1,
    //                     'xdr:rowOff': 0
    //                 },
    //                 'xdr:pic': {},
    //                 'xdr:clientData': {}
    //             }]
    //         }
    //     };
    // } else {
    //     const drawingXml = await zip.file(drawingFileName).async('string');
    //     drawingObj = await parser.parseStringPromise(drawingXml);
    // }

    // // Update relationships file
    // const relsPath = `xl/drawings/_rels/drawing1.xml.rels`;
    // let relsXml = await zip.file(relsPath).async('string');
    // let relsObj = await parser.parseStringPromise(relsXml);

    // if (!relsObj.Relationships) {
    //     relsObj.Relationships = { 'Relationship': [] };
    // }

    // // Calculate new relationship ID
    // const newRelId = `rId${relsObj.Relationships.Relationship.length + 1}`;

    // // Define imageDetails object with dynamic r:link
    // const imageDetails = {
    //     'xdr:nvPicPr': {
    //         'xdr:cNvPr': {
    //             '$': {
    //                 id: '1',
    //                 name: 'image1'
    //             }
    //         },
    //         'xdr:cNvPicPr': {
    //             'a:picLocks': {
    //                 '$': {
    //                     noChangeAspect: '1'
    //                 }
    //             }
    //         }
    //     },
    //     'xdr:blipFill': {
    //         'a:blip': {
    //             '$': {
    //                 'r:link': newRelId  // Use the dynamically calculated relationship ID
    //             }
    //         },
    //         'a:stretch': {
    //             'a:fillRect': {}
    //         }
    //     },
    //     'xdr:spPr': {
    //         'a:xfrm': {
    //             'a:off': {
    //                 '$': {
    //                     x: '0',
    //                     y: '0'
    //                 }
    //             },
    //             'a:ext': {
    //                 '$': {
    //                     cx: '0',
    //                     cy: '0'
    //                 }
    //             }
    //         },
    //         'a:prstGeom': {
    //             'a:avLst': {}
    //         }
    //     }
    // };

    // // Insert imageDetails into the drawing object
    // if (!drawingObj['xdr:wsDr']['xdr:twoCellAnchor'][0]['xdr:pic']) {
    //     console.log('HERE');
    //     drawingObj['xdr:wsDr']['xdr:twoCellAnchor'][0]['xdr:pic'] = imageDetails;
    // }

    // const builder = new xml2js.Builder();
    // const updatedDrawingXml = builder.buildObject(drawingObj);
    // zip.file(drawingFileName, updatedDrawingXml);

    // // Add new relationship with dynamic ID
    // const newRel = {
    //     'Relationship': {
    //         '$': {
    //             'Id': newRelId,
    //             'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    //             'Target': imageUrl,
    //             'TargetMode': 'External'
    //         }
    //     }
    // };

    // relsObj.Relationships.Relationship.push(newRel);
    // const updatedRelsXml = builder.buildObject(relsObj);
    // zip.file(relsPath, updatedRelsXml);

    // const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    // fs.writeFileSync(outputFilePath, buffer);

    // console.log('Image appended and new xlsx file saved.');
}

appendImageToXlsx('./test.xlsx', './output.xlsx', 'http://localhost:5001/image.png');