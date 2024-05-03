'use strict';

const JSZip = require('jszip');
const fs = require('fs');
const xml2js = require('xml2js');
const parser = new xml2js.Parser();
const builder = new xml2js.Builder();

async function appendImageToDocx(docxPath, imageUrl) {
    const content = await fs.promises.readFile(docxPath);
    const zip = await JSZip.loadAsync(content);

    const relsPath = 'word/_rels/document.xml.rels';
    const docPath = 'word/document.xml';

    const [relsXml, docXml] = await Promise.all([
        zip.file(relsPath).async('string'),
        zip.file(docPath).async('string')
    ]);

    const [relsObj, docObj] = await Promise.all([
        parser.parseStringPromise(relsXml),
        parser.parseStringPromise(docXml)
    ]);

    const newId = `rId${relsObj.Relationships.Relationship.length + 1}`;
    const newRelationship = {
        $: {
            Id: newId,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            Target: imageUrl,
            TargetMode: 'External'
        }
    };
    relsObj.Relationships.Relationship.push(newRelationship);

    const newRelsXml = builder.buildObject(relsObj);
    zip.file(relsPath, newRelsXml);

    // Create a new paragraph with the provided drawing XML structure
    const drawing = {
        'w:p': [{
            'w:r': [{
                'w:drawing': [{
                    'wp:inline': [{
                        '$': {
                            'distT': '0',
                            'distB': '0',
                            'distL': '0',
                            'distR': '0'
                        },
                        'wp:extent': [{
                            '$': {
                                'cx': '0',
                                'cy': '0'
                            }
                        }],
                        'wp:effectExtent': [{
                            '$': {
                                'l': '0',
                                't': '0',
                                'r': '0',
                                'b': '0'
                            }
                        }],
                        'wp:docPr': [{
                            '$': {
                                'id': '1',
                                'name': '08f13379-b480-4b66-81d9-a4d3d9d45743'
                            }
                        }],
                        'wp:cNvGraphicFramePr': [{
                            'a:graphicFrameLocks': [{
                                '$': {
                                    'noChangeAspect': '1',
                                    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                                }
                            }]
                        }],
                        'a:graphic': [{
                            '$': {
                                'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                            },
                            'a:graphicData': [{
                                '$': {
                                    'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
                                },
                                'pic:pic': [{
                                    '$': {
                                        'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
                                    },
                                    'pic:nvPicPr': [{
                                        'pic:cNvPr': [{
                                            '$': {
                                                'id': '1',
                                                'name': '08f13379-b480-4b66-81d9-a4d3d9d45743'
                                            }
                                        }],
                                        'pic:cNvPicPr': [{}]
                                    }],
                                    'pic:blipFill': [{
                                        'a:blip': [{
                                            '$': {
                                                'r:link': newId,
                                                'cstate': 'print'
                                            },
                                            'a:extLst': [{
                                                'a:ext': [{
                                                    '$': {
                                                        'uri': '{28A0092B-C50C-407E-A947-70E740481C1C}'
                                                    }
                                                }]
                                            }]
                                        }],
                                        'a:stretch': [{
                                            'a:fillRect': [{}]
                                        }]
                                    }],
                                    'pic:spPr': [{
                                        'a:xfrm': [{
                                            'a:off': [{
                                                '$': {
                                                    'x': '0',
                                                    'y': '0'
                                                }
                                            }],
                                            'a:ext': [{
                                                '$': {
                                                    'cx': '0',
                                                    'cy': '0'
                                                }
                                            }]
                                        }],
                                        'a:prstGeom': [{
                                            '$': {
                                                'prst': 'rect'
                                            },
                                            'a:avLst': [{}]
                                        }]
                                    }]
                                }]
                            }]
                        }]
                    }]
                }]
            }]
        }]
    };

    if (!docObj['w:document']['w:body']) {
        docObj['w:document']['w:body'] = [];
    }

    docObj['w:document']['w:body'].push(drawing);

    const newDocXml = builder.buildObject(docObj);
    zip.file(docPath, newDocXml);

    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    await fs.promises.writeFile('./manual_result.docx', buffer);

    console.log('Image appended and DOCX file saved as output.docx');
}

appendImageToDocx('./test_docx.docx', 'http://localhost:5001/image.png').catch(console.error);