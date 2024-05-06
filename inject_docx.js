// Inject tracking pixel url in document. Support formats: .docx, .docm, .dotx. Node js realization of c# realization https://github.com/wavvs/doctrack

'use strict';

const { randomUUID } = require('node:crypto');
const fs = require('node:fs/promises');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const inputFile = './test.docx';
// const inputFile = './test_dotx.dotx';
// const inputFile = './empty.docx';
// const inputFile = './test_with_image.docx';
// const inputFile = './manual_result.docx';
// const inputFile = './output.docx';
// const outputFile = './manual_result.dotx';
const outputFile = './output.docx';
const imageUrl = 'http://localhost:5001/image.png';
// const imageUrl = 'http://localhost:5001/image2.png';

appendImageToDocx(inputFile, outputFile, imageUrl).catch(console.error);

async function appendImageToDocx(inputPath, outputPath, imageUrl) {
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();

    const content = await fs.readFile(inputPath);
    const zip = await JSZip.loadAsync(content);

    const relsPath = 'word/_rels/document.xml.rels';
    const docPath = 'word/document.xml';

    const [relsXml, docXml] = await Promise.all([
        zip.file(relsPath).async('string'),
        zip.file(docPath).async('string'),
    ]);

    const [relsObj, docObj] = await Promise.all([
        parser.parseStringPromise(relsXml),
        parser.parseStringPromise(docXml),
    ]);

    const lastId = relsObj.Relationships.Relationship.length;
    const newRId = `rId${lastId+1}`;
    const newRelationship = {
        $: {
            Id: newRId,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            Target: imageUrl,
            TargetMode: 'External'
        }
    };
    relsObj.Relationships.Relationship.push(newRelationship);

    const newRelsXml = builder.buildObject(relsObj);
    zip.file(relsPath, newRelsXml);

    // console.dir(docObj['w:document']['w:body'], { depth: 3 });

    const pictureName = randomUUID(); // 08f13379-b480-4b66-81d9-a4d3d9d45743

    // New paragraph with blank draw
      const drawing = {
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
                                'name': pictureName
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
                                                'name': pictureName
                                            }
                                        }],
                                        'pic:cNvPicPr': [{}]
                                    }],
                                    'pic:blipFill': [{
                                        'a:blip': [{
                                            '$': {
                                                'r:link': newRId,
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
    };


    if (!docObj['w:document']['w:body']) {
        console.log('Initiazation!');
        docObj['w:document']['w:body'] = [{ 'w:p': [] }];
    }

    const wBody = docObj['w:document']['w:body'];
    wBody[0]['w:p'].push(drawing);

    // console.dir(docObj['w:document']['w:body'], { depth: 3 });

    zip.file(docPath, builder.buildObject(docObj));


    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    await fs.writeFile(outputPath, buffer);

    console.log(`SUCCESS: outputPath ===> ${outputPath}`);
}

