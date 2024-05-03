'use strict';

const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

async function appendHyperlinkToDocx(filePath, newFilePath, hyperlink) {
    // Read the DOCX file
    const content = await fs.promises.readFile(filePath);
    const zip = new JSZip();
    // Load zip content
    await zip.loadAsync(content);

    // Get the document.xml and document.xml.rels files
    const xmlContent = await zip.file('word/document.xml').async('string');
    const relsContent = await zip.file('word/_rels/document.xml.rels').async('string');

    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();

    // Parse the XML content
    const doc = await parser.parseStringPromise(xmlContent);
    const rels = await parser.parseStringPromise(relsContent);

    // Append a hyperlink to the body
    const body = doc['w:document']['w:body'][0];
    const relId = 'rId100'; // Example new relationship ID
    const hyperlinkXml = {
        'w:hyperlink': [{
            'w:r': [{
                'w:rPr': [{}],
                'w:t': [{ '_': hyperlink.text }]
            }],
            '$': { 'r:id': relId, 'w:history': '1' }
        }]
    };

    // Add the hyperlink to the body
    body['w:p'] = (body['w:p'] || []).concat(hyperlinkXml);

    // Add or update the relationship in document.xml.rels
    rels['Relationships']['Relationship'] = (rels['Relationships']['Relationship'] || []).concat({
        '$': {
            'Id': relId,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'Target': hyperlink.url,
            'TargetMode': 'External'
        }
    });

    // Build the new XML from the JS object
    const newXmlContent = builder.buildObject(doc);
    const newRelsContent = builder.buildObject(rels);

    // Replace the old XML with the new one in the ZIP file
    zip.file('word/document.xml', newXmlContent);
    zip.file('word/_rels/document.xml.rels', newRelsContent);

    // Save the new DOCX file
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    await fs.promises.writeFile(newFilePath, buffer);
}

// Example usage
const hyperlink = { text: 'Click here', url: 'http://localhost:5001/image.png' };
appendHyperlinkToDocx('./test_docx.docx', './output/test_docx.docx', hyperlink)
    .then(() => console.log('Hyperlink added successfully!'))
    .catch(err => console.error('Error:', err));

// appendHyperlinkToDocx('./test_docx.docx', './output.docx', hyperlink)
//     .then(() => console.log('Hyperlink added successfully!'))
//     .catch(err => console.error('Error:', err));