'use strict';

const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

async function appendExternalImageUrl(filePath, newFilePath, imageUrl) {
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

    // Define a new relationship ID for the hyperlink
    const hyperlinkRelId = 'rIdHyperlinkImage'; // Ensure this ID is unique in the document

    // Define the hyperlink XML structure
    const hyperlinkXml = {
        'w:hyperlink': [{
            'w:r': [{
                'w:rPr': [{}],
                'w:t': [{ '_': 'Click here to view the image' }]
            }],
            '$': { 'r:id': hyperlinkRelId, 'w:history': '1' }
        }]
    };

    // Append the hyperlink to the document body
    const body = doc['w:document']['w:body'][0];
    body['w:p'] = (body['w:p'] || []).concat(hyperlinkXml);

    // Add the hyperlink relationship to document.xml.rels
    rels['Relationships']['Relationship'].push({
        '$': {
            'Id': hyperlinkRelId,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'Target': imageUrl,
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
const imageUrl = 'http://localhost:5001/image.png';
appendExternalImageUrl('./test_docx.docx', './output/test_docx.docx', imageUrl)
    .then(() => console.log('External image URL appended successfully!'))
    .catch(err => console.error('Error:', err))