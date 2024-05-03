'use strict';

const { readFileSync } = require('fs');
const http = require('http');
const server = http.createServer();

const png1x1 = readFileSync('./1x1.png');
server.on('request', (req, res) => {
    console.log('Start sendping png => ', req.url);
    // Set the content type to PNG
    res.writeHead(200, { 'Content-Type': 'image/png' });
    res.end(png1x1);
}).listen(5001);