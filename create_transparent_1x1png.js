const sharp = require('sharp');

// Create a 1x1 transparent PNG
const transparentPixel = Buffer.from([0, 0, 0, 0]); // RGBA values

sharp({
  create: {
    width: 1,
    height: 1,
    channels: 4,
    background: { r: 0, g: 0, b: 0, alpha: 0 }
  }
})
.png()
.toBuffer()
.then(data => {
  require('fs').writeFileSync('1x1.png', data);
  console.log('Transparent 1x1 PNG file created successfully!');
})
.catch(err => {
  console.error('Error creating PNG file:', err);
});