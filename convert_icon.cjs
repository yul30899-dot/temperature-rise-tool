const Jimp = require('jimp');
const pngToIco = require('png-to-ico');
const fs = require('fs');

async function convert() {
  console.log('Reading image...');
  const image = await Jimp.read('build/icon.png');
  console.log('Resizing to 256x256...');
  image.resize(256, 256);
  console.log('Saving as true PNG...');
  await image.writeAsync('build/icon_true.png');
  
  console.log('Converting to ICO...');
  const buf = await pngToIco('build/icon_true.png');
  fs.writeFileSync('build/icon.ico', buf);
  console.log('Done!');
}

convert().catch(console.error);
