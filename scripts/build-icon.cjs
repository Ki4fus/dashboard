const fs = require('node:fs/promises');
const sharp = require('sharp');
(async () => {
  const chunks = [];
  for (const [type, size] of [['icp4',16],['icp5',32],['icp6',64],['ic07',128],['ic08',256],['ic09',512],['ic10',1024],['ic11',32],['ic12',64],['ic13',256],['ic14',512]]) {
    const png = await sharp('public/netting-logo.svg').resize(size,size).png().toBuffer();
    const header = Buffer.alloc(8); header.write(type); header.writeUInt32BE(png.length+8,4);
    chunks.push(header,png);
  }
  const header = Buffer.alloc(8);header.write('icns');header.writeUInt32BE(8+chunks.reduce((sum,b)=>sum+b.length,0),4);
  await fs.mkdir('build',{recursive:true});await fs.writeFile('build/netting.icns',Buffer.concat([header,...chunks]));
})();
