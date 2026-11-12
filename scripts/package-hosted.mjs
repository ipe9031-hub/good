import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { zipSync, strToU8 } from 'fflate';
const manifest = readFileSync('manifest.production.xml');
if (manifest.toString().includes('localhost') || !manifest.toString().includes('https://good-red.vercel.app')) throw new Error('Hosted manifest origin is incorrect');
const data = zipSync({
  'PPT-AI-Studio-Hosted/manifest.xml': new Uint8Array(manifest),
  'PPT-AI-Studio-Hosted/설치안내.md': new Uint8Array(readFileSync('docs/INSTALL-HOSTED.md')),
  'PPT-AI-Studio-Hosted/README.txt': strToU8('PPT AI Studio - Hosted Windows add-in\r\nServer: https://good-red.vercel.app/\r\nNo Node.js or localhost certificate installation required.\r\nRegister manifest.xml in PowerPoint. See 설치안내.md.\r\nWindows PowerPoint host validation: pending.\r\n')
}, {level:6});
mkdirSync('downloads', {recursive:true});
const path = 'downloads/PPT-AI-Studio-Hosted-Windows.zip';
writeFileSync(path, data);
writeFileSync(path+'.sha256', createHash('sha256').update(data).digest('hex')+'  PPT-AI-Studio-Hosted-Windows.zip\n');
console.log('Created',path,data.length,'bytes');

const version=JSON.parse(readFileSync('package.json','utf8')).version;
const versionedName=`PPT-AI-Studio-Online-${version}.zip`;
writeFileSync('downloads/'+versionedName,data);
writeFileSync('downloads/'+versionedName+'.sha256',createHash('sha256').update(data).digest('hex')+'  '+versionedName+'\n');
console.log('Created versioned online package:', versionedName);
