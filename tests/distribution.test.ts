import {describe,it,expect} from 'vitest';
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs';
import {join,resolve} from 'node:path';
import {tmpdir} from 'node:os';
import {execFileSync} from 'node:child_process';
import {parseStringPromise} from 'xml2js';
const script=resolve('scripts/manifest-production.mjs');
const dev=readFileSync('manifest.xml','utf8');
const prod=readFileSync('manifest.production.xml','utf8');
describe('개발형과 온라인 등록 파일 분리',()=>{
 it('별도 ID, 이름, 캐시 구분 주소를 생성하고 개발형은 보존한다',async()=>{
  const dir=mkdtempSync(join(tmpdir(),'ppt-manifest-'));
  try{
   writeFileSync(join(dir,'manifest.xml'),dev);writeFileSync(join(dir,'package.json'),JSON.stringify({version:'0.2.1'}));
   execFileSync(process.execPath,[script,'https://good-red.vercel.app'],{cwd:dir});
   const output=readFileSync(join(dir,'artifacts/manifest.production.xml'),'utf8');
   const parsed=await parseStringPromise(output);
   expect(parsed.OfficeApp.Id[0]).not.toBe((await parseStringPromise(dev)).OfficeApp.Id[0]);
   expect(parsed.OfficeApp.DisplayName[0].$.DefaultValue).toBe('PPT AI Studio 온라인');
   expect(parsed.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).toBe('https://good-red.vercel.app/index.html?addin=hosted&v=0.2.1');
   expect(output).not.toContain('localhost');expect(readFileSync(join(dir,'manifest.xml'),'utf8')).toBe(dev);
  }finally{rmSync(dir,{recursive:true,force:true});}
 });
 it('소스 ZIP의 개발 템플릿을 이용해 다른 배포 주소도 생성한다',()=>{
  const dir=mkdtempSync(join(tmpdir(),'ppt-manifest-'));
  try{
   writeFileSync(join(dir,'manifest.xml'),prod);writeFileSync(join(dir,'manifest.development.xml'),dev);writeFileSync(join(dir,'package.json'),JSON.stringify({version:'0.2.1'}));
   execFileSync(process.execPath,[script,'https://example.vercel.app'],{cwd:dir});
   const output=readFileSync(join(dir,'artifacts/manifest.production.xml'),'utf8');expect(output).toContain('https://example.vercel.app/index.html');expect(output).not.toContain('good-red.vercel.app');
  }finally{rmSync(dir,{recursive:true,force:true});}
 });
 it('배포 HTML 엔트리에 캐시 방지 헤더가 설정돼 있다',()=>{
  const config=JSON.parse(readFileSync('vercel.json','utf8'));
  for(const source of ['/','/index.html','/commands.html'])expect(config.headers.find((h:any)=>h.source===source).headers).toContainEqual({key:'Cache-Control',value:'no-store, max-age=0'});
 });
});
