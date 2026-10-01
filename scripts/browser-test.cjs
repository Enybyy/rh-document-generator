/* Run against local Python on 5081 and the static demo on 5085. */
const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const fs=require('node:fs');
const path=require('node:path');
const assert=require('node:assert/strict');
(async()=>{
 const root=path.resolve(__dirname,'..');fs.mkdirSync(path.join(root,'assets/screenshots'),{recursive:true});fs.mkdirSync(path.join(root,'_qa'),{recursive:true});
 const browser=await chromium.launch({headless:true});const page=await browser.newPage({viewport:{width:1440,height:1080}});const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('http://127.0.0.1:5085/');await page.waitForSelector('#rows .person');assert.equal(await page.locator('#rows tr').count(),3);
 await page.screenshot({path:path.join(root,'assets/screenshots/rh-desktop.png'),fullPage:true});await page.screenshot({path:path.join(root,'assets/screenshots/rh-upwork.png')});
 await page.locator('#edit-name').fill('Prueba á & < >');await page.locator('#save-edit').click();assert.match(await page.locator('#preview').innerText(),/Prueba á & < >/);
 await page.locator('#select-all').uncheck();assert.equal(await page.locator('#generate').isDisabled(),true);await page.locator('#select-all').check();
 const download=page.waitForEvent('download');await page.locator('#generate').click();await (await download).saveAs(path.join(root,'_qa/demo.zip'));
 await page.locator('#csv').setInputFiles({name:'personal.csv',mimeType:'text/csv',buffer:Buffer.from('DNI,NOMBRE,CARGO,CIUDAD,FECHA_INICIO,FECHA_FIN,PAGO\n00000011,"Demo, María",Analista,Lima,01/10/2026,31/12/2026,3000\nmal,Nombre,Cargo,Lima,31/02/2026,31/12/2026,-1')});await page.locator('#load').click();await page.waitForFunction(()=>document.querySelector('#status').textContent.includes('1 registros listos'));assert.equal(await page.locator('#rows input:disabled').count(),1);
 await page.setViewportSize({width:360,height:800});await page.screenshot({path:path.join(root,'assets/screenshots/rh-mobile.png'),fullPage:true});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
 await page.setViewportSize({width:1440,height:1080});await page.goto('http://127.0.0.1:5081/');await page.waitForFunction(()=>document.querySelector('#mode').textContent.includes('Aplicación local'));
 const xlsx=await (await page.request.get('http://127.0.0.1:5081/api/samples/personal.xlsx')).body();const docx=await (await page.request.get('http://127.0.0.1:5081/api/samples/plantilla.docx')).body();
 await page.locator('#excel').setInputFiles({name:'personal.xlsx',mimeType:'application/octet-stream',buffer:xlsx});await page.locator('#template').setInputFiles({name:'plantilla.docx',mimeType:'application/octet-stream',buffer:docx});await page.locator('#load').click();await page.waitForFunction(()=>document.querySelector('#source').textContent.includes('Archivo cargado'));assert.equal(await page.locator('#rows tr').count(),3);
 await page.locator('#excel').setInputFiles({name:'changed.xlsx',mimeType:'application/octet-stream',buffer:xlsx});assert.equal(await page.locator('#generate').isDisabled(),true);await page.locator('#load').click();await page.waitForFunction(()=>!document.querySelector('#generate').disabled);const localDownload=page.waitForEvent('download');await page.locator('#generate').click();await (await localDownload).saveAs(path.join(root,'_qa/local.zip'));assert.deepEqual(errors,[]);
 console.log('PASS RH browser: edit, selection, real ZIP, CSV errors, 360px, XLSX+DOCX local upload and ZIP.');await browser.close();
})().catch(e=>{console.error(e);process.exit(1);});
