'use strict';
let reportLibraryPromise;
function loadReportLibrary(){
 if(reportLibraryPromise)return reportLibraryPromise;
 reportLibraryPromise=(async()=>{
  if(!window.jspdf)await new Promise((resolve,reject)=>{const script=document.createElement('script');script.src='vendor/jspdf.umd.min.js';script.onload=resolve;script.onerror=()=>reject(Error('PDF library'));document.head.appendChild(script)});
  const response=await fetch('vendor/report-font.ttf');if(!response.ok)throw Error('PDF font');const bytes=new Uint8Array(await response.arrayBuffer());let raw='';for(let i=0;i<bytes.length;i+=8192)raw+=String.fromCharCode(...bytes.subarray(i,i+8192));return {jsPDF:window.jspdf.jsPDF,font:btoa(raw)};
 })().catch(e=>{reportLibraryPromise=null;throw e});return reportLibraryPromise;
}
async function downloadPDF(){
 const sections=Array.from(document.querySelectorAll('[name=reportSection]:checked')).map(x=>x.value);
 if(!sections.length){notify('En az bir bölüm seç.');return}
 const lang=$('#reportLang').value,vat=$('#reportVat').value,button=$('#downloadPdf');
 button.disabled=true;button.textContent=t('Rapor hazırlanıyor');
 try{
  const content=buildPrint(sections,lang,vat),{jsPDF,font}=await loadReportLibrary();
  const doc=new jsPDF({unit:'pt',format:'a4'});doc.addFileToVFS('report.ttf',font);doc.addFont('report.ttf','ESTA','normal');doc.setFont('ESTA');
  const parsed=new DOMParser().parseFromString(content,'text/html'),W=doc.internal.pageSize.getWidth(),H=doc.internal.pageSize.getHeight(),margin=40,available=W-margin*2;let y=90,page=1;
  const label=s=>lang==='ru'?(TRANSLATIONS_RU[s]||s):s;
  function header(){doc.setFillColor(23,27,34);doc.rect(0,0,W,64,'F');doc.setFontSize(14);doc.setTextColor(196,168,118);doc.text('ESTA CONSTRUCTION · '+DATA.meta.project,margin,29);doc.setFontSize(9);doc.setTextColor(243,245,247);doc.text(label('Finans ve saha raporu')+' · '+DATA.meta.date,margin,47);doc.setFontSize(8);doc.setTextColor(95,105,120);doc.text(label('Tutarlar milyon RUB (₽)'),margin,H-24);doc.text(String(page),W-margin,H-24,{align:'right'});y=90}
  function newPage(){doc.addPage();page++;header()}
  function ensure(height){if(y+height>H-48)newPage()}
  function paragraph(text,size=9){doc.setFontSize(size);const lines=doc.splitTextToSize(text,available);for(const line of lines){ensure(size*1.5);doc.setTextColor(80,91,108);doc.text(line,margin,y);y+=size*1.5}y+=5}
  function heading(text){doc.setFontSize(12);const lines=doc.splitTextToSize(text,available);ensure(lines.length*17+45);doc.setTextColor(50,58,72);doc.text(lines,margin,y);y+=lines.length*17;doc.setDrawColor(196,168,118);doc.line(margin,y-5,W-margin,y-5);y+=9}
  function table(el){const rows=Array.from(el.querySelectorAll('tr')).map(tr=>Array.from(tr.children).map(td=>td.textContent.trim()));if(!rows.length)return;const n=rows[0].length;const widths=n===2?[available*.64,available*.36]:n===3?[available*.48,available*.3,available*.22]:n===4?[available*.3,available*.22,available*.16,available*.32]:[available*.3,...Array(n-1).fill(available*.7/(n-1))];
   function row(cells,isHead){doc.setFontSize(isHead?8:8.5);const lines=cells.map((c,i)=>doc.splitTextToSize(c,widths[i]-12));const count=Math.max(...lines.map(l=>l.length));const height=count*12+14;if(y+height>H-48){newPage();if(!isHead)row(rows[0],true)}if(isHead){doc.setFillColor(237,240,244);doc.rect(margin,y-3,available,height,'F')}let x=margin;lines.forEach((ls,i)=>{doc.setTextColor(isHead?90:30,isHead?100:40,isHead?115:55);doc.text(ls,x+6,y+10);x+=widths[i]});y+=height;doc.setDrawColor(220,225,232);doc.line(margin,y-3,W-margin,y-3)}
   rows.forEach((r,i)=>row(r,i===0));y+=14;
  }
  header();paragraph(label('Finansal tutarlar rapor tarihine aittir. Grafikler kendi kayıt dönemini gösterir.'));
  Array.from(parsed.querySelectorAll('.print-section')).forEach((section,index)=>{if(index)newPage();for(const child of section.children){if(child.tagName==='H2')heading(child.textContent);else if(child.tagName==='TABLE')table(child);else if(child.tagName==='P')paragraph(child.textContent)}});
  doc.save(`ESTA-${DATA.meta.project}-${DATA.meta.date}-${lang.toUpperCase()}.pdf`);closeDialog();
 }catch(error){notify('PDF hazırlanamadı; PDF / Yazdır seçeneğini kullan.');button.disabled=false;button.textContent=t('PDF indir')}
}
