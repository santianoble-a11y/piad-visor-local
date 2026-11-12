'use strict';
(() => {
 const $ = id => document.getElementById(id);
 const norm = v => String(v == null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]+/g,' ').trim();
 const display = v => v == null ? '' : v instanceof Date ? v.toLocaleDateString('es-EC',{timeZone:'UTC'}) : String(v).trim();
 const sheet = {records:[], filtered:[], page:0, pdfURL:null};
 const pageSize=25;
 const column = (headers,...labels) => { for (const label of labels) {const idx=headers.indexOf(norm(label));if(idx>=0)return idx;} return -1; };
 function resetOptions(id, values){const el=$(id), old=el.value;el.replaceChildren(new Option('Todos',''));[...new Set(values.filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es')).forEach(x=>el.add(new Option(x,x)));if(values.includes(old))el.value=old;}
 function parseFile(buf){
  const wb=XLSX.read(buf,{type:'array',cellDates:true,sheets:['OPERATIVOS'],dense:true});
  const sn=wb.SheetNames.find(n=>norm(n).includes('OPERATIVO'));if(!sn||!wb.Sheets[sn]) throw Error('No se encontró la hoja OPERATIVOS.');
  const raw=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null,blankrows:true});
  if(raw.length<6)throw Error('La hoja OPERATIVOS no contiene las filas esperadas.');
  const h=raw[4].map(norm), ci={orden:column(h,'Nº (ORDEN)','N ORDEN'), numero:column(h,'NRO. OPERATIVO'), provincia:column(h,'PROVINCIA'),distrito:column(h,'DISTRITO'),delito:column(h,'PRESUNTO DELITO'),fecha:column(h,'FECHA DE OPERATIVO'),nombre:column(h,'NOMBRE DE LA OPERACIÓN REALIZADA'),apreh:column(h,'PERSONAS APREHENDIDAS NUMERICO')};
  if(ci.numero<0||ci.provincia<0||ci.delito<0)throw Error('La hoja OPERATIVOS no tiene las columnas esperadas en la fila 5.');
  const get=(row,k)=>ci[k]<0?'':display(row[ci[k]]);
  return raw.slice(5).filter(r=>r && (display(r[ci.orden<0?0:ci.orden]) || get(r,'numero'))).map(row=>({row, headers:raw[4],numero:get(row,'numero'),provincia:get(row,'provincia'),distrito:get(row,'distrito'),delito:get(row,'delito'),fecha:get(row,'fecha'),nombre:get(row,'nombre'),apreh:Number(row[ci.apreh])||0, text:row.map(display).join(' ')}));
 }
 function update(){
  const query=norm($('search').value), prov=$('provincia').value, dis=$('distrito').value, delito=$('delito').value;
  sheet.filtered=sheet.records.filter(r=>(!query||norm(r.text).includes(query))&&(!prov||r.provincia===prov)&&(!dis||r.distrito===dis)&&(!delito||r.delito===delito));
  sheet.page=Math.min(sheet.page,Math.max(0,Math.ceil(sheet.filtered.length/pageSize)-1));render();
 }
 function render(){
  const rs=sheet.filtered; $('summary').textContent=`${rs.length.toLocaleString('es-EC')} resultados de ${sheet.records.length.toLocaleString('es-EC')} registros cargados.`;
  $('metrics').replaceChildren();for(const [name,n] of [['Operativos',rs.length],['Aprehendidos reportados',rs.reduce((a,r)=>a+r.apreh,0)]]){const item=document.createElement('div');item.className='metric';let b=document.createElement('b');b.textContent=n.toLocaleString('es-EC');item.append(b,document.createTextNode(name));$('metrics').append(item);}
  $('results').replaceChildren(); if(!rs.length){const e=document.createElement('div');e.className='empty';e.textContent=sheet.records.length?'Sin resultados con esos filtros.':'Aquí aparecerán los registros de la hoja OPERATIVOS.';$('results').append(e);}
  rs.slice(sheet.page*pageSize,(sheet.page+1)*pageSize).forEach(r=>{const el=document.createElement('button');el.type='button';el.className='result';const strong=document.createElement('strong');strong.textContent=`Operativo ${r.numero || 'sin número'}${r.nombre ? ' · '+r.nombre : ''}`;const sm=document.createElement('small');sm.textContent=[r.fecha,r.provincia,r.distrito,r.delito].filter(Boolean).join(' · ');el.append(strong,sm);el.addEventListener('click',()=>detail(r));$('results').append(el);});
  const pages=Math.ceil(rs.length/pageSize);$('pageinfo').textContent=pages?`Página ${sheet.page+1} de ${pages}`:'';$('prev').disabled=sheet.page===0;$('next').disabled=sheet.page+1>=pages;
 }
 function detail(r){const container=$('detailbody');container.replaceChildren();const p=document.createElement('p');p.className='detailnote';p.textContent='Datos disponibles en la fila del Excel. El PDF se abre por separado; esta versión no vincula los casos automáticamente.';container.append(p);const grid=document.createElement('div');grid.className='detailgrid';r.headers.forEach((h,i)=>{let val=display(r.row[i]);if(!h||!val)return;const div=document.createElement('div');div.className='field';const label=document.createElement('b');label.textContent=String(h);const v=document.createElement('span');v.textContent=val;div.append(label,v);grid.append(div);});container.append(grid);$('detail').hidden=false;$('detail').scrollIntoView({behavior:'smooth',block:'start'});$('detail').focus({preventScroll:true});}
 $('excel').addEventListener('change',async e=>{const file=e.target.files[0];if(!file)return;$('excelstatus').textContent='Leyendo el Excel en este dispositivo…';await new Promise(r=>setTimeout(r,40));try{if(typeof XLSX==='undefined')throw Error('Falta la biblioteca local para leer Excel.');const recs=parseFile(await file.arrayBuffer());sheet.records=recs;sheet.page=0;$('excelstatus').textContent=`${file.name}: ${recs.length.toLocaleString('es-EC')} registros cargados en memoria.`;['search','provincia','distrito','delito'].forEach(id=>$(id).disabled=false);$('search').value='';resetOptions('provincia',recs.map(r=>r.provincia));resetOptions('distrito',recs.map(r=>r.distrito));resetOptions('delito',recs.map(r=>r.delito));$('detail').hidden=true;update();}catch(err){$('excelstatus').textContent='No se pudo leer el archivo: '+err.message;}});
 $('pdf').addEventListener('change',e=>{const f=e.target.files[0];if(!f)return;if(!f.name.toLowerCase().endsWith('.pdf')&&f.type!=='application/pdf'){$('pdfstatus').textContent='Selecciona un archivo PDF.';return;}if(sheet.pdfURL)URL.revokeObjectURL(sheet.pdfURL);sheet.pdfURL=URL.createObjectURL(f);$('pdfview').src=sheet.pdfURL;$('pdflink').href=sheet.pdfURL;$('pdfstatus').textContent=`${f.name} seleccionado en este dispositivo.`;$('pdfsection').hidden=false;});
 for(const id of ['search','provincia','distrito','delito'])$(id).addEventListener(id==='search'?'input':'change',()=>{sheet.page=0;update();});
 $('prev').addEventListener('click',()=>{sheet.page--;render();});$('next').addEventListener('click',()=>{sheet.page++;render();});$('closeDetail').addEventListener('click',()=>{$('detail').hidden=true;});
 $('clear').addEventListener('click',()=>{sheet.records=[];sheet.filtered=[];sheet.page=0;if(sheet.pdfURL)URL.revokeObjectURL(sheet.pdfURL);sheet.pdfURL=null;$('pdfview').removeAttribute('src');$('pdflink').removeAttribute('href');$('pdfsection').hidden=true;$('detail').hidden=true;$('detailbody').replaceChildren();$('excel').value='';$('pdf').value='';$('excelstatus').textContent='Sin archivo seleccionado.';$('pdfstatus').textContent='Sin PDF seleccionado.';['search','provincia','distrito','delito'].forEach(id=>{$(id).value='';$(id).disabled=true;});$('metrics').replaceChildren();$('summary').textContent='Carga el Excel para comenzar.';$('results').replaceChildren();const e=document.createElement('div');e.className='empty';e.textContent='Aquí aparecerán los registros de la hoja OPERATIVOS.';$('results').append(e);$('pageinfo').textContent='';$('prev').disabled=true;$('next').disabled=true;});
})();
