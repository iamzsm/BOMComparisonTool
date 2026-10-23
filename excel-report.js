// BOM Studio system-import Excel writer. Strict seven-value Modify Mode vocabulary.
const NAVY='FF183153',PALE='FFEAF0F7',AMBER='FFFFE9B5';
const fill=c=>({type:'pattern',pattern:'solid',fgColor:{argb:c}});
const simple=x=>x==null?'':typeof x==='object'?x.text??x.result??x.richText?.map(t=>t.text).join('')??'':x;
function heading(ws,row,title){let c=ws.getCell(row,1);c.value=title;c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill=fill(NAVY);ws.getRow(row).height=26}
function fields(ws,row,values,flag=false){values.forEach((v,i)=>{let c=ws.getCell(row,i+1);c.value=simple(v);c.alignment={wrapText:true,vertical:'top'};if(flag)c.fill=fill(AMBER)});ws.getRow(row).height=34}
const role=p=>p?.sub?'Sub':p?.main?'Main':'';
const addMode=p=>p?.sub?'Add Sub':p?.main&&p?.group?'Add Main':'Add';
export async function makeReport(wb,ws,hdr,result,oldFile,newFile){
  const origWidth=Math.max(...Object.values(hdr));const mainCol=hdr['MAIN/SUB'];
  if(!mainCol)throw Error('OLD BOM is missing Main/Sub header');
  // Store source data and formatting before reconstructing rows. Other worksheets
  // remain in workbook; this prototype cannot guarantee formulas/merges survive.
  const entries=[];const rowsByOld=new Map();
  for(const p of result.old){let values=[],styles=[];for(let c=1;c<=origWidth;c++){
    const x=ws.getRow(p.row).getCell(c);values.push(x.value);styles.push(structuredClone(x.style||{}));
  }rowsByOld.set(p.row,{values,styles});}
  // Unchanged OLD rows, including deleted originals, keep historical SBR IDs.
  for(const p of result.old){if(result.retained.has(p.row))continue;
    const row=rowsByOld.get(p.row);const mode=result.modes.get(p.row)||'';
    const match=result.matches.find(([a])=>a.row===p.row)?.[1];
    // Respect explicitly reported new Main/Sub role, even when nothing else changed.
    if(match&&mode!=='Delete'&&role(match))row.values[mainCol-1]=role(match);
    if(mode!=='Delete')for(const [field,nv] of result.changes.get(p.row)||[]){
      let col=hdr[field.toUpperCase()];if(col)row.values[col-1]=nv;
    }
    entries.push({values:row.values,styles:row.styles,mode,group:p.group,part:p,review:result.manual.has(p.row),oldRow:p.row,after:p.row});
  }
  const materialize=(p,mode,review=false,after=null,forcedGroup='')=>{
    const values=Array(origWidth).fill(null);for(const [key,c] of Object.entries(hdr)){
      if(['SBR ID','HOME','NO.'].includes(key))continue;
      values[c-1]=simple(p.data[key]);
    }
    if(role(p))values[mainCol-1]=role(p);
    return {values,styles:null,mode,group:forcedGroup||p.group,part:p,review,after,oldRow:null};
  };
  for(const e of [...result.added,...result.extraAdds]){
    if(e.sourceOldRow){const historical=rowsByOld.get(e.sourceOldRow);
      let values=[...historical.values];if(hdr['SBR ID'])values[hdr['SBR ID']-1]=null;
      if(hdr['HOME'])values[hdr['HOME']-1]=null;
      if(hdr['NO.'])values[hdr['NO.']-1]=null;
      entries.push({values,styles:historical.styles,mode:e.mode,part:e.b,group:e.b.group,review:true,after:e.after,oldRow:null});
    }else entries.push(materialize(e.b,e.mode||addMode(e.b),Boolean(e.review),e.after));
  }
  // For matched group members, keep OLD group identifier if NEW does not supply one.
  const groupMembership=new Map();for(const e of entries){const g=e.group;
    if(!g)continue;if(!groupMembership.has(g))groupMembership.set(g,[]);groupMembership.get(g).push(e)}
  // Stable group-aware sorting: keeps original block placement, Main before Sub,
  // while leaving ungrouped components in their original relative order.
  const ordered=[],emitted=new Set();for(const e of entries){
    if(emitted.has(e))continue;
    const block=e.group?groupMembership.get(e.group):[e];
    const sorted=[...block].sort((a,b)=>{
      const r=x=>String(x.values[mainCol-1]||'').trim().toUpperCase();
      const n=x=>r(x)==='MAIN'?0:r(x)==='SUB'?1:2;
      return n(a)-n(b)||entries.indexOf(a)-entries.indexOf(b);
    });for(const x of sorted){if(!emitted.has(x)){ordered.push(x);emitted.add(x)}}
  }
  // Rebuild a flat table, inserting exactly one column immediately before Main/Sub.
  // Uses header *names*, independent of the physical column order of either input.
  const newWidth=origWidth+1;const header=[];for(let c=1;c<=origWidth;c++)header.push(ws.getRow(1).getCell(c).value);
  header.splice(mainCol-1,0,'Modify Mode');
  const roleColumn=mainCol+1;
  if(ws.rowCount)ws.spliceRows(1,ws.rowCount);
  header.forEach((value,i)=>{let c=ws.getRow(1).getCell(i+1);c.value=value;c.font={bold:true}});
  for(let i=0;i<ordered.length;i++){
    const e=ordered[i],row=ws.getRow(i+2),v=[...e.values];v.splice(mainCol-1,0,e.mode||'');
    for(let c=1;c<=newWidth;c++){
      let cell=row.getCell(c);cell.value=simple(v[c-1]);
      const sourceCol=c>mainCol?c-1:c;
      if(c!==mainCol&&e.styles?.[sourceCol-1]){
        let style=structuredClone(e.styles[sourceCol-1]);
        // No inherited comparison colour: restore all rows to normal no fill.
        delete style.fill;cell.style=style;
      }
      cell.fill={type:'pattern',pattern:'none'};
      if(e.review)cell.fill=fill(AMBER);
    }
    row.getCell(mainCol).alignment={vertical:'middle',wrapText:true};
  }
  ws.getColumn(mainCol).width=24;ws.getColumn(roleColumn).width=16;
  ws.views=[];
  // Worksheet 2 ONLY SECTION E: concise manual import validation.
  // A-D are rendered on the website dashboard, not duplicated in Excel.
  const sum=wb.addWorksheet('Summary & Review');sum.views=[];
  [24,22,21,26,30,55].forEach((w,i)=>sum.getColumn(i+1).width=w);
  heading(sum,1,'E. MAIN/SUB & IMPORT VALIDATION');
  fields(sum,2,['REFERENCE ONLY — manually verify each item before importing the BOM.']);
  fields(sum,4,['Check','OLD Row','NEW Row','Group','Modify Mode','Reason']);
  let rr=5;const diagnostics=[];
  for(const [oldRow,msgs] of result.manual){const p=result.old.find(x=>x.row===oldRow);
    const b=result.matches.find(([a])=>a.row===oldRow)?.[1];
    for(const msg of msgs)diagnostics.push(['MANUAL CHECK',oldRow,b?.row??'',p?.group??'',result.modes.get(oldRow)??'',msg]);}
  for(const rec of result.review)diagnostics.push(['MATCH REVIEW',rec[1],rec[2],rec[3],'',rec[0]]);
  const gm=new Map();for(const e of ordered){if(!e.group||e.mode==='Delete')continue;
    if(!gm.has(e.group))gm.set(e.group,[]);gm.get(e.group).push(e)}
  for(const [g,parts]of gm){const mains=parts.filter(p=>String(p.values[mainCol-1]||'').toUpperCase()==='MAIN');
    if(mains.length!==1)diagnostics.push(['GROUP STRUCTURE','','',g,'',`Expected one Main; found ${mains.length}`]);}
  if(!diagnostics.length)fields(sum,rr++,['No exceptions identified. Manual engineering verification is still required.']);
  else for(const d of diagnostics)fields(sum,rr++,d,true);
  const change=wb.addWorksheet('Change Log');fields(change,1,['Status','Original OLD Row','Original NEW Row','Old SBR ID','Location','Field','Old Value','New Value']);
  result.log.forEach((r,i)=>fields(change,i+2,r));
  const review=wb.addWorksheet('Needs Review');fields(review,1,['Reason','Original OLD Row','Possible NEW Row(s)','Identifier']);
  result.review.forEach((r,i)=>fields(review,i+2,r,true));
  sum.views=[];return wb.xlsx.writeBuffer();
}
