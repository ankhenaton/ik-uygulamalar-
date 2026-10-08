/*
 * SGK / İK Hesaplama Merkezi
 * Bordro Denetim Motoru V3
 * Model: standart şablon + yapısal doğrulama + istisna tabanlı denetim
 * Veri işleme: yalnızca tarayıcıda; sunucuya yükleme yapılmaz.
 */

const TEMPLATE_VERSION = 'PAYROLL-AUDIT-TR-2026-V3';
const TEMPLATE_SHEET = 'Bordro_Veri';
const PARAM_SHEET = 'Parametreler';

const HEADERS = [
  'Sicil No','Ad Soyad','Bordro Dönemi','İşyeri Sicil No','Organizasyon Birimi','Pozisyon',
  'Sigortalılık Statüsü','İşe Giriş Tarihi','İşten Çıkış Tarihi','Prim Günü','Eksik Gün','Eksik Gün Kodu',
  'Normal Brüt Ücret','Fazla Mesai','Prim / İkramiye','Brüt Yemek Ödemesi','Brüt Yol Ödemesi','Diğer Brüt Ek Ödeme',
  'Brüt Toplam','SGK PEK','SGK İşçi Oranı %','SGK İşçi Primi','İşsizlik İşçi Oranı %','İşsizlik İşçi Primi',
  'GV Matrahı','Kümülatif GV Matrahı','Gelir Vergisi','Asgari Ücret GV İstisnası','Damga Vergisi',
  'Asgari Ücret Damga Vergisi İstisnası','BES / Özel Kesintiler','İcra / Haciz Kesintisi','Diğer Kesintiler',
  'Net Ücret','İşveren SGK Primi','İşveren İşsizlik Primi','Toplam İşveren Maliyeti','Teşvik / İstisna Durumu','Açıklama'
];

const REQUIRED_HEADERS = [
  'Sicil No','Ad Soyad','Bordro Dönemi','Sigortalılık Statüsü','Prim Günü','Eksik Gün',
  'Normal Brüt Ücret','Brüt Toplam','SGK PEK','SGK İşçi Primi','İşsizlik İşçi Primi',
  'GV Matrahı','Kümülatif GV Matrahı','Gelir Vergisi','Damga Vergisi','Net Ücret','Toplam İşveren Maliyeti'
];

const PARAM_DEFAULTS = {
  minWage: 33030,
  pekMin: 33030,
  pekMax: 297270,
  dailyPekMin: 1101,
  dailyPekMax: 9909,
  sgkEmp: 0.14,
  unempEmp: 0.01,
  stamp: 0.00759,
  daysMax: 30,
  variancePct: 0.25
};

const APP = {
  workbook: null,
  rows: [],
  headers: [],
  findings: [],
  params: {...PARAM_DEFAULTS},
  validation: null,
  totals: {},
  sourceFile: '',
  period: '',
  sheetName: ''
};

const $ = id => document.getElementById(id);
const escapeHtml = s => String(s ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

function num(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  let s = String(v).trim().replace(/\s/g, '');
  if (!s) return null;
  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',','.');
    else s = s.replace(/,/g,'');
  } else if (s.includes(',')) {
    s = s.replace(',','.');
  }
  s = s.replace(/[^0-9+\-.]/g,'');
  const x = Number(s);
  return Number.isFinite(x) ? x : null;
}

function fmt(x) {
  return x == null || !Number.isFinite(Number(x)) ? '—' : Number(x).toLocaleString('tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
}

function pct(x) {
  return x == null ? '—' : (Number(x)*100).toLocaleString('tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2}) + '%';
}

function parseDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && window.XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m-1, d.d);
  }
  const s = String(v ?? '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(+m[1],+m[2]-1,+m[3]);
  m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function periodNorm(v) {
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}`;
  const s = String(v ?? '').trim();
  const m = s.match(/^(\d{4})[-/.](\d{1,2})$/);
  return m ? `${m[1]}-${String(+m[2]).padStart(2,'0')}` : s;
}

function periodBounds(period) {
  const m = String(period).match(/^(\d{4})-(\d{2})$/);
  if (!m) return null;
  const y=+m[1], mo=+m[2];
  if (mo<1 || mo>12) return null;
  return {start:new Date(y,mo-1,1), end:new Date(y,mo,0), year:y, month:mo, days:new Date(y,mo,0).getDate()};
}

function addFinding(type, code, msg, row=0, detail='', field='') {
  APP.findings.push({type,code,msg,row,detail,field});
}

function csvCell(v) { return '"' + String(v ?? '').replace(/"/g,'""') + '"'; }

function downloadBlob(blob, name) {
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}

function requiredHeaderCheck(headers) {
  const missing = REQUIRED_HEADERS.filter(h => !headers.includes(h));
  const unknown = headers.filter(h => !HEADERS.includes(h));
  const orderedIssues = HEADERS.filter((h,i) => headers[i] !== h).slice(0,8);
  return {missing, unknown, orderedIssues, exact:missing.length===0};
}

function paramFromSheet(wb) {
  const p = {...PARAM_DEFAULTS};
  if (!wb.Sheets[PARAM_SHEET]) return {params:p, found:false};
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[PARAM_SHEET], {header:1, defval:'', raw:true});
  const map = new Map();
  rows.slice(1).forEach(r => { if (String(r[0]??'').trim()) map.set(String(r[0]).trim(), r[1]); });
  const get = k => num(map.get(k));
  p.minWage = get('Brüt asgari ücret') ?? p.minWage;
  p.pekMin = get('PEK alt sınır') ?? p.pekMin;
  p.pekMax = get('PEK üst sınır') ?? p.pekMax;
  p.dailyPekMin = get('Günlük PEK alt sınır') ?? p.dailyPekMin;
  p.dailyPekMax = get('Günlük PEK üst sınır') ?? p.dailyPekMax;
  p.sgkEmp = get('SGK işçi oranı') ?? p.sgkEmp;
  p.unempEmp = get('İşsizlik işçi oranı') ?? p.unempEmp;
  p.stamp = get('Damga vergisi oranı') ?? p.stamp;
  p.daysMax = get('Prim günü üst sınırı') ?? p.daysMax;
  p.variancePct = (get('Aykırılık inceleme eşiği') ?? 25) / 100;
  return {params:p, found:true, version:String(map.get('Şablon sürümü')??'')};
}

async function readWorkbook(file) {
  if (/\.csv$/i.test(file.name)) {
    throw new Error('CSV kabul ediliyor ancak profesyonel denetim için bu sayfada standart Excel şablonunun .xlsx olarak kullanılması önerilir.');
  }
  if (!window.XLSX) throw new Error('Excel motoru yüklenemedi. Sayfayı yenileyip tekrar deneyin.');
  const ab=await file.arrayBuffer();
  const wb=XLSX.read(ab,{type:'array',cellDates:true,raw:true});
  if (!wb.SheetNames.includes(TEMPLATE_SHEET)) throw new Error(`"${TEMPLATE_SHEET}" sayfası bulunamadı. Lütfen sitedeki Bordro Denetim Şablonunu kullanın.`);
  const ws=wb.Sheets[TEMPLATE_SHEET];
  const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:false});
  if (!matrix.length) throw new Error('Bordro_Veri sayfası boş.');
  const headers=matrix[0].map(x=>String(x??'').trim());
  const check=requiredHeaderCheck(headers);
  if (check.missing.length) throw new Error('Şablon eksik. Eksik zorunlu kolonlar: '+check.missing.join(', '));
  const rows=matrix.slice(1)
    .filter(r=>r.some(v=>String(v??'').trim()!==''))
    .map(arr=>Object.fromEntries(headers.map((h,i)=>[h,arr[i] ?? ''])));
  if (!rows.length) throw new Error('Bordro_Veri sayfasında en az bir personel satırı bulunmalı.');
  return {wb,headers,rows,sheetName:TEMPLATE_SHEET,paramInfo:paramFromSheet(wb),matrix};
}

function setStatus(title, detail, type='info') {
  const el=$('schemaStatus');
  el.className='schema-status '+type;
  el.innerHTML=`<b>${escapeHtml(title)}</b><br><span>${escapeHtml(detail)}</span>`;
}

function validateWorkbook(obj) {
  const {headers,rows,paramInfo,wb}=obj;
  const check=requiredHeaderCheck(headers);
  APP.params=paramInfo.params;
  APP.validation={...check,version:paramInfo.version||'',sheets:wb.SheetNames.slice(),rows:rows.length};
  let errs=[];
  let warns=[];
  if (!paramInfo.found) warns.push('Parametreler sayfası bulunamadı; 2026 varsayılanları kullanılacak.');
  else if (paramInfo.version && paramInfo.version !== TEMPLATE_VERSION) errs.push(`Şablon sürümü ${paramInfo.version}; beklenen ${TEMPLATE_VERSION}.`);
  else if (!paramInfo.version) warns.push('Şablon sürümü bulunamadı.');
  if (check.missing.length) errs.push('Eksik zorunlu kolon: '+check.missing.join(', '));
  if (check.unknown.length) warns.push('Ek/standart dışı kolonlar bulundu: '+check.unknown.join(', '));
  if (!headers.includes('AM') && false) {} // future compatibility guard
  if (errs.length) {
    setStatus('Şablon doğrulaması başarısız', errs.join(' | '), 'error');
  } else if (warns.length) {
    setStatus('Şablon geçerli — uzman uyarısı mevcut', warns.join(' | '), 'warn');
  } else {
    setStatus('Şablon doğrulandı', `${rows.length} personel satırı, ${headers.length} kolon, sürüm ${paramInfo.version || TEMPLATE_VERSION}.`, 'ok');
  }
  return {ok:errs.length===0,errors:errs,warnings:warns,check};
}

function rowVal(r,h) { return r[h]; }

function auditRows(rows, headers) {
  APP.findings=[];
  APP.totals={gross:0,pek:0,net:0,cost:0,employees:rows.length,rows:rows.length};
  const p=APP.params;
  const seen=new Map();
  const periods=[];
  const grossVals=[];

  rows.forEach((r,idx)=>{
    const excelRow=idx+2;
    const sicil=String(rowVal(r,'Sicil No')??'').trim();
    const ad=String(rowVal(r,'Ad Soyad')??'').trim();
    const period=periodNorm(rowVal(r,'Bordro Dönemi'));
    const status=String(rowVal(r,'Sigortalılık Statüsü')??'').trim();
    const incentive=String(rowVal(r,'Teşvik / İstisna Durumu')??'').trim().toLowerCase('tr-TR');
    const statSpecial=status && status!=='4A Normal';

    if (!sicil) addFinding('ERROR','COM-001','Sicil numarası boş',excelRow,'Personel kimliği zorunludur.','Sicil No');
    else if (seen.has(sicil)) addFinding('ERROR','DUP-001','Mükerrer sicil numarası',excelRow,`İlk kayıt Excel ${seen.get(sicil)} satırında.`,'Sicil No');
    else seen.set(sicil,excelRow);
    if (!ad) addFinding('WARN','COM-002','Ad soyad alanı boş',excelRow,'Kimlik doğrulama ve izlenebilirlik için tamamlanmalı.','Ad Soyad');
    if (period) periods.push(period); else addFinding('ERROR','PER-001','Bordro dönemi boş/okunamadı',excelRow,'YYYY-AA biçimi beklenir.','Bordro Dönemi');
    if (status!=='4A Normal' && status!=='SGDP' && status!=='Diğer') addFinding('WARN','COM-003','Sigortalılık statüsü standart seçeneklerden değil',excelRow,'Analitik prim kontrolleri temkinli uygulanır.','Sigortalılık Statüsü');

    const numFields=['Prim Günü','Eksik Gün','Normal Brüt Ücret','Fazla Mesai','Prim / İkramiye','Brüt Yemek Ödemesi','Brüt Yol Ödemesi','Diğer Brüt Ek Ödeme','Brüt Toplam','SGK PEK','SGK İşçi Oranı %','SGK İşçi Primi','İşsizlik İşçi Oranı %','İşsizlik İşçi Primi','GV Matrahı','Kümülatif GV Matrahı','Gelir Vergisi','Asgari Ücret GV İstisnası','Damga Vergisi','Asgari Ücret Damga Vergisi İstisnası','BES / Özel Kesintiler','İcra / Haciz Kesintisi','Diğer Kesintiler','Net Ücret','İşveren SGK Primi','İşveren İşsizlik Primi','Toplam İşveren Maliyeti'];
    const v={};
    numFields.forEach(h=>{
      const raw=rowVal(r,h);
      if (raw!=='' && raw!==null && raw!==undefined && num(raw)===null) addFinding('ERROR','DAT-001','Sayısal alan okunamıyor',excelRow,`Alan: ${h}; değer: ${String(raw)}` ,h);
      v[h]=num(raw);
    });

    ['Normal Brüt Ücret','Fazla Mesai','Prim / İkramiye','Brüt Yemek Ödemesi','Brüt Yol Ödemesi','Diğer Brüt Ek Ödeme','Brüt Toplam','SGK PEK','SGK İşçi Primi','İşsizlik İşçi Primi','GV Matrahı','Kümülatif GV Matrahı','Gelir Vergisi','Damga Vergisi','BES / Özel Kesintiler','İcra / Haciz Kesintisi','Diğer Kesintiler','Net Ücret','İşveren SGK Primi','İşveren İşsizlik Primi','Toplam İşveren Maliyeti'].forEach(h=>{
      if (v[h]!=null && v[h]<0) addFinding('ERROR','AMT-001','Negatif parasal değer',excelRow,`${h}: ${fmt(v[h])}`,h);
    });

    const days=v['Prim Günü'], missing=v['Eksik Gün'];
    if (days!=null && (days<0 || days>p.daysMax)) addFinding('ERROR','DAY-001','Prim günü geçersiz',excelRow,`${fmt(days)} gün; izin verilen aralık 0–${p.daysMax}.`,'Prim Günü');
    if (missing!=null && (missing<0 || missing>p.daysMax)) addFinding('ERROR','DAY-002','Eksik gün geçersiz',excelRow,`${fmt(missing)} gün; izin verilen aralık 0–${p.daysMax}.`,'Eksik Gün');
    const join=parseDate(rowVal(r,'İşe Giriş Tarihi')), leave=parseDate(rowVal(r,'İşten Çıkış Tarihi')), pb=periodBounds(period);
    if (join && leave && join>leave) addFinding('ERROR','DATE-001','İşe giriş tarihi işten çıkış tarihinden sonra',excelRow,'Tarih sırası tutarsız.','İşe Giriş Tarihi');
    if (pb && join && join>pb.end) addFinding('WARN','DATE-002','İşe giriş tarihi bordro döneminden sonra',excelRow,'Personel bu dönemde çalışmıyor olabilir.','İşe Giriş Tarihi');
    if (pb && leave && leave<pb.start) addFinding('WARN','DATE-003','İşten çıkış tarihi bordro döneminden önce',excelRow,'Personel bu dönemde çalışmıyor olabilir.','İşten Çıkış Tarihi');
    if (days!=null && missing!=null && days+missing>30) addFinding('ERROR','DAY-003','Prim günü + eksik gün 30 günü aşıyor',excelRow,`Toplam ${fmt(days+missing)} gün.`,'Prim Günü');
    if (!join && !leave && days!=null && missing!=null && days+missing!==30) addFinding('WARN','DAY-004','Tam ay kaydında prim günü + eksik gün 30 değil',excelRow,`Toplam ${fmt(days+missing)} gün. İşe giriş/çıkış/özel durum bilgisi yok.`,'Prim Günü');
    if (join || leave) {
      if (days===0 && (v['Normal Brüt Ücret']||0)>0) addFinding('WARN','DAY-005','Sıfır prim günü ile ücret kaydı birlikte bulunuyor',excelRow,'İşe giriş/çıkış, ücretsiz izin veya özel sigortalılık nedeni doğrulanmalı.','Prim Günü');
    }

    const grossComponents=['Normal Brüt Ücret','Fazla Mesai','Prim / İkramiye','Brüt Yemek Ödemesi','Brüt Yol Ödemesi','Diğer Brüt Ek Ödeme'];
    const grossCalc=grossComponents.reduce((s,h)=>s+(v[h]||0),0);
    if (grossCalc>0) APP.totals.gross+=grossCalc;
    if (v['Brüt Toplam']!=null && grossCalc>=0 && Math.abs(v['Brüt Toplam']-grossCalc)>0.01) addFinding('ERROR','REC-001','Brüt toplam mutabakatı tutmuyor',excelRow,`Bordro: ${fmt(v['Brüt Toplam'])}; kalem toplamı: ${fmt(grossCalc)}.`,'Brüt Toplam');
    if (v['Brüt Toplam']==null && grossCalc>0) addFinding('ERROR','REC-002','Brüt toplam boş',excelRow,'Şablondaki formül/alan korunmalı.','Brüt Toplam');

    const pek=v['SGK PEK'];
    if (pek!=null) APP.totals.pek+=pek;
    if (days!=null && days>0 && pek!=null) {
      const min= p.dailyPekMin*days;
      const max= p.dailyPekMax*days;
      if (pek < min-0.01 && !incentive.includes('istisna') && !statSpecial) addFinding('WARN','SGK-001','PEK alt sınır analitik kontrolünde düşük',excelRow,`PEK ${fmt(pek)}; yaklaşık günlük alt sınır karşılığı ${fmt(min)}. Özel kazanç/istisna durumları ayrıca incelenmeli.`,'SGK PEK');
      if (pek > max+0.01) addFinding('ERROR','SGK-002','PEK üst sınırını aşıyor',excelRow,`PEK ${fmt(pek)}; ${days} gün için tavan yaklaşık ${fmt(max)}.`,'SGK PEK');
    }
    const gross=v['Brüt Toplam']!=null ? v['Brüt Toplam'] : grossCalc;
    if (gross>0 && pek!=null && pek>gross*1.5) addFinding('WARN','SGK-003','PEK, brüt toplama göre olağandışı yüksek',excelRow,'Ücret dışı prime tabi ödemeler, geçmiş dönem aktarımı veya özel kayıtlar incelenmeli.','SGK PEK');
    if (gross>0 && pek!=null && pek<gross*0.5 && !incentive.includes('istisna') && !statSpecial) addFinding('INFO','SGK-004','PEK, brüt toplama göre belirgin düşük',excelRow,'Prime tabi olmayan kalemler/istisnalar/özel durumlar kontrol edilebilir.','SGK PEK');

    // Standard 4A analytical premium checks; special statuses are deliberately downgraded/omitted.
    const sgkRate=v['SGK İşçi Oranı %']!=null ? v['SGK İşçi Oranı %']/100 : p.sgkEmp;
    const unRate=v['İşsizlik İşçi Oranı %']!=null ? v['İşsizlik İşçi Oranı %']/100 : p.unempEmp;
    if (!statSpecial && pek!=null) {
      const expSgk=pek*sgkRate;
      if (v['SGK İşçi Primi']!=null && Math.abs(v['SGK İşçi Primi']-expSgk)>Math.max(1,expSgk*0.02)) addFinding('WARN','PRM-001','SGK işçi primi analitik kontrolden farklı',excelRow,`Raporlanan ${fmt(v['SGK İşçi Primi'])}; oran kontrolü ${fmt(expSgk)}. Teşvik/istisna/yuvarlama parametreleri nedeniyle kesin hata değildir.`,'SGK İşçi Primi');
      const expUn=pek*unRate;
      if (v['İşsizlik İşçi Primi']!=null && Math.abs(v['İşsizlik İşçi Primi']-expUn)>Math.max(1,expUn*0.02)) addFinding('WARN','PRM-002','İşsizlik işçi primi analitik kontrolden farklı',excelRow,`Raporlanan ${fmt(v['İşsizlik İşçi Primi'])}; oran kontrolü ${fmt(expUn)}.`,'İşsizlik İşçi Primi');
    } else if ((v['SGK İşçi Primi']!=null || v['İşsizlik İşçi Primi']!=null)) {
      addFinding('INFO','PRM-003','Özel sigortalılık nedeniyle oran denetimi temkinli',excelRow,`Statü: ${status}. Prim tutarları kayıt/belge ile ayrıca doğrulanmalı.`,'Sigortalılık Statüsü');
    }

    const gv=v['GV Matrahı'], cum=v['Kümülatif GV Matrahı'], tax=v['Gelir Vergisi'], stamp=v['Damga Vergisi'];
    if (gv!=null && gv<0) addFinding('ERROR','TAX-001','Gelir vergisi matrahı negatif',excelRow,'Matrah tutarı incelenmeli.','GV Matrahı');
    if (tax!=null && tax<0) addFinding('ERROR','TAX-002','Gelir vergisi negatif',excelRow,'Vergi alanı negatif olamaz.','Gelir Vergisi');
    if (stamp!=null && stamp<0) addFinding('ERROR','TAX-003','Damga vergisi negatif',excelRow,'Vergi alanı negatif olamaz.','Damga Vergisi');
    if (cum!=null && gv!=null && cum+0.01<gv) addFinding('ERROR','TAX-004','Kümülatif matrah mevcut dönem matrahından küçük',excelRow,`Kümülatif ${fmt(cum)}; dönem ${fmt(gv)}.`,'Kümülatif GV Matrahı');
    if (cum!=null && gv!=null && gv>0 && Math.abs(cum-gv)<0.01 && period && /^\d{4}-0[1]$/.test(period)) addFinding('INFO','TAX-005','Kümülatif matrah dönem matrahına eşit',excelRow,'Ocak dönemi için normal olabilir; dönem bilgisi doğrulanmalı.','Kümülatif GV Matrahı');

    if (gross>0 && stamp!=null) {
      const expectedStamp=gross*p.stamp;
      if (Math.abs(stamp-expectedStamp)>Math.max(1,expectedStamp*0.05)) addFinding('INFO','TAX-006','Damga vergisi analitik tutardan farklı',excelRow,`Raporlanan ${fmt(stamp)}; basit oran kontrolü ${fmt(expectedStamp)}. Asgari ücret istisnası ve diğer istisnalar nedeniyle kesin hata değildir.`,'Damga Vergisi');
    }

    const sgk=v['SGK İşçi Primi']||0, un=v['İşsizlik İşçi Primi']||0;
    const otherDed=(v['BES / Özel Kesintiler']||0)+(v['İcra / Haciz Kesintisi']||0)+(v['Diğer Kesintiler']||0);
    if (gross!=null && v['Net Ücret']!=null) {
      const expectedNet=gross-sgk-un-(tax||0)-(stamp||0)-otherDed;
      APP.totals.net+=v['Net Ücret'];
      if (Math.abs(v['Net Ücret']-expectedNet)>2) addFinding('WARN','REC-003','Net ücret mutabakatı farklı',excelRow,`Raporlanan ${fmt(v['Net Ücret'])}; temel kesintiler sonrası ${fmt(expectedNet)}. Bordroda bu şablonda bulunmayan ödeme/kesinti kalemi varsa inceleyin.`,'Net Ücret');
      if (v['Net Ücret']>gross+0.01) addFinding('ERROR','NET-001','Net ücret brüt toplamdan yüksek',excelRow,`Net ${fmt(v['Net Ücret'])}; brüt ${fmt(gross)}.`,'Net Ücret');
      if (gross>0 && Math.abs(gross-v['Net Ücret'])/gross>0.70) addFinding('INFO','NET-002','Brüt-net farkı yüksek',excelRow,'İcra/haciz, BES, diğer kesintiler, vergi ve özel ödemelerle açıklanıp açıklanmadığı incelenmeli.','Net Ücret');
    }

    const employer=v['Toplam İşveren Maliyeti'];
    if (employer!=null) {
      APP.totals.cost+=employer;
      if (gross>0 && employer<gross-0.01) addFinding('ERROR','COST-001','Toplam işveren maliyeti brütün altında',excelRow,`Maliyet ${fmt(employer)}; brüt ${fmt(gross)}.`,'Toplam İşveren Maliyeti');
    }
    if (v['BES / Özel Kesintiler']!=null && v['BES / Özel Kesintiler']<0) addFinding('ERROR','CUT-001','BES / özel kesinti negatif',excelRow,'Kesinti alanı kontrol edilmeli.','BES / Özel Kesintiler');
    if (v['İcra / Haciz Kesintisi']!=null && v['İcra / Haciz Kesintisi']<0) addFinding('ERROR','CUT-002','İcra / haciz kesintisi negatif',excelRow,'Kesinti alanı kontrol edilmeli.','İcra / Haciz Kesintisi');

    if (gross>0) grossVals.push({row:excelRow,gross});
  });

  // Cross-row period consistency
  const uniquePeriods=[...new Set(periods.filter(Boolean))];
  APP.period=uniquePeriods.length===1?uniquePeriods[0]:(uniquePeriods.length>1?'KARIŞIK':'');
  if (uniquePeriods.length>1) addFinding('ERROR','PER-002','Tek dosyada birden fazla bordro dönemi bulundu',0,uniquePeriods.join(', '),'Bordro Dönemi');

  // Duplicate person-name / sicil anomaly is separate from exact duplicate.
  const nameMap=new Map();
  rows.forEach((r,idx)=>{
    const name=String(r['Ad Soyad']??'').trim().toLocaleLowerCase('tr-TR');
    if (!name) return;
    if (nameMap.has(name) && String(r['Sicil No']??'').trim()!==String(rows[nameMap.get(name)]['Sicil No']??'').trim()) {
      addFinding('INFO','DUP-002','Aynı ad soyad farklı sicil numaralarında geçiyor',idx+2,`İlk görüldüğü Excel satırı ${nameMap.get(name)+2}. Kimlik doğrulaması yapın.`,'Ad Soyad');
    } else if(!nameMap.has(name)) nameMap.set(name,idx);
  });

  // Robust outlier: median + configurable multiplier; information only.
  if (grossVals.length>=7) {
    const sorted=grossVals.map(x=>x.gross).sort((a,b)=>a-b);
    const med=sorted[Math.floor(sorted.length/2)];
    const limit=1 + (APP.params.variancePct*4);
    grossVals.forEach(x=>{
      if (med>0 && (x.gross>med*limit || x.gross<med/limit)) addFinding('INFO','OUT-001','Ücret dağılımında güçlü aykırı değer',x.row,`Medyan ${fmt(med)}; kayıt ${fmt(x.gross)}.`,'Brüt Toplam');
    });
  }

  // Aggregate totals
  APP.stats={
    rows:rows.length,
    errors:APP.findings.filter(f=>f.type==='ERROR').length,
    warns:APP.findings.filter(f=>f.type==='WARN').length,
    info:APP.findings.filter(f=>f.type==='INFO').length
  };
  return APP.stats;
}

function renderPreview(rows) {
  const showHeaders=['Sicil No','Ad Soyad','Bordro Dönemi','Prim Günü','Normal Brüt Ücret','SGK PEK','Net Ücret','Teşvik / İstisna Durumu'];
  let html='<table class="table compact"><thead><tr>'+showHeaders.map(h=>`<th>${escapeHtml(h)}</th>`).join('')+'</tr></thead><tbody>';
  rows.slice(0,6).forEach(r=>{html+='<tr>'+showHeaders.map(h=>`<td>${escapeHtml(h.includes('Ücret')||h.includes('PEK')?fmt(num(r[h])):r[h])}</td>`).join('')+'</tr>'});
  html+='</tbody></table>';
  if (rows.length>6) html+=`<p class="small muted">İlk 6 satır gösteriliyor. Toplam ${rows.length} kayıt denetlenecek.</p>`;
  $('preview').innerHTML=html;
}

function renderSummary() {
  const s=APP.stats||{};
  const f=APP.findings;
  $('summary').innerHTML=`
    <div class="audit-kpis">
      <div class="kpi-box"><span>Personel</span><strong>${s.rows||0}</strong></div>
      <div class="kpi-box danger"><span>HATA</span><strong>${s.errors||0}</strong></div>
      <div class="kpi-box warn"><span>UYARI</span><strong>${s.warns||0}</strong></div>
      <div class="kpi-box info"><span>İNCELEME</span><strong>${s.info||0}</strong></div>
    </div>
    <div class="summary-lines">
      <div><b>Dönem:</b> ${escapeHtml(APP.period||'Belirsiz')}</div>
      <div><b>Brüt toplam (hesaplanan):</b> ${fmt(APP.totals.gross)} TL</div>
      <div><b>PEK toplam:</b> ${fmt(APP.totals.pek)} TL</div>
      <div><b>Net toplam:</b> ${fmt(APP.totals.net)} TL</div>
      <div><b>İşveren maliyeti toplam:</b> ${fmt(APP.totals.cost)} TL</div>
      <div><b>Bulgu:</b> ${f.length} kayıt</div>
    </div>`;
}

function renderFindings() {
  const typeFilter=$('severityFilter').value;
  const search=String($('findingSearch').value||'').toLocaleLowerCase('tr-TR');
  const list=APP.findings.filter(f=>{
    if(typeFilter!=='ALL' && f.type!==typeFilter) return false;
    const hay=`${f.code} ${f.msg} ${f.detail} ${f.field}`.toLocaleLowerCase('tr-TR');
    return !search || hay.includes(search);
  });
  const group={ERROR:'HATA',WARN:'UYARI',INFO:'İNCELEME'};
  let html='<table class="table"><thead><tr><th>Seviye</th><th>Kod</th><th>Excel Satırı</th><th>Alan</th><th>Denetim</th><th>Açıklama</th></tr></thead><tbody>';
  list.slice(0,1000).forEach(f=>{
    html+=`<tr class="${f.type==='ERROR'?'danger-row':f.type==='WARN'?'warn-row':'info-row'}"><td><b>${group[f.type]}</b></td><td class="mono">${escapeHtml(f.code)}</td><td>${f.row||'—'}</td><td>${escapeHtml(f.field||'')}</td><td>${escapeHtml(f.msg)}</td><td>${escapeHtml(f.detail)}</td></tr>`;
  });
  if(!list.length) html+='<tr><td colspan="6">Seçilen filtrelerle bulgu bulunamadı.</td></tr>';
  html+='</tbody></table>';
  $('findings').innerHTML=html;
}

function runAudit() {
  if(!APP.rows.length) return;
  const validation=APP.validation;
  if(!validation || !validation.exact){
    $('loadMsg').textContent='Şablon doğrulaması geçerli değil; denetim başlatılmadı.';
    return;
  }
  auditRows(APP.rows,APP.headers);
  renderSummary();
  renderFindings();
  $('auditPanel').hidden=false;
  $('exportBtn').disabled=APP.findings.length===0;
  $('auditMsg').textContent=`Denetim tamamlandı. ${APP.stats.errors} hata, ${APP.stats.warns} uyarı, ${APP.stats.info} inceleme bulgusu.`;
}

function exportFindings() {
  if(!APP.findings.length) return;
  const rows=[['Seviye','Kod','Excel Satırı','Alan','Denetim','Açıklama'],...APP.findings.map(f=>[f.type,f.code,f.row,f.field,f.msg,f.detail])];
  const text='\ufeff'+rows.map(r=>r.map(csvCell).join(';')).join('\r\n');
  downloadBlob(new Blob([text],{type:'text/csv;charset=utf-8'}),'bordro-denetim-bulgulari.csv');
}

function exportFullAudit() {
  if(!APP.rows.length) return;
  const rows=[
    ['Bordro Denetim Raporu', '', '', '', '', ''],
    ['Dosya',APP.sourceFile,'Dönem',APP.period,'Şablon',TEMPLATE_VERSION],
    ['Personel',APP.stats.rows,'Hata',APP.stats.errors,'Uyarı',APP.stats.warns],
    [],
    ['Seviye','Kod','Excel Satırı','Alan','Denetim','Açıklama'],
    ...APP.findings.map(f=>[f.type,f.code,f.row,f.field,f.msg,f.detail])
  ];
  const text='\ufeff'+rows.map(r=>r.map(csvCell).join(';')).join('\r\n');
  downloadBlob(new Blob([text],{type:'text/csv;charset=utf-8'}),'bordro-denetim-raporu.csv');
}

function loadFile(file) {
  APP.sourceFile=file.name;
  $('fileName').textContent=file.name;
  $('loadMsg').textContent='Excel okunuyor...';
  readWorkbook(file).then(obj=>{
    APP.workbook=obj.wb; APP.rows=obj.rows; APP.headers=obj.headers; APP.sheetName=obj.sheetName;
    const validation=validateWorkbook(obj);
    renderPreview(APP.rows);
    $('rowCount').textContent=APP.rows.length;
    $('colCount').textContent=APP.headers.length;
    $('runBtn').disabled=!validation.ok;
    $('auditMsg').textContent=validation.ok?'Denetime hazır.':'Denetim için şablon hatalarını düzeltin.';
    if(validation.ok) {
      $('paramInfo').textContent=`Excel içindeki Parametreler sayfası okundu. PEK: ${fmt(APP.params.pekMin)}–${fmt(APP.params.pekMax)} TL/ay · SGK işçi: ${pct(APP.params.sgkEmp)} · İşsizlik işçi: ${pct(APP.params.unempEmp)}.`;
    }
  }).catch(err=>{
    APP.rows=[]; APP.validation=null; $('runBtn').disabled=true; $('auditPanel').hidden=true;
    setStatus('Dosya yüklenemedi',err.message,'error');
    $('loadMsg').textContent=err.message;
  });
}

function loadDemo() {
  const rows=[
    {
      'Sicil No':'1001','Ad Soyad':'Örnek Personel 1','Bordro Dönemi':'2026-10','Sigortalılık Statüsü':'4A Normal','Prim Günü':30,'Eksik Gün':0,
      'Normal Brüt Ücret':50000,'Fazla Mesai':2500,'Prim / İkramiye':0,'Brüt Yemek Ödemesi':0,'Brüt Yol Ödemesi':0,'Diğer Brüt Ek Ödeme':0,'Brüt Toplam':52500,
      'SGK PEK':52500,'SGK İşçi Oranı %':14,'SGK İşçi Primi':7350,'İşsizlik İşçi Oranı %':1,'İşsizlik İşçi Primi':525,'GV Matrahı':44625,'Kümülatif GV Matrahı':510000,
      'Gelir Vergisi':6693.75,'Asgari Ücret GV İstisnası':0,'Damga Vergisi':398.48,'Asgari Ücret Damga Vergisi İstisnası':0,'BES / Özel Kesintiler':500,'İcra / Haciz Kesintisi':0,'Diğer Kesintiler':0,
      'Net Ücret':37032.77,'İşveren SGK Primi':8500,'İşveren İşsizlik Primi':1050,'Toplam İşveren Maliyeti':62050,'Teşvik / İstisna Durumu':'Yok','Açıklama':'Demo'
    },
    {
      'Sicil No':'1002','Ad Soyad':'Örnek Personel 2','Bordro Dönemi':'2026-10','Sigortalılık Statüsü':'4A Normal','Prim Günü':30,'Eksik Gün':0,
      'Normal Brüt Ücret':60000,'Fazla Mesai':0,'Prim / İkramiye':0,'Brüt Yemek Ödemesi':0,'Brüt Yol Ödemesi':0,'Diğer Brüt Ek Ödeme':0,'Brüt Toplam':60000,
      'SGK PEK':310000,'SGK İşçi Oranı %':14,'SGK İşçi Primi':8400,'İşsizlik İşçi Oranı %':1,'İşsizlik İşçi Primi':600,'GV Matrahı':51000,'Kümülatif GV Matrahı':561000,
      'Gelir Vergisi':8000,'Asgari Ücret GV İstisnası':0,'Damga Vergisi':455.4,'Asgari Ücret Damga Vergisi İstisnası':0,'BES / Özel Kesintiler':0,'İcra / Haciz Kesintisi':0,'Diğer Kesintiler':0,
      'Net Ücret':50944.6,'İşveren SGK Primi':9500,'İşveren İşsizlik Primi':1200,'Toplam İşveren Maliyeti':70700,'Teşvik / İstisna Durumu':'Yok','Açıklama':'Demo — PEK tavan hatası örneği'
    },
    {
      'Sicil No':'1002','Ad Soyad':'Mükerrer Kayıt','Bordro Dönemi':'2026-10','Sigortalılık Statüsü':'4A Normal','Prim Günü':30,'Eksik Gün':0,
      'Normal Brüt Ücret':45000,'Fazla Mesai':0,'Prim / İkramiye':0,'Brüt Yemek Ödemesi':0,'Brüt Yol Ödemesi':0,'Diğer Brüt Ek Ödeme':0,'Brüt Toplam':45000,
      'SGK PEK':45000,'SGK İşçi Oranı %':14,'SGK İşçi Primi':6300,'İşsizlik İşçi Oranı %':1,'İşsizlik İşçi Primi':450,'GV Matrahı':38250,'Kümülatif GV Matrahı':548250,
      'Gelir Vergisi':5737.5,'Asgari Ücret GV İstisnası':0,'Damga Vergisi':341.55,'Asgari Ücret Damga Vergisi İstisnası':0,'BES / Özel Kesintiler':0,'İcra / Haciz Kesintisi':0,'Diğer Kesintiler':0,
      'Net Ücret':32170.95,'İşveren SGK Primi':7500,'İşveren İşsizlik Primi':900,'Toplam İşveren Maliyeti':53400,'Teşvik / İstisna Durumu':'Yok','Açıklama':'Demo — duplicate örneği'
    }
  ];
  APP.sourceFile='demo-bordro.xlsx'; APP.rows=rows; APP.headers=HEADERS.slice(); APP.sheetName=TEMPLATE_SHEET;
  APP.validation={exact:true,version:TEMPLATE_VERSION}; APP.params={...PARAM_DEFAULTS};
  setStatus('Örnek bordro hazır','Şablon yapısı simüle edildi. Şimdi Denetimi Çalıştır düğmesine basın.','ok');
  renderPreview(rows); $('rowCount').textContent=rows.length; $('colCount').textContent=HEADERS.length; $('runBtn').disabled=false; $('auditPanel').hidden=true;
}

$('fileInput').addEventListener('change',e=>{const f=e.target.files?.[0]; if(f) loadFile(f);});
$('runBtn').addEventListener('click',runAudit);
$('demoBtn').addEventListener('click',loadDemo);
$('exportBtn').addEventListener('click',exportFullAudit);
$('exportFindingsBtn').addEventListener('click',exportFindings);
$('severityFilter').addEventListener('change',renderFindings);
$('findingSearch').addEventListener('input',renderFindings);
$('resetBtn').addEventListener('click',()=>{location.reload();});

setStatus('Dosya bekleniyor','Sitedeki resmi Excel şablonunu indirip doldurun; kolon adlarını değiştirmeden tekrar yükleyin.','info');
