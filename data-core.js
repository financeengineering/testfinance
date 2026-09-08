'use strict';
// Compatible with the current Excel export. Amounts are in millions of RUB.
const Core = (() => {
  const parseM = value => {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string' || !value.trim()) return null;
    const clean = value.replace(/[−–]/g, '-').replace(/[M₽\s]/gi, '').replace(/\./g, '').replace(/,/g, '.');
    return /^-?\d+(\.\d+)?$/.test(clean) ? Number(clean) : null;
  };
  const ratio = (actual, target) => Number.isFinite(actual) && Number.isFinite(target) && target > 0 ? Math.round(actual / target * 100) : null;
  const sum = xs => xs.reduce((a,b) => a + (Number.isFinite(b) ? b : 0), 0);
  const width = n => Math.min(100, Math.max(0, Number.isFinite(n) ? n : 0));
  const id = (kind, label) => kind + '-' + Array.from(String(label)).map(c => c.codePointAt(0).toString(16)).join('-');
  const monthKey = s => { const [m,y] = s.split('.'); return Number(y)*12 + ['Oca','Şub','Mar','Nis','May','Haz','Tem','Ağu','Eyl','Eki','Kas','Ara'].indexOf(m); };
  const gaugeStats = (g,d) => {
    const value = parseM(g.big);
    const gauge = name => parseM(d.gauges.find(x=>x.label===name)?.big);
    const definitions = {
      'Toplam Tahsilat': [parseM(d.contractPotential.total),'Toplam Potansiyel Sözleşme Bedeli'],
      'Saha Tahakkuk': [parseM(d.contractPotential.total),'Toplam Potansiyel Sözleşme Bedeli'],
      'Toplam Harcama': [parseM(d.butce),'Bütçe'],
      'Avans Kesintisi': [parseM(d.avansSummary[0]?.[1]),'Toplam Avans Tahsilat'],
      'Brüt Hakediş': [parseM(d.contractPotential.total),'Toplam Potansiyel Sözleşme Bedeli'],
      'Açık İşler': [gauge('Saha Tahakkuk'),'Saha Tahakkuk'],
      'Borç': [gauge('Toplam Harcama'),'Toplam Harcama'],
      'Banka Bakiye': [parseM(d.contractPotential.total),'Toplam Potansiyel Sözleşme Bedeli']
    };
    const [target,base] = definitions[g.label] || [null,''];
    return {value,target,base,pct:ratio(value,target),contractPct:ratio(value,parseM(d.contractCurrent.total))};
  };
  const monthlyCategories = d => {
    const map = new Map();
    d.aylikDagilim.forEach(m => m.kalemler.forEach(([name,value])=>{
      if (!map.has(name)) map.set(name,{name,total:0,months:[]});
      const k=map.get(name); k.total+=value; k.months.push({month:m.ay,value});
    }));
    return [...map.values()].sort((a,b)=>b.total-a.total);
  };
  const generalSeries = d => {
    const spend=new Map((d.harcamalar?.aylar||[]).map((m,i)=>[m,d.harcamalar.toplam.degerler[i]]));
    const receipts=new Map(d.monthlyNet.map(r=>[r.m,r.tahsilat]));
    const gross=new Map(d.monthlyBrut.map(r=>[r.m,r.v]));
    const months=[...new Set([...spend.keys(),...receipts.keys(),...gross.keys()])].sort((a,b)=>monthKey(a)-monthKey(b));
    return {months,series:[['harcama','Toplam Harcama',spend],['tahsilat','Tahsilat',receipts],['hakedis','Brüt Hakediş',gross]].map(([key,label,map])=>({key,label,values:months.map(m=>map.has(m)?map.get(m):null)}))};
  };
  const validate = d => {
    if (!d || typeof d!=='object' || !d.meta || typeof d.meta.date!=='string') throw Error('schema');
    for (const k of ['gauges','avansSummary','bankaGarantili','bankaGarantisiz','monthlyBrut','monthlyNet','hakedisTipleri','sahaItems','kesintiler','aylikDagilim','brutHakedisDetay']) if (!Array.isArray(d[k])) throw Error('schema:'+k);
    for (const k of ['contractCurrent','contractPotential']) if (!d[k] || parseM(d[k].total)===null || !Array.isArray(d[k].items)) throw Error('schema:'+k);
    if (!d.personel || !d.kalanGun || !d.hakedisTotals?.kdvHaric || !d.hakedisTotals?.kdvli || !d.potansiyelDusulecek) throw Error('schema');
    const number = x => { if (!Number.isFinite(x)) throw Error('schema:number'); };
    const label = x => { if (typeof x !== 'string' || !x.trim()) throw Error('schema:label'); };
    const amount = x => { if (parseM(x) === null) throw Error('schema:amount'); };
    d.gauges.forEach(g => { label(g.label); amount(g.big); });
    for (const key of ['avansSummary','bankaGarantili','bankaGarantisiz','hakedisTipleri']) d[key].forEach(row => {
      if (!Array.isArray(row)) throw Error('schema:row'); label(row[0]); amount(row[1]); amount(row[2]);
    });
    for (const key of ['brutHakedisDetay','kesintiler']) d[key].forEach(row => {
      if (!Array.isArray(row)) throw Error('schema:row'); label(row[0]); amount(row[1]);
    });
    d.sahaItems.forEach(k => { label(k.name); label(k.unit); [k.proje,k.gerc,k.kalan,k.ay?.plan,k.ay?.gerc,k.ay?.kalan,k.haftalik].forEach(number); });
    d.monthlyBrut.forEach(m => { label(m.m); number(m.v); });
    d.monthlyNet.forEach(m => { label(m.m); number(m.tahsilat); number(m.net); });
    d.aylikDagilim.forEach(m => { label(m.ay); if (!Array.isArray(m.kalemler)) throw Error('schema:monthly'); m.kalemler.forEach(([k,v])=>{label(k);number(v);}); });
    ['toplam','estaDirekt','estaEndirekt','taseronDirekt','taseronEndirekt','make','oran'].forEach(k=>number(d.personel[k]));
    ['cmpGun','pnrGun','gecenPnr'].forEach(k=>number(d.kalanGun[k]));
    ['cmpTarih','pnrTarih'].forEach(k=>label(d.kalanGun[k]));
    ['kdvHaric','kdvli'].forEach(k=>{amount(d.hakedisTotals[k].finansal);amount(d.hakedisTotals[k].brut);});
    ['kdvsiz','kdvli'].forEach(k=>amount(d.potansiyelDusulecek[k]));
    amount(d.brutHakedisToplam); amount(d.butce);
    if (d.harcamalar) {
      const h=d.harcamalar;
      if (!Array.isArray(h.aylar)||!Array.isArray(h.kategoriler)||!Array.isArray(h.toplam?.degerler)||h.aylar.length!==h.toplam.degerler.length) throw Error('schema:harcamalar');
      for (const k of h.kategoriler) if (!Array.isArray(k.aylik)||k.aylik.length!==h.aylar.length||(k.altKalemler||[]).some(a=>!Array.isArray(a.aylik)||a.aylik.length!==h.aylar.length)) throw Error('schema:series');
      h.aylar.forEach(label); h.toplam.degerler.forEach(number);
      h.kategoriler.forEach(k=>{label(k.ad);k.aylik.forEach(number);(k.altKalemler||[]).forEach(a=>{label(a.ad);a.aylik.forEach(number);});});
    }
    const finite = v => { if(typeof v==='number'&&!Number.isFinite(v))throw Error('schema:number'); if(v&&typeof v==='object')Object.values(v).forEach(finite); }; finite(d);
    return d;
  };
  return {parseM,ratio,sum,width,id,monthKey,gaugeStats,monthlyCategories,generalSeries,validate};
})();
