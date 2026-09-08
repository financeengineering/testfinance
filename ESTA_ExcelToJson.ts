// ============================================================================
// ESTA · 110 SPPC — Excel (MASTER sayfası) → data.json dönüştürücü
// ============================================================================
// KULLANIM:
//  1) Excel Online'da dosyanı aç (OneDrive for Business üzerinde).
//  2) Üst menüden "Automate" sekmesine gir.
//  3) "New Script" de, açılan editördeki örnek kodu SİL, bu dosyanın
//     tamamını yapıştır, kaydet (adını "ExcelToJson" yap).
//  4) Power Automate akışındaki "Run script" adımında bu script'i seçeceğiz.
//
// Bu script MASTER sayfasındaki bölüm başlıklarını ve etiketleri ARAYARAK
// çalışır (sabit satır numarasına güvenmez) — yani ileride satır eklesen/
// çıkarsan script bozulmaz, otomatik yeni konumu bulur.
// ============================================================================

function main(workbook: ExcelScript.Workbook): string {
  // MASTER sayfasını konumuna/ismine göre değil, İÇERİĞİNE göre bul —
  // böylece sayfa adını değiştirsen ya da başka sayfalar eklesen bile
  // (hangi sırada olursa olsun) her zaman doğru sayfa okunur.
  const allSheets = workbook.getWorksheets();
  let sheet: ExcelScript.Worksheet | undefined;
  for (let i = 0; i < allSheets.length; i++) {
    const v = allSheets[i].getRange("A1").getValue();
    if (typeof v === "string" && v.indexOf("ESTA") === 0 && v.indexOf("MASTER") > -1) {
      sheet = allSheets[i];
      break;
    }
  }
  if (!sheet) throw new Error("ESTA MASTER sayfası bulunamadı. A1 başlığını kontrol edin.");
  const used = sheet.getUsedRange();
  if (!used) throw new Error("MASTER sayfası boş.");
  // Include A1 even if the used range starts further down/right.
  const range = sheet.getRangeByIndexes(0, 0, used.getRowIndex() + used.getRowCount(), used.getColumnIndex() + used.getColumnCount());
  const values = range.getValues() as (string | number | boolean)[][];
  const nRows = values.length;

  // -------------------- YARDIMCI FONKSİYONLAR --------------------

  function findSectionRow(titleStartsWith: string, fromRow: number = 0): number {
    for (let i = fromRow; i < nRows; i++) {
      const v = values[i][0];
      if (typeof v === "string" && v.indexOf(titleStartsWith) === 0) return i;
    }
    throw new Error("Zorunlu bölüm bulunamadı: " + titleStartsWith);
  }

  function findLabelRow(startRow: number, endRow: number, label: string): number {
    const end = endRow === -1 ? nRows : endRow;
    for (let i = startRow; i < end; i++) {
      const v = values[i][1];
      if (v === label) return i;
    }
    throw new Error("Zorunlu etiket bulunamadı: " + label);
  }

  function num(r: number, c: number): number {
    if (r < 0 || r >= nRows) return 0;
    const v = values[r][c];
    if (typeof v === "number") return v;
    if (typeof v === "string") {
      // Bazı hücreler biçimlendirilmiş metin (boşluk/₽/virgül içeren) olarak
      // gelebiliyor — bunları da sayıya çevirmeyi dene.
      const cleaned = v.replace(/[−–]/g, "-").replace(/[₽\s]/g, "");
      const n = parseFloat(cleaned.replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", "."));
      if (cleaned === "") return 0;
      if (!Number.isFinite(n)) throw new Error("Sayısal değer okunamadı: satır " + (r + 1) + ", sütun " + (c + 1));
      return n;
    }
    return 0;
  }

  function str(r: number, c: number): string {
    if (r < 0 || r >= nRows) return "";
    const v = values[r][c];
    if (v === null || v === undefined) return "";
    return String(v);
  }

  function trGroup(n: number): string {
    const neg = n < 0;
    const abs = Math.round(Math.abs(n));
    let s = String(abs);
    let out = "";
    let cnt = 0;
    for (let i = s.length - 1; i >= 0; i--) {
      out = s[i] + out;
      cnt++;
      if (cnt % 3 === 0 && i !== 0) out = "." + out;
    }
    return (neg ? "− " : "") + out;
  }

  function mFromM(v: number): string {
    return trGroup(v) + " M";
  }

  function mFromRuble(v: number): string {
    return trGroup(v / 1000000) + " M";
  }

  function dateStr(r: number, c: number): string {
    const v = values[r][c];
    // Excel Online / Office Scripts, tarih hücrelerini Date nesnesi olarak
    // DEĞİL, seri numarası (gün say\u0131s\u0131) olarak d\u00f6nd\u00fcr\u00fcyor. Bunu manuel \u00e7evir.
    if (typeof v === "number" && v > 20000) {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      const dd = String(d.getUTCDate()).padStart(2, "0");
      const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
      const yyyy = d.getUTCFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }
    if (Object.prototype.toString.call(v) === "[object Date]") {
      const d = v as unknown as Date;
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }
    return str(r, c);
  }

  const AY_MAP: { [k: number]: string } = {
    0: "Oca", 1: "Şub", 2: "Mar", 3: "Nis", 4: "May", 5: "Haz",
    6: "Tem", 7: "Ağu", 8: "Eyl", 9: "Eki", 10: "Kas", 11: "Ara"
  };

  function ayLabel(r: number, c: number): string {
    const v = values[r][c];
    if (typeof v === "number" && v > 20000) {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return `${AY_MAP[d.getUTCMonth()]}.${String(d.getUTCFullYear()).slice(2)}`;
    }
    if (Object.prototype.toString.call(v) === "[object Date]") {
      const d = v as unknown as Date;
      return `${AY_MAP[d.getMonth()]}.${String(d.getFullYear()).slice(2)}`;
    }
    return str(r, c);
  }

  // Bir bölümdeki tüm "değer satırlarını" (B sütunu dolu olan) döndürür.
  function sectionRows(sectionTitle: string, nextSectionTitle: string, headerOffset: number = 1): number[] {
    const start = findSectionRow(sectionTitle);
    const end = nextSectionTitle ? findSectionRow(nextSectionTitle, start + 1) : nRows;
    const rows: number[] = [];
    for (let i = start + 1 + headerOffset; i < (end === -1 ? nRows : end); i++) {
      if (values[i][1] !== null && values[i][1] !== undefined && values[i][1] !== "") {
        rows.push(i);
      }
    }
    return rows;
  }

  // -------------------- META --------------------
  const secMeta = findSectionRow("META");
  const rTarih = findLabelRow(secMeta, findSectionRow("GÜNCEL SÖZLEŞME"), "Rapor Tarihi (gg.aa.yyyy)");
  const meta = {
    project: "110 SPPC",
    date: dateStr(rTarih, 2),
    generatedAt: new Date().toISOString(),
    schemaVersion: 1,
    currency: "RUB",
    amountUnit: "million",
    client: "OOO «Эста Констракшен» · Severstal"
  };

  // -------------------- SÖZLEŞME (Güncel + Potansiyel) --------------------
  function sozlesmeBlok(sectionTitle: string, nextTitle: string, toplamLabel: string, label: string, dateStr2: string) {
    const start = findSectionRow(sectionTitle);
    const end = findSectionRow(nextTitle, start + 1);
    const rToplam = findLabelRow(start, end, toplamLabel);
    const rIscilik = findLabelRow(start, end, "İşçilik (M)");
    const rTest = findLabelRow(start, end, "İşçilik - Test ve Devreye Alma (M)");
    const rVinc = findLabelRow(start, end, "100 Ton Üstü Vinç Maliyetleri (M)");
    const rMob = findLabelRow(start, end, "Mobilizasyon (M)");
    const rMalzeme = findLabelRow(start, end, "Malzeme (M)");
    return {
      label: label,
      date: dateStr2,
      total: mFromM(num(rToplam, 2)),
      items: [
        ["İşçilik", mFromM(num(rIscilik, 2))],
        ["İşçilik - Test ve Devreye Alma", mFromM(num(rTest, 2))],
        ["100 Ton Üstü Vinç Maliyetleri", mFromM(num(rVinc, 2))],
        ["Mobilizasyon (Ek Sözleşme №6)", mFromM(num(rMob, 2))],
        ["Malzeme", mFromM(num(rMalzeme, 2))]
      ]
    };
  }

  const contractCurrent = sozlesmeBlok(
    "GÜNCEL SÖZLEŞME", "POTANSİYEL SÖZLEŞME",
    "Güncel Sözleşme Bedeli Toplam (M)",
    "Güncel Sözleşme Bedeli — Доп. № 20", "24.07.2024"
  );
  const contractPotential = (() => {
    const b = sozlesmeBlok(
      "POTANSİYEL SÖZLEŞME", "FİNANSAL GÖSTERGELER",
      "Potansiyel Sözleşme Bedeli Toplam (M)",
      "Toplam Potansiyel Sözleşme Bedeli", ""
    );
    return { label: b.label, total: b.total, items: b.items };
  })();

  // -------------------- FİNANSAL GÖSTERGELER (ham M) --------------------
  const secGost = findSectionRow("FİNANSAL GÖSTERGELER");
  const secSurec = findSectionRow("SÜREÇ TAKVİMİ");
  function gVal(label: string): number {
    return num(findLabelRow(secGost, secSurec, label), 2);
  }
  const gToplamTahsilat = gVal("Toplam Tahsilat (M)");
  const gMalzemeKesintisi = gVal("Malzeme Kesintisi (M)");
  const gGarantiOdemesi = gVal("Garanti Ödemesi (M)");
  const gSahaTahakkuk = gVal("Saha Tahakkuk Tutarı (M)");
  const gToplamHarcama = gVal("Toplam Harcama (M)");
  const gProjeTransfer = gVal("Proje Transfer (M)");
  const gButce = gVal("Bütçe (M)");
  const gBankaBakiye = gVal("Banka Bakiye - Borç Hariç (M)");
  const gAcikIsler = gVal("Açık İşler (M)");
  const gBorc = gVal("Borç (M)");

  // -------------------- SÜREÇ TAKVİMİ --------------------
  const secBrutDetay = findSectionRow("BRÜT HAKEDİŞ DETAYI");
  function sVal(label: string): number { return num(findLabelRow(secSurec, secBrutDetay, label), 2); }
  function sDate(label: string): string { return dateStr(findLabelRow(secSurec, secBrutDetay, label), 2); }
  const kalanGun = {
    cmpGun: Math.round(sVal("Kalan Gün (CMP)")),
    cmpTarih: sDate("CMP Tarih (gg.aa.yyyy)"),
    pnrGun: Math.round(sVal("Kalan Gün (ПНР)")),
    pnrTarih: sDate("ПНР Tarih (gg.aa.yyyy)"),
    gecenPnr: Math.round(sVal("Geçen Gün Oranı % (ПНР)") * 100)
  };

  // -------------------- BRÜT HAKEDİŞ DETAYI --------------------
  const secKesinti = findSectionRow("KESİNTİLER");
  function bVal(label: string): number { return num(findLabelRow(secBrutDetay, secKesinti, label), 2); }
  const brutHakedisDetay = [
    ["İşçilik", mFromM(bVal("İşçilik (M)"))],
    ["Malzeme", mFromM(bVal("Malzeme (M)"))],
    ["100 Ton Üstü Vinçler", mFromM(bVal("100 Ton Üstü Vinçler (M)"))],
    ["Kur Farkı", mFromM(bVal("Kur Farkı (M)"))]
  ];
  const brutHakedisToplamNum = bVal("Brüt Hakediş Toplam (M)");
  const brutHakedisToplam = mFromM(brutHakedisToplamNum);

  // -------------------- KESİNTİLER --------------------
  const secAvans = findSectionRow("AVANS TABLOSU");
  function kVal(label: string): number { return num(findLabelRow(secKesinti, secAvans, label), 2); }
  const garantiKesinti10 = kVal("Garanti Kesintisi %10 (Ödenen Hariç)");
  const garantiKesinti5 = kVal("Garanti Kesintisi %5");
  const kesintiler: (string | number | null)[][] = [
    ["Avans Kesintileri", mFromM(kVal("Avans Kesintileri (M)")), null],
    ["Banka Garantisiz Kesinti", mFromM(kVal("Banka Garantisiz Avans Kesintisi (M)")), null],
    ["Banka Garantili Kesinti", mFromM(kVal("Banka Garantili Avans Kesintisi (M)")), null],
    ["İSG Kesintileri", mFromM(kVal("İSG Kesintileri (M)")), null],
    ["Elektrik Kesintileri", mFromM(kVal("Elektrik Kesintileri (M)")), null],
    ["Garanti Kesintileri", mFromM(garantiKesinti10 + garantiKesinti5), null],
    ["Garanti Kesintisi %10 (Ödenen Hariç)", mFromM(garantiKesinti10), null],
    ["Garanti Kesintisi %5", mFromM(garantiKesinti5), null]
  ];

  // -------------------- AVANS TABLOSU --------------------
  const secPersonel = findSectionRow("PERSONEL");
  function avansRow(label: string): [string, string] {
    const r = findLabelRow(secAvans, secPersonel, label);
    // NOT: Avans Tablosu artık ham ₽ değil, doğrudan M cinsinden geliyor
    // (Sözleşme bölümünden farklı) — bu yüzden mFromM kullanıyoruz.
    return [mFromM(num(r, 2)), mFromM(num(r, 3))];
  }
  const avansSummary: (string | number | null)[][] = [
    ["Toplam Avans Tahsilat", ...avansRow("Toplam Avans Tahsilat"), null],
    ["Toplam Kesilecek Avans", ...avansRow("Toplam Kesilecek Avans"), null],
    ["Toplam Avans Kesintisi", ...avansRow("Toplam Avans Kesintisi"), null],
    ["Kalan Avans Borcu", ...avansRow("Kalan Avans Borcu"), null]
  ];
  const bankaGarantili: (string | number | null)[][] = [
    ["Alınan Banka Garantili Avans", ...avansRow("Alınan Banka Garantili Avans"), null],
    ["Kesilen Banka Garantili Avans Mektubu", ...avansRow("Kesilen Banka Garantili Avans Mektubu"), null],
    ["Kalan Banka Garantili Avans Mektubu", ...avansRow("Kalan Banka Garantili Avans Mektubu"), null]
  ];
  const bankaGarantisiz: (string | number | null)[][] = [
    ["Alınan Banka Garantisiz Avans", ...avansRow("Alınan Banka Garantisiz Avans"), null],
    ["Kesilen Banka Garantisiz Avans", ...avansRow("Kesilen Banka Garantisiz Avans"), null],
    ["Kalan Banka Garantisiz Avans", ...avansRow("Kalan Banka Garantisiz Avans"), null]
  ];
  const potDusRow = avansRow("Potansiyel Düşülecek Banka Mektubu");
  const potansiyelDusulecek = { kdvsiz: potDusRow[0], kdvli: potDusRow[1] };

  // -------------------- PERSONEL --------------------
  const secHakedisTipi = findSectionRow("HAKEDİŞ TİPİ DAĞILIMI");
  function pVal(label: string): number { return num(findLabelRow(secPersonel, secHakedisTipi, label), 2); }
  const pToplam = Math.round(pVal("Toplam Personel"));
  const pEstaDirekt = Math.round(pVal("ESTA Direkt Personel"));
  const pEstaEndirekt = Math.round(pVal("ESTA Endirekt Personel"));
  const pEstaToplam = pEstaDirekt + pEstaEndirekt;
  const personel = {
    toplam: pToplam,
    estaDirekt: pEstaDirekt,
    estaEndirekt: pEstaEndirekt,
    taseronDirekt: Math.round(pVal("Taşeron Direkt Personel")),
    taseronEndirekt: Math.round(pVal("Taşeron Endirekt Personel")),
    make: Math.round(pVal("Make")),
    oran: pEstaToplam > 0 ? Math.round((pEstaEndirekt / pEstaToplam) * 100) : 0
  };

  // -------------------- HAKEDİŞ TİPİ DAĞILIMI --------------------
  const secSaha = findSectionRow("SAHA İLERLEME");
  const hakedisRows = sectionRows("HAKEDİŞ TİPİ DAĞILIMI", "SAHA İLERLEME", 1).filter(r => {
    const label = str(r, 1);
    return label.indexOf("TOPLAM") !== 0;
  });
  const hakedisTipleri = hakedisRows.map(r => [str(r, 1), mFromM(num(r, 2)), mFromM(num(r, 3))]);

  const rToplamKdvHaric = findLabelRow(secHakedisTipi, secSaha, "TOPLAM (KDV hariç)");
  const rToplamKdvli = findLabelRow(secHakedisTipi, secSaha, "TOPLAM (KDV'li)");
  const hakedisTotals = {
    kdvHaric: { finansal: mFromM(num(rToplamKdvHaric, 2)), brut: mFromM(num(rToplamKdvHaric, 3)) },
    kdvli: { finansal: mFromM(num(rToplamKdvli, 2)), brut: mFromM(num(rToplamKdvli, 3)) }
  };

  // -------------------- SAHA İLERLEME --------------------
  const secAylikBrut = findSectionRow("AYLIK BRÜT HAKEDİŞ");
  const sahaHeaderRow = secSaha + 1;
  const sahaHeaders = values[sahaHeaderRow] as string[];
  function sahaCol(headerText: string): number {
    for (let c = 0; c < sahaHeaders.length; c++) {
      if (sahaHeaders[c] === headerText) return c;
    }
    throw new Error("Saha sütunu bulunamadı: " + headerText);
  }
  const cBirim = sahaCol("Birim");
  const cProje = sahaCol("Proje Metraj");
  const cGerc = sahaCol("Gerçekleşen (Toplam)");
  const cKalan = sahaCol("Kalan (Toplam)");
  const cAyPlan = sahaCol("Bu Ay Planlanan");
  const cAyGerc = sahaCol("Bu Ay Gerçekleşen");
  const cAyKalan = sahaCol("Bu Ay Kalan");
  const cHaftalik = sahaCol("Haftalık Ortalama");

  const sahaRows = sectionRows("SAHA İLERLEME", "AYLIK BRÜT HAKEDİŞ", 1);
  const sahaItems = sahaRows.map(r => ({
    name: str(r, 1),
    unit: str(r, cBirim),
    proje: Math.round(num(r, cProje)),
    gerc: Math.round(num(r, cGerc)),
    kalan: Math.round(num(r, cKalan)),
    ay: {
      plan: Math.round(num(r, cAyPlan)),
      gerc: Math.round(num(r, cAyGerc)),
      kalan: Math.round(num(r, cAyKalan))
    },
    haftalik: Math.round(num(r, cHaftalik))
  }));

  // -------------------- AYLIK BRÜT HAKEDİŞ --------------------
  const secAylikTahsilat = findSectionRow("AYLIK TAHSİLAT VE NET HAKEDİŞ");
  const monthlyBrutRows = sectionRows("AYLIK BRÜT HAKEDİŞ", "AYLIK TAHSİLAT VE NET HAKEDİŞ", 1);
  const monthlyBrut = monthlyBrutRows
    .filter(r => values[r][2] !== "" && values[r][2] !== null && values[r][2] !== undefined)
    .map(r => ({ m: ayLabel(r, 1), v: Math.round(num(r, 2)) }));

  // -------------------- AYLIK TAHSİLAT VE NET HAKEDİŞ --------------------
  const secAylikKalem = findSectionRow("AYLIK KALEM DAĞILIMI");
  const monthlyNetRows = sectionRows("AYLIK TAHSİLAT VE NET HAKEDİŞ", "AYLIK KALEM DAĞILIMI", 1);
  const monthlyNet = monthlyNetRows
    .filter(r => [2, 3].some(c => values[r][c] !== "" && values[r][c] !== null && values[r][c] !== undefined))
    .map(r => ({ m: ayLabel(r, 1), tahsilat: Math.round(num(r, 2)), net: Math.round(num(r, 3)) }));

  // -------------------- AYLIK KALEM DAĞILIMI --------------------
  const secHarcamalarAna = findSectionRow("HARCAMALAR — ANA KALEMLER");
  const kalemHeaderRow = secAylikKalem + 1;
  const kalemHeaders = values[kalemHeaderRow] as string[];
  const kalemRows = sectionRows("AYLIK KALEM DAĞILIMI", "HARCAMALAR — ANA KALEMLER", 1);
  const aylikDagilim = kalemRows
    .filter(r => {
      // en az bir kalemde veri var mı?
      for (let c = 2; c < kalemHeaders.length; c++) {
        const v = values[r][c];
        if (typeof v === "number") return true;
      }
      return false;
    })
    .map(r => {
      const kalemler: [string, number][] = [];
      let toplam = 0;
      for (let c = 2; c < kalemHeaders.length; c++) {
        const h = kalemHeaders[c];
        if (!h) continue;
        const v = num(r, c);
        if (values[r][c] !== "" && values[r][c] !== null && values[r][c] !== undefined) {
          kalemler.push([h, Math.round(v)]);
          toplam += v;
        }
      }
      return { ay: ayLabel(r, 1), toplam: Math.round(toplam), kalemler };
    });

  // -------------------- HARCAMALAR --------------------
  const secHarcamalarTaseron = findSectionRow("HARCAMALAR — TAŞERON ALT KALEMLERİ");
  const anaHeaderRow = secHarcamalarAna + 1;
  const anaHeaders = values[anaHeaderRow] as string[];
  const anaRows = sectionRows("HARCAMALAR — ANA KALEMLER", "HARCAMALAR — TAŞERON ALT KALEMLERİ", 1);
  const harcamalarAylar = anaRows.map(r => ayLabel(r, 1));

  const taseronHeaderRow = secHarcamalarTaseron + 1;
  const taseronHeaders = values[taseronHeaderRow] as string[];
  const taseronRows = sectionRows("HARCAMALAR — TAŞERON ALT KALEMLERİ", "", 1);

  function seriesFor(rows: number[], col: number): number[] {
    return rows.map(r => Math.round(num(r, col)));
  }

  const harcamalarKategoriler = [] as {
    ad: string; odenen: number; aylik: number[]; altKalemler: { ad: string; odenen: number; aylik: number[] }[];
  }[];

  for (let c = 2; c < anaHeaders.length; c++) {
    const ad = anaHeaders[c];
    if (!ad) continue;
    const aylik = seriesFor(anaRows, c);
    const odenen = Math.round(aylik.reduce((a, b) => a + b, 0));
    const altKalemler: { ad: string; odenen: number; aylik: number[] }[] = [];
    if (ad === "Taşeron") {
      const taseronMonthRows = new Map<string, number>();
      taseronRows.forEach(r => {
        const month = ayLabel(r, 1);
        if (taseronMonthRows.has(month)) throw new Error("Tekrarlanan taşeron ayı: " + month);
        taseronMonthRows.set(month, r);
      });
      for (let tc = 2; tc < taseronHeaders.length; tc++) {
        const tAd = taseronHeaders[tc];
        if (!tAd) continue;
        const tAylik = harcamalarAylar.map(month => {
          const row = taseronMonthRows.get(month);
          if (row === undefined) throw new Error("Taşeron tablosunda ay eksik: " + month);
          return Math.round(num(row, tc));
        });
        const tOdenen = Math.round(tAylik.reduce((a, b) => a + b, 0));
        altKalemler.push({ ad: tAd, odenen: tOdenen, aylik: tAylik });
      }
    }
    harcamalarKategoriler.push({ ad, odenen, aylik, altKalemler });
  }

  const harcamalarToplamDegerler = harcamalarAylar.map((_, i) =>
    Math.round(harcamalarKategoriler.reduce((sum, k) => sum + (k.aylik[i] || 0), 0))
  );

  // Yalnızca boş gelecek ayları kırp; açıkça girilmiş sıfırları koru.
  let lastRealIdx = harcamalarToplamDegerler.length - 1;
  while (lastRealIdx >= 0 && anaHeaders.every((header, c) =>
    c < 2 || !header || values[anaRows[lastRealIdx]][c] === "" || values[anaRows[lastRealIdx]][c] === null || values[anaRows[lastRealIdx]][c] === undefined
  )) lastRealIdx--;
  const harcamalarAylarTrim = harcamalarAylar.slice(0, lastRealIdx + 1);
  const harcamalarToplamDegerlerTrim = harcamalarToplamDegerler.slice(0, lastRealIdx + 1);
  harcamalarKategoriler.forEach(k => {
    k.aylik = k.aylik.slice(0, lastRealIdx + 1);
    k.altKalemler.forEach(ak => { ak.aylik = ak.aylik.slice(0, lastRealIdx + 1); });
  });

  const harcamalarToplamOdenen = Math.round(harcamalarToplamDegerlerTrim.reduce((a, b) => a + b, 0));

  const harcamalar = {
    aylar: harcamalarAylarTrim,
    kategoriler: harcamalarKategoriler,
    toplam: { ay: harcamalarAylarTrim, degerler: harcamalarToplamDegerlerTrim, odenen: harcamalarToplamOdenen }
  };

  // -------------------- GÖSTERGELER (gauges) --------------------
  const gauges = [
    {
      big: mFromM(gToplamTahsilat),
      sub: "",
      label: "Toplam Tahsilat",
      accent: "gold"
    },
    { big: mFromM(gSahaTahakkuk), sub: "", label: "Saha Tahakkuk", accent: "gold" },
    { big: mFromM(gToplamHarcama), sub: `Proje Transfer ${mFromM(gProjeTransfer)}`, label: "Toplam Harcama", accent: "red" },
    { big: mFromM(kVal("Avans Kesintileri (M)")), sub: `Avans Tahsilat ${avansRow("Toplam Avans Tahsilat")[0]}`, label: "Avans Kesintisi", accent: "red" },
    { big: mFromM(brutHakedisToplamNum), sub: "", label: "Brüt Hakediş", accent: "gold" },
    { big: mFromM(gAcikIsler), sub: `Banka Bakiye ${mFromM(gBankaBakiye)}`, label: "Açık İşler", accent: "orange" },
    { big: mFromM(gBorc), sub: "", label: "Borç", accent: "red" },
    { big: mFromM(gBankaBakiye), sub: "", label: "Banka Bakiye", accent: "gold" }
  ];

  const bankaBakiye = {
    bakiye: mFromM(gBankaBakiye),
    harcama: mFromM(gToplamHarcama),
    acikIsler: mFromM(gAcikIsler),
    pct: gSahaTahakkuk > 0 ? Math.round((gAcikIsler / gSahaTahakkuk) * 100) : 0
  };

  const butce = mFromM(gButce);

  // -------------------- SONUÇ --------------------
  const result = {
    meta,
    contractCurrent,
    contractPotential,
    gauges,
    brutHakedisDetay,
    kalanGun,
    bankaBakiye,
    avansSummary,
    bankaGarantili,
    bankaGarantisiz,
    potansiyelDusulecek,
    monthlyBrut,
    monthlyNet,
    hakedisTipleri,
    hakedisTotals,
    personel,
    sahaItems,
    brutHakedisToplam,
    kesintiler,
    aylikDagilim,
    harcamalar,
    butce
  };

  return JSON.stringify(result);
}
