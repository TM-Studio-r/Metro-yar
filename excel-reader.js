/* MetroYar Excel Reader - browser-side parser for .xls/.xlsx */
(function(){
  const SHEETJS_CDNS = [
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
  ];
  let scriptPromise;
  const workbookCache = new Map();
  const lineCache = new Map();

  function loadScript(src){return new Promise((resolve,reject)=>{const s=document.createElement('script');s.src=src;s.async=true;s.onload=()=>resolve();s.onerror=()=>reject(new Error('SheetJS load failed'));document.head.appendChild(s);});}
  async function ensureXLSX(){
    if(window.XLSX) return window.XLSX;
    if(!scriptPromise){scriptPromise=(async()=>{let last;for(const url of SHEETJS_CDNS){try{await loadScript(url);if(window.XLSX)return window.XLSX;}catch(e){last=e;}}throw last||new Error('SheetJS unavailable');})();}
    return scriptPromise;
  }

  function norm(s){
    return String(s??'').replace(/\u200c|\u200f|\u200e/g,'').replace(/[يى]/g,'ی').replace(/ك/g,'ک').replace(/[ۀة]/g,'ه').replace(/ـ/g,'').replace(/[أإ]/g,'ا').replace(/\(\s*(?:ره|عج|س)\s*\)/g,'').replace(/[()،,؛؛:._\-]/g,' ').replace(/[\u064B-\u065F]/g,'').replace(/\s+/g,'').trim();
  }
  const manualRaw={
    'ميدانوليعصر':'میدان حضرت ولی عصر','میدانولیعصر':'میدان حضرت ولی عصر',
    'ميدانحضرتولي عصر':'میدان حضرت ولی عصر','میدانولیعصر':'میدان حضرت ولی عصر','ميدانوليعصر':'میدان حضرت ولی عصر','هروي':'میدان هروی',
    'شهيدآرمان':'شهید آرمان علی وردی (کوهسار)','شهیدآرمان':'شهید آرمان علی وردی (کوهسار)',
    'دولتآباد':'شهدای دولت آباد','شهدایدولتآباد':'شهدای دولت آباد',
    'شاهد':'شاهد-باقرشهر','شاهدباقرشهر':'شاهد-باقرشهر',
    'ارمسبز':'ارم سبز','دانشگاهتربیتمدرس':'دانشگاه تربیت مدرس','تربیتمدرس':'دانشگاه تربیت مدرس',
    'تربيتمدرس':'دانشگاه تربیت مدرس','تربیت مدرس':'دانشگاه تربیت مدرس','تربيت مدرس':'دانشگاه تربیت مدرس','بهشتي':'شهید بهشتی','شهیدبهشتی':'شهید بهشتی',
    'آیتالهکاشانی':'آیت الله کاشانی','آیتاللهکاشانی':'آیت الله کاشانی',
    'شهداي17شهريور':'شهدای 17 شهریور','شهدای17شهریور':'شهدای 17 شهریور',
    'شهداي17شهریور':'شهدای 17 شهریور','شهدايهفدهشهریور':'شهدای 17 شهریور',
    'قیام':'میدان قیام','ميدانقيام':'میدان قیام','میدانقیام':'میدان قیام',
    'كتاب':'میدان کتاب','کتاب':'میدان کتاب','تختی':'ورزشگاه تختی','بسیج':'بسیج',
    'هلالاحمر':'هلال احمر','بریانک':'بریانک','کمیل':'کمیل','مولوی':'مولوی','رودکی':'رودکی',
    'برجمیلاد':'برج میلاد تهران','میدانصنعت':'میدان صنعت','شهیددادمان':'شهید دادمان',
    'مدافعانسلامت':'مدافعان سلامت','بوستانگفتگو':'بوستان گفتگو','توحید':'توحید',
    'شهدایکن':'شهدای کن','شهرزیبا':'شهر زیبا','شهران':'شهران','مرزداران':'مرزداران',
    'شهرکآزمایش':'شهرک آزمایش','یادگارامام':'یادگار امام','اشرفیاصفهانی':'شهید اشرفی اصفهانی',
    'شهیدستاری':'شهید ستاری','کارگر':'کارگر','بوستانلاله':'بوستان لاله','بهارشیراز':'بهار شیراز',
    'امامحسین':'امام حسین','میدانشهدا':'میدان شهدا','امیرکبیر':'امیرکبیر','میدانخراسان':'میدان خراسان',
    'شهیدرضایی':'شهید رضایی','کیانشهر':'کیان شهر','بعثت':'بعثت',
    'شهیدنوابصفوی':'شهید نواب صفوی','دانشگاهامامعلی':'دانشگاه امام علی','دانشگاهشریف':'دانشگاه شریف',
    'حسنآباد':'حسن آباد','ملت':'ملت','بهارستان':'بهارستان','دروازهشمیران':'دروازه شمیران','شهیدمدنی':'شهید مدنی','فدک':'فدک','جانبازان':'جانبازان','سرسبز':'سرسبز','تهرانپارس':'تهرانپارس','فرهنگسرا':'فرهنگسرا',
    'مصلیامامخمینی':'مصلی امام خمینی','امامخمینی':'امام خمینی','حرممطهرامامخمینی':'حرم مطهر امام خمینی',
    'میدانمحمدیه':'میدان محمدیه','پایانهجنوب':'پایانه جنوب','شهیدبخارایی':'شهید بخارایی','علیآباد':'علی آباد','جوانمردقصاب':'جوانمرد قصاب','شهرری':'شهرری','پالایشگاه':'پالایشگاه','حرممطهرامامخمینی':'حرم مطهر امام خمینی','کهریزک':'کهریزک',
    'شهدايهفتتير':'شهدای هفتم تیر','شهیدبخارایی':'شهید بخارایی','شهيدبخارائي':'شهید بخارایی','خواجهعبداللهانصاری':'خواجه عبدالله انصاری','خواجهعبدالهانصاري':'خواجه عبدالله انصاری','زمزم':'زمزم','تجریش':'تجریش','قیطریه':'قیطریه','شهیدصدر':'شهید صدر','قلهک':'قلهک','دکترشریعتی':'دکتر شریعتی','میرداماد':'میرداماد','شهیدحقانی':'شهید حقانی','شهیدهمت':'شهید همت','شهیدمفتح':'شهید مفتح','شهدایهفتمتیر':'شهدای هفتم تیر','طالقانی':'آیت الله طالقانی','سعدی':'سعدی','پانزدهخرداد':'پانزده خرداد','خیام':'خیام','شوش':'شوش','شهرآفتاب':'نمایشگاه شهر آفتاب','نمایشگاهشهرآفتاب':'نمایشگاه شهر آفتاب','فرودگاهامامخمینی':'فرودگاه امام خمینی','شهرپرند':'شهر پرند','پرند':'شهر پرند',
    'شهیدکلاهدوز':'شهید کلاهدوز','نیروهوایی':'نیروهوایی','نيرو هوايي':'نیروهوایی','نبرد':'نبرد','پیروزی':'پیروزی','ابنسینا':'ابن سینا','میدانآزادی':'میدان آزادی','دکترحبیباله':'دکتر حبیب الله','استادمعین':'استاد معین','شهرکاکباتان':'شهرک اکباتان','علامهجعفری':'علامه جعفری',
    'تهرانصادقیه':'تهران (صادقیه)','صادقیه':'تهران (صادقیه)','تهران':'تهران (صادقیه)','طرشت':'طرشت','دانشگاهعلموصنعت':'دانشگاه علم و صنعت',
    'آزادگان':'آزادگان','نعمتآباد':'نعمت آباد','عبدلآباد':'عبدل آباد','شهرکشریعتی':'شهرک شریعتی','زمزم':'زمزم','جوادیه':'جوادیه','راهآهن':'راه آهن','منیریه':'منیریه','میدانحضرتولیعصر':'میدان حضرت ولی عصر','میدانجهاد':'میدان جهاد','میرزایشیرازی':'میرزای شیرازی','سهروردی':'سهروردی','شهیدقدوسی':'شهید قدوسی','شهیدصیادشیرازی':'شهید صیاد شیرازی','خواجهعبداللهانصاری':'خواجه عبدالله انصاری','شهیدزینالدین':'شهید زین الدین','حسینآباد':'حسین آباد','نوبنیاد':'نوبنیاد','اقدسیه':'اقدسیه','شهیدمحلاتی':'شهید محلاتی','قائم':'قائم',
    'گلشهر':'گلشهر','کرج':'کرج','محمدشهر':'محمدشهر','گرمدره':'گرم دره','اتمسفر':'اتمسفر','وردآورد':'وردآورد','ایرانخودرو':'ایران خودرو','چیتگر':'چیتگر','ورزشگاهآزادی':'ورزشگاه آزادی','شهیدفخریزاده':'شهید فخری زاده','شهیدسپهبدقاسمسلیمانی':'شهید سپهبد قاسم سلیمانی','هشتگرد':'هشتگرد','ماموت':'ماموت',
    'بیمه':'بیمه','پایانه12فرودگاه مهرآباد':'پایانه 1 و 2 فرودگاه مهرآباد','پایانه46فرودگاه مهرآباد':'پایانه 4 و 6 فرودگاه مهرآباد','پایانه12فرودگاه مهرآّباد':'پایانه 1 و 2 فرودگاه مهرآباد',
    'شهدای هفت تیر':'شهدای هفتم تیر','شهداي هفت تير':'شهدای هفتم تیر','شهید بخارایی':'شهید بخارایی','شهيد بخارائي':'شهید بخارایی','خواجه عبداله انصاري':'خواجه عبدالله انصاری',
  };

  const manual=Object.fromEntries(Object.entries(manualRaw).map(([k,v])=>[norm(k),v]));

  function canonicalStation(cell, allStations){
    if(cell==null || typeof cell==='object') return null;
    const n=norm(cell);
    if(!n) return null;
    if(n.startsWith('اعزاماز')) return canonicalStation(n.slice(7),allStations);
    if(n.startsWith('دریافت')) return canonicalStation(n.slice(6),allStations);
    const control=['سرفاصله','هدوی','ریال','ريال','بسمت','از','به','اعزام','دریافت','دريافت'];
    if(control.includes(n)||n.startsWith('روز')||n.startsWith('جمعه')||n.startsWith('شنبه')||n.startsWith('پنجشنبه')||n.startsWith('تعطیل')) return null;
    if(manual[n]) return manual[n];
    for(const s of allStations){if(norm(s)===n)return s;}
    for(const s of allStations){const sn=norm(s);if(sn.length>=6 && (n.includes(sn)||sn.includes(n)))return s;}
    return null;
  }

  function classifyDay(title){
    const t=norm(title);
    if(t.includes('پنجشنبه') || t.includes('پنجشنبهها') || t.includes('پنجشنبه')) return 'thursday';
    if(t.includes('جمعه') || t.includes('تعطیل')) return 'holiday';
    return 'normal';
  }
  function toMinutes(v){
    if(v instanceof Date) return v.getHours()*60+v.getMinutes();
    if(typeof v==='number' && Number.isFinite(v)){
      if(Math.abs(v)<1e-100) return null;
      if(v>=0 && v<2) return Math.round(v*1440) % 1440;
      if(v>=0 && v<=2359 && Math.floor(v/100)<=23 && Math.floor(v%100)<=59) return Math.floor(v/100)*60 + Math.floor(v%100);
    }
    if(typeof v==='string'){
      let s=v.trim().replace(/[۰-۹]/g,d=>String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)));
      const m=s.match(/(\d{1,2})\s*[:٫.]\s*(\d{2})/); if(m)return Number(m[1])*60+Number(m[2]);
      const n=Number(s); if(Number.isFinite(n)) return toMinutes(n);
    }
    return null;
  }

  function findHeader(rows, allStations){
    let best=null;
    for(let r=0;r<Math.min(rows.length,25);r++){
      const row=rows[r]||[]; let found=[]; for(let c=0;c<row.length;c++){const st=canonicalStation(row[c],allStations);if(st)found.push({c,st});}
      const uniq=[...new Set(found.map(x=>x.st))];
      if(uniq.length>=2 && (!best || uniq.length>best.uniq.length)) best={row:r,found,uniq};
    }
    if(!best) return null;
    // keep actual order and drop duplicated columns
    const columns=[]; const seen=new Set();
    for(const x of best.found){if(!seen.has(x.st)){seen.add(x.st);columns.push(x);}}
    return {...best,columns};
  }

  function parseWorkbook(workbook, allStations){
    const services=[];
    workbook.SheetNames.forEach(sheetName=>{
      const ws=workbook.Sheets[sheetName];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
      const header=findHeader(rows,allStations); if(!header) return;
      const day=classifyDay(sheetName+' '+rows.slice(0,header.row+1).flat().filter(x=>typeof x==='string').join(' '));
      const stationOrder=header.columns.map(x=>x.st);
      const colByStation=Object.fromEntries(header.columns.map(x=>[x.st,x.c]));
      const trips=[];
      for(let r=header.row+1;r<rows.length;r++){
        const row=rows[r]||[]; const times={}; let count=0;
        for(const st of stationOrder){const mins=toMinutes(row[colByStation[st]]); if(mins!=null){times[st]=mins;count++;}}
        if(count>=2) trips.push({times});
      }
      if(trips.length) services.push({sheetName,day,stationOrder,trips});
    });
    return services;
  }

  async function loadWorkbook(url,allStations){
    if(workbookCache.has(url))return workbookCache.get(url);
    const X=await ensureXLSX();
    const res=await fetch(url,{cache:'no-cache'}); if(!res.ok)throw new Error('Excel fetch failed: '+url+' ['+res.status+']');
    const ab=await res.arrayBuffer(); const wb=X.read(ab,{type:'array',cellDates:false,dense:true});
    const services=parseWorkbook(wb,allStations); if(!services.length)throw new Error('هیچ جدول زمان‌بندی قابل تشخیصی در '+url+' پیدا نشد.');
    workbookCache.set(url,services); return services;
  }

  function selectDay(services,day){
    let exact=services.filter(s=>s.day===day); if(exact.length)return exact;
    if(day==='thursday'){let fallback=services.filter(s=>s.day==='normal');if(fallback.length)return fallback;}
    if(day==='holiday'){let fallback=services.filter(s=>s.day==='normal');if(fallback.length)return fallback;}
    return services;
  }
  async function loadSources(urls,allStations,day){
    const all=[]; for(const url of urls){const ss=await loadWorkbook(url,allStations);all.push(...selectDay(ss,day));} return all;
  }

  window.MetroExcel={loadSources,ensureXLSX,toMinutes,norm};
})();
