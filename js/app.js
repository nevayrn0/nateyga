(function(){
  const DATA_FILE = 'data/results.xlsx'; // ضع ملف النتيجة هنا بهذا الاسم بالضبط

  const statusCard = document.getElementById('status-card');
  const searchCard = document.getElementById('search-card');
  const searchInput = document.getElementById('search-input');
  const resultsEl = document.getElementById('results');

  let records = [];
  let columns = [];
  let nameKey = null;
  let seatKey = null;
  let statusKey = null;
  let totalKey = null;
  let debounceTimer = null;

  // ---------- helpers ----------
  function normalizeArabic(s){
    if(s===undefined||s===null) return '';
    s = s.toString();
    s = s.replace(/[\u064B-\u0652\u0670\u0640]/g,'');
    s = s.replace(/[إأآا]/g,'ا');
    s = s.replace(/ى/g,'ي');
    s = s.replace(/ة/g,'ه');
    s = s.replace(/ؤ/g,'و');
    s = s.replace(/ئ/g,'ي');
    s = s.replace(/\s+/g,' ').trim().toLowerCase();
    return s;
  }
  function normalizeDigits(s){
    if(s===undefined||s===null) return '';
    const map = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'};
    return s.toString().replace(/[٠-٩]/g, d=>map[d]).trim();
  }
  function isPureDigits(s){
    const d = normalizeDigits(s).replace(/\s/g,'');
    return d.length>0 && /^\d+$/.test(d);
  }
  function findColumn(keywords){
    return columns.find(c => {
      const nc = normalizeArabic(c);
      return keywords.some(k => nc.includes(normalizeArabic(k)));
    }) || null;
  }
  function escapeHtml(s){
    if(s===undefined||s===null) return '';
    return s.toString()
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;');
  }

  // ---------- load data from the site itself ----------
  loadData();

  function loadData(){
    fetch(DATA_FILE)
      .then(res=>{
        if(!res.ok) throw new Error('not-found');
        return res.arrayBuffer();
      })
      .then(buf=>{
        const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {defval:'', raw:false});
        if(!json.length) throw new Error('empty');

        records = json;
        columns = Object.keys(json[0]);
        nameKey = findColumn(['اسم الطالب','الاسم رباعي','الاسم كامل','الاسم','name']);
        seatKey = findColumn(['رقم الجلوس','رقم جلوس','الجلوس','seat','رقم الطالب']);
        statusKey = findColumn(['الحاله','الحالة','النتيجه','النتيجة','result','status']);
        totalKey = findColumn(['المجموع','مجموع الدرجات','النسبه','النسبة','total','percentage']);

        statusCard.innerHTML = `
          <div class="status-ready">
            <svg viewBox="0 0 24 24" fill="none" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
            <span>البيانات جاهزة — ${records.length} طالب</span>
          </div>`;
        searchCard.style.display = 'block';
        searchInput.disabled = false;
        searchInput.focus();
      })
      .catch(()=>{
        statusCard.innerHTML = `
          <div class="status-error">
            تعذّر تحميل بيانات النتيجة. تأكد من وجود ملف <code>${DATA_FILE}</code> داخل مجلد الموقع.
          </div>`;
      });
  }

  // ---------- search ----------
  searchInput.addEventListener('input', ()=>{
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(runSearch, 150);
  });

  function runSearch(){
    const q = searchInput.value.trim();
    if(!q){ resultsEl.innerHTML=''; return; }

    let matches = [];
    if(isPureDigits(q) && seatKey){
      const qd = normalizeDigits(q).replace(/\s/g,'');
      matches = records.filter(r => normalizeDigits(r[seatKey]).includes(qd));
    } else if(nameKey){
      const qn = normalizeArabic(q);
      const words = qn.split(' ').filter(Boolean);
      matches = records.filter(r=>{
        const rn = normalizeArabic(r[nameKey]);
        return words.every(w => rn.includes(w));
      });
    }

    if(matches.length===0){
      resultsEl.innerHTML = `
        <div class="state-msg error">
          <svg viewBox="0 0 24 24" fill="none" stroke-width="1.6"><circle cx="12" cy="12" r="9"/><path d="M9 9l6 6M15 9l-6 6"/></svg>
          <div>لا توجد نتيجة مطابقة. تأكد من كتابة الاسم رباعيًا كما هو مسجل، أو من صحة رقم الجلوس.</div>
        </div>`;
    } else if(matches.length===1){
      resultsEl.innerHTML = '';
      resultsEl.appendChild(renderCertificate(matches[0]));
    } else if(matches.length<=40){
      renderCandidateList(matches);
    } else {
      resultsEl.innerHTML = `
        <div class="state-msg">عدد النتائج كبير جدًا (${matches.length}). أضف كلمة أخرى من الاسم لتضييق البحث.</div>`;
    }
  }

  function renderCandidateList(matches){
    let html = `<ul class="candidates">`;
    matches.slice(0,40).forEach((r,i)=>{
      html += `<li><button data-idx="${i}">
        <span class="cname">${escapeHtml(nameKey ? r[nameKey] : '—')}</span>
        <span class="cseat">${seatKey ? escapeHtml(r[seatKey]) : ''}</span>
      </button></li>`;
    });
    html += `</ul>`;
    resultsEl.innerHTML = html;
    resultsEl.querySelectorAll('button[data-idx]').forEach(btn=>{
      btn.addEventListener('click', ()=>{
        const rec = matches[parseInt(btn.getAttribute('data-idx'),10)];
        resultsEl.innerHTML = '';
        resultsEl.appendChild(renderCertificate(rec));
      });
    });
  }

  function classifyStatus(val){
    const nv = normalizeArabic(val);
    if(!nv) return 'neutral';
    if(nv.includes('ناجح') || nv.includes('pass')) return 'pass';
    if(nv.includes('راسب') || nv.includes('دور') || nv.includes('fail')) return 'fail';
    return 'neutral';
  }

  function renderCertificate(record){
    const wrap = document.createElement('div');
    wrap.className = 'cert';

    const statusVal = statusKey ? record[statusKey] : '';
    const cls = classifyStatus(statusVal);
    const sealIcon = cls==='fail'
      ? '<path d="M18 6L6 18M6 6l12 12"/>'
      : '<path d="M20 6L9 17l-5-5"/>';

    let topHtml = '';
    if(statusKey || totalKey){
      topHtml += '<div class="cert-top">';
      if(totalKey){
        topHtml += `<div class="stat-block"><div class="lbl">${escapeHtml(totalKey)}</div><div class="val">${escapeHtml(record[totalKey])}</div></div>`;
      }
      if(statusKey){
        topHtml += `<div class="stat-block"><div class="lbl">${escapeHtml(statusKey)}</div><div class="val ${cls}">${escapeHtml(record[statusKey])}</div></div>`;
      }
      topHtml += '</div>';
    }

    let gridHtml = '<div class="field-grid">';
    columns.forEach(col=>{
      if(col===statusKey || col===totalKey) return;
      gridHtml += `<div class="field"><div class="lbl">${escapeHtml(col)}</div><div class="val">${escapeHtml(record[col]!==''?record[col]:'—')}</div></div>`;
    });
    gridHtml += '</div>';

    wrap.innerHTML = `
      <div class="seal ${cls}"><svg viewBox="0 0 24 24" fill="none" stroke-width="2.4">${sealIcon}</svg></div>
      ${topHtml}
      ${gridHtml}
    `;
    return wrap;
  }
})();
