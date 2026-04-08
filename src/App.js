import { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

const SK = 'akash_hr_v4';
const COLS = ['sl','name','age','address','university','degree','jobMatch','phone','email','careerSummary','bdJobsPage'];
const LBL = {
  sl:'#', name:'Name', age:'Age', address:'Address',
  university:'University', degree:'Degree / Program',
  jobMatch:'Job Match', phone:'Phone(s)', email:'Email',
  careerSummary:'Career Summary', bdJobsPage:'Bdjobs Page'
};

// ─── PARSER ───────────────────────────────────────────────────────────────────

function parseNameCell(raw) {
  if (!raw) return null;
  const text = String(raw).trim();
  if (!text) return null;

  const firstLine = text.split('\n')[0].trim();
  const skipPatterns = [
    /applicant summary/i, /job title/i, /downloaded on/i,
    /page no/i, /^sl$/i, /^image$/i, /^name$/i, /^career summary/i,
    /^experience/i, /^applied on/i, /^remarks/i
  ];
  if (skipPatterns.some(p => p.test(firstLine))) return null;
  if (/^\d+$/.test(firstLine)) return null;
  if (firstLine.length < 2) return null;

  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const name = lines[0].replace(/^name[:\s]*/i, '').trim();
  if (!name || name.length < 2) return null;

  const r = {
    name, age: '', address: '',
    university: '', degree: '',
    jobMatch: '', phone: '', email: ''
  };

  const phones = [];
  const addressParts = [];
  const universityParts = [];
  const degreeParts = [];

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];

    // ── Age ──────────────────────────────────────────────────
    if (/^age\s*[:\-]?\s*/i.test(line)) {
      const ageVal = line.replace(/^age\s*[:\-]?\s*/i, '').replace(/\s*years?\.?$/i, '').trim();
      if (ageVal && /\d/.test(ageVal)) r.age = ageVal;
      continue;
    }

    // ── Job Match ────────────────────────────────────────────
    if (/job\s*matching?\s*[:\-]?\s*\d/i.test(line)) {
      const m = line.match(/(\d+\.?\d*)\s*%?/);
      r.jobMatch = m ? m[1] + '%' : line.replace(/.*?[:\-]/, '').trim();
      continue;
    }

    // ── Email ────────────────────────────────────────────────
    const emailMatch = line.match(/[\w.+\-]+@[\w.\-]+\.\w{2,}/);
    if (emailMatch) { r.email = emailMatch[0]; continue; }

    // ── Phones ───────────────────────────────────────────────
    const cleanLine = line.replace(/[\s\-\.]/g, '');
    const phoneMatches = cleanLine.match(/(?:\+?880)?01[3-9]\d{8}/g);
    if (phoneMatches) {
      phoneMatches.forEach(p => {
        let norm = p.replace(/^\+?880/, '');
        if (!norm.startsWith('0')) norm = '0' + norm;
        if (!phones.includes(norm)) phones.push(norm);
      });
      continue;
    }

    // ── Degree ───────────────────────────────────────────────
    const degreeKw = /bachelor|master|phd|b\.?sc|b\.?b\.?a|m\.?b\.?a|b\.?eng|m\.?sc|b\.?a\b|m\.?a\b|diploma|hons|honours|llb|mbbs|b\.?arch|b\.?tech|m\.?tech|mphil|d\.?pharm/i;
    if (degreeKw.test(line)) {
      degreeParts.push(line); continue;
    }

    // ── University ───────────────────────────────────────────
    const uniKw = /university|college|institute of|polytechnic|academy|buet|nsu|diu|brac|ulab|aust|iub|aiub|sau|ciu|puc|ruet|cuet|kuet|ju\b|bu\b|sust|hstu|pstu|mist|national university|open university|islamic university/i;
    if (uniKw.test(line)) {
      universityParts.push(line); continue;
    }

    // ── Address ──────────────────────────────────────────────
    const addrKw = /dhaka|chittagong|ctg|chattogram|sylhet|rajshahi|khulna|barishal|barisal|mymensingh|rangpur|comilla|narayanganj|gazipur|tangail|bogura|bogra|jessore|jashore|faridpur|manikganj|narsingdi|habiganj|moulvibazar|sunamganj|brahmanbaria|chandpur|lakshmipur|noakhali|feni|cox|bandarban|rangamati|khagrachhari|kurigram|lalmonirhat|nilphamari|panchagarh|thakurgaon|dinajpur|joypurhat|naogaon|natore|sirajganj|sirajdikhan|pabna|meherpur|chuadanga|jhenaidah|magura|narail|satkhira|bagerhat|pirojpur|jhalokathi|bhola|patuakhali|barguna|shariatpur|madaripur|gopalganj|munshiganj|mohammadpur|mirpur|gulshan|banani|uttara|dhanmondi|motijheel|rampura|badda|khilgaon|khilgoan|lalbagh|hazaribagh|farmgate|tejgaon|tongi|savar|keraniganj|demra|jatrabari|mugda|malibagh|bashundhara|baridhara|nikunja|kafrul|pallabi|diabari|turag|cantonment|wari|shyampur|gendaria|sabujbagh|paltan|kamrangirchar|uttar khan|halishahar|north|south|east|west|housing|sadar|road|lane|avenue|district|thana|upazila|village|para|bazar|nagar|gram|sector|block|flat|apt|floor|house|plot|holding/i;
    if (addrKw.test(line)) {
      addressParts.push(line); continue;
    }

    // ── Fallback: unclassified line 2 treated as address ─────
    if (
      i === 1
      && !addressParts.length
      && line.length < 120
      && !/^\d+$/.test(line)
      && !/\d{7,}/.test(line)
    ) {
      addressParts.push(line);
    }
  }

  r.phone      = phones.join(', ');
  r.address    = addressParts.join(', ');
  r.university = universityParts.join(' ');
  r.degree     = degreeParts.join(' ');

  return r;
}

function parseCareersCell(raw) {
  if (!raw) return '';
  return String(raw).trim().replace(/\n/g, ' | ');
}

// ─── BDJOBS PAGE CALCULATOR ───────────────────────────────────────────────────
// Bdjobs shows 50 applicants per page
// SL 1-50 = Page 1, 51-100 = Page 2, etc.
function getBdjobsPage(slNumber) {
  return Math.ceil(slNumber / 50);
}

// ─── COLUMN & ROW DETECTION ───────────────────────────────────────────────────

function findDataStart(json) {
  for (let i = 0; i < Math.min(json.length, 25); i++) {
    const row = json[i] || [];
    let hasSL = false, hasName = false;
    row.forEach(cell => {
      const v = String(cell || '').trim();
      if (/^sl\.?$/i.test(v)) hasSL = true;
      if (/^name$/i.test(v)) hasName = true;
    });
    if (hasSL && hasName) return { headerIdx: i, dataStart: i + 1 };
  }
  for (let i = 0; i < Math.min(json.length, 25); i++) {
    const row = json[i] || [];
    const hasMultiline = row.some(cell => {
      const v = String(cell || '').trim();
      return v.includes('\n') && /^[A-Za-z]/.test(v) && v.length > 10;
    });
    if (hasMultiline) return { headerIdx: -1, dataStart: i };
  }
  return { headerIdx: -1, dataStart: 6 };
}

function findColumns(json, headerIdx, dataStart) {
  let nameCol = -1, careerCol = -1;

  if (headerIdx >= 0) {
    const headerRow = json[headerIdx] || [];
    headerRow.forEach((cell, i) => {
      const v = String(cell || '').trim().toLowerCase();
      if (v === 'name') nameCol = i;
      if (v.includes('career') || v.includes('summary')) careerCol = i;
    });
  }

  if (nameCol === -1) {
    const colScores = {};
    for (let r = dataStart; r < Math.min(json.length, dataStart + 15); r++) {
      const row = json[r] || [];
      row.forEach((cell, c) => {
        const v = String(cell || '').trim();
        if (v.includes('\n') && /^[A-Z]/.test(v) && /age\s*:/i.test(v)) {
          colScores[c] = (colScores[c] || 0) + 3;
        } else if (v.includes('\n') && /^[A-Z]/.test(v) && v.length > 15) {
          colScores[c] = (colScores[c] || 0) + 1;
        }
      });
    }
    const best = Object.entries(colScores).sort((a, b) => b[1] - a[1])[0];
    if (best) nameCol = parseInt(best[0]);
  }

  if (nameCol === -1) nameCol = 2;

  if (careerCol === -1) {
    const colScores = {};
    for (let r = dataStart; r < Math.min(json.length, dataStart + 15); r++) {
      const row = json[r] || [];
      for (let c = nameCol + 1; c < row.length; c++) {
        const v = String(row[c] || '').trim();
        if (v.length > 15) {
          colScores[c] = (colScores[c] || 0) + 1;
        }
      }
    }
    const best = Object.entries(colScores).sort((a, b) => b[1] - a[1])[0];
    if (best) careerCol = parseInt(best[0]);
  }

  if (careerCol === -1) careerCol = nameCol + 1;

  return { nameCol, careerCol };
}

// ─── BADGE ────────────────────────────────────────────────────────────────────

function MatchBadge({ val }) {
  const n = parseFloat(val);
  if (!val) return <span className="empty-val">—</span>;
  const cls = n >= 80 ? 'badge-high' : n >= 60 ? 'badge-med' : 'badge-low';
  return <span className={`badge ${cls}`}>{val}</span>;
}

function PageBadge({ page }) {
  if (!page) return <span className="empty-val">—</span>;
  return <span className="page-badge">Page {page}</span>;
}

// ─── APP ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [rows, setRows]             = useState([]);
  const [search, setSearch]         = useState('');
  const [uniFilter, setUni]         = useState('All');
  const [ageMin, setAgeMin]         = useState('');
  const [ageMax, setAgeMax]         = useState('');
  const [jobMin, setJobMin]         = useState('');
  const [summaryFilter, setSummary] = useState('All');
  const [pageFilter, setPageFilter] = useState('All');
  const [error, setError]           = useState('');
  const [info, setInfo]             = useState('');
  const fileRef = useRef();

  useEffect(() => {
    try {
      const saved = localStorage.getItem(SK);
      if (saved) setRows(JSON.parse(saved));
    } catch(e) {}
  }, []);

  // University dropdown
  const universities = useMemo(() => {
    const set = new Set();
    rows.forEach(r => { if (r.university) set.add(r.university.trim()); });
    return ['All', ...Array.from(set).sort()];
  }, [rows]);

  // Page dropdown — unique pages from data
  const pageOptions = useMemo(() => {
    const set = new Set();
    rows.forEach(r => { if (r.bdJobsPage) set.add(r.bdJobsPage); });
    return ['All', ...Array.from(set).sort((a, b) => a - b)];
  }, [rows]);

  // Total pages count
  const totalPages = useMemo(() => {
    if (!rows.length) return 0;
    return Math.ceil(rows.length / 50);
  }, [rows]);

  // Apply all filters
  const filtered = useMemo(() => {
    return rows.filter(r => {
      const q = search.toLowerCase();
      if (q && !Object.values(r).some(v => v && String(v).toLowerCase().includes(q))) return false;
      if (uniFilter !== 'All' && r.university !== uniFilter) return false;
      const age = parseFloat(r.age);
      if (ageMin && !isNaN(age) && age < parseFloat(ageMin)) return false;
      if (ageMax && !isNaN(age) && age > parseFloat(ageMax)) return false;
      const jm = parseFloat(r.jobMatch);
      if (jobMin && !isNaN(jm) && jm < parseFloat(jobMin)) return false;
      if (summaryFilter === 'With Summary'    && !r.careerSummary) return false;
      if (summaryFilter === 'Without Summary' &&  r.careerSummary) return false;
      if (pageFilter !== 'All' && String(r.bdJobsPage) !== String(pageFilter)) return false;
      return true;
    });
  }, [rows, search, uniFilter, ageMin, ageMax, jobMin, summaryFilter, pageFilter]);

  const handleFile = (file) => {
    if (!file) return;
    setError(''); setInfo('Reading file...');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: 'array', cellText: false, raw: false,
          cellDates: true, sheetStubs: true
        });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, {
          header: 1, raw: false, defval: '', blankrows: false
        });

        const { headerIdx, dataStart } = findDataStart(json);
        const { nameCol, careerCol }   = findColumns(json, headerIdx, dataStart);

        console.log('=== Akash HR Debug ===');
        console.log(`Header row index : ${headerIdx}`);
        console.log(`Data starts at   : row ${dataStart + 1}`);
        console.log(`Name column      : ${nameCol}`);
        console.log(`Career column    : ${careerCol}`);

        const parsed = [];
        let sl = 1;
        for (let r = dataStart; r < json.length; r++) {
          const row = json[r];
          if (!row) continue;
          const nameRaw = row[nameCol];
          if (!nameRaw || !String(nameRaw).trim()) continue;
          const res = parseNameCell(nameRaw);
          if (res) {
            res.sl           = sl;
            res.careerSummary = parseCareersCell(row[careerCol]);
            res.bdJobsPage   = getBdjobsPage(sl); // ← auto page number
            sl++;
            parsed.push(res);
          }
        }

        if (!parsed.length) {
          setError(`No valid records found. Name col: ${nameCol}, Career col: ${careerCol}, data from row ${dataStart + 1}.`);
          setInfo(''); return;
        }

        setInfo(`Loaded ${parsed.length} applicants across ${Math.ceil(parsed.length / 50)} Bdjobs pages.`);
        setTimeout(() => setInfo(''), 4000);
        setRows(parsed);
        localStorage.setItem(SK, JSON.stringify(parsed));
      } catch(ex) {
        setError('Error reading file: ' + ex.message);
        setInfo('');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportExcel = () => {
    const exportCols = ['sl','name','age','address','university','degree','jobMatch','phone','email','careerSummary','bdJobsPage'];
    const data = [exportCols.map(c => LBL[c])];
    filtered.forEach(r => data.push(exportCols.map(c => r[c] || '')));
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [5, 22, 6, 30, 34, 28, 12, 22, 26, 38, 12].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Applicants');
    XLSX.writeFile(wb, 'akash_hr_applicants.xlsx');
  };

  const clearAll = () => {
    if (!window.confirm('Clear all loaded data?')) return;
    setRows([]); setSearch(''); setUni('All');
    setAgeMin(''); setAgeMax(''); setJobMin('');
    setSummary('All'); setPageFilter('All');
    localStorage.removeItem(SK);
  };

  const resetFilters = () => {
    setSearch(''); setUni('All');
    setAgeMin(''); setAgeMax(''); setJobMin('');
    setSummary('All'); setPageFilter('All');
  };

  const avgMatch = () => {
    const nums = rows.map(r => parseFloat(r.jobMatch)).filter(n => !isNaN(n));
    return nums.length
      ? Math.round(nums.reduce((a, b) => a + b, 0) / nums.length) + '%'
      : '–';
  };

  return (
    <div className="app-wrapper">
      <div className="app">

        {/* ── Header ── */}
        <div className="header">
          <div className="logo-row">
            <div className="logo-icon">A</div>
            <div>
              <h1>Akash HR</h1>
              <p>Applicant Contact Parser — Bdjobs Excel Importer</p>
            </div>
          </div>
        </div>

        {/* ── Drop Zone ── */}
        <div
          className="drop-zone"
          onClick={() => fileRef.current.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }}
        >
          <div className="drop-icon">📂</div>
          <p><strong>Drop your .xlsx file here</strong> or click to browse</p>
          <p className="sub">
            Reads Name &amp; Career Summary — auto-calculates Bdjobs page number (50 per page)
          </p>
          <input
            ref={fileRef} type="file" accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={(e) => { handleFile(e.target.files[0]); e.target.value = ''; }}
          />
        </div>

        {info  && <div className="msg info">{info}</div>}
        {error && <div className="msg error">{error}</div>}

        {rows.length > 0 && (
          <>
            {/* ── Stats ── */}
            <div className="stats">
              <div className="stat">
                <div className="stat-label">Total Applicants</div>
                <div className="stat-val">{rows.length}</div>
              </div>
              <div className="stat">
                <div className="stat-label">Showing</div>
                <div className="stat-val">{filtered.length}</div>
              </div>
              <div className="stat">
                <div className="stat-label">Bdjobs Pages</div>
                <div className="stat-val">{totalPages}</div>
              </div>
              <div className="stat">
                <div className="stat-label">With Phone</div>
                <div className="stat-val">{rows.filter(r => r.phone).length}</div>
              </div>
              <div className="stat">
                <div className="stat-label">With Email</div>
                <div className="stat-val">{rows.filter(r => r.email).length}</div>
              </div>
              <div className="stat">
                <div className="stat-label">Avg Job Match</div>
                <div className="stat-val">{avgMatch()}</div>
              </div>
            </div>

            {/* ── Page Info Banner ── */}
            <div className="page-banner">
              <span className="page-banner-icon">📄</span>
              <div>
                <strong>Bdjobs Page Reference</strong>
                <span>
                  {Array.from({ length: totalPages }, (_, i) => (
                    <span key={i} className="page-pill">
                      Page {i + 1}: SL {i * 50 + 1}–{Math.min((i + 1) * 50, rows.length)}
                    </span>
                  ))}
                </span>
              </div>
            </div>

            {/* ── Filters ── */}
            <div className="filter-box">
              <div className="filter-row">
                <div className="filter-group full">
                  <label>Search</label>
                  <input
                    className="finput" type="text"
                    placeholder="Search name, phone, email, address, university..."
                    value={search}
                    onChange={e => setSearch(e.target.value)}
                  />
                </div>
              </div>

              <div className="filter-row">
                <div className="filter-group">
                  <label>University</label>
                  <select className="finput" value={uniFilter} onChange={e => setUni(e.target.value)}>
                    {universities.map(u => <option key={u} value={u}>{u}</option>)}
                  </select>
                </div>

                <div className="filter-group">
                  <label>Age Range</label>
                  <div className="range-row">
                    <input
                      className="finput small" type="number"
                      placeholder="Min" value={ageMin}
                      onChange={e => setAgeMin(e.target.value)}
                      min="0" max="100"
                    />
                    <span className="range-sep">–</span>
                    <input
                      className="finput small" type="number"
                      placeholder="Max" value={ageMax}
                      onChange={e => setAgeMax(e.target.value)}
                      min="0" max="100"
                    />
                  </div>
                </div>

                <div className="filter-group">
                  <label>Min Job Match %</label>
                  <input
                    className="finput" type="number"
                    placeholder="e.g. 70" value={jobMin}
                    onChange={e => setJobMin(e.target.value)}
                    min="0" max="100"
                  />
                </div>

                <div className="filter-group">
                  <label>Career Summary</label>
                  <select className="finput" value={summaryFilter} onChange={e => setSummary(e.target.value)}>
                    <option>All</option>
                    <option>With Summary</option>
                    <option>Without Summary</option>
                  </select>
                </div>

                <div className="filter-group">
                  <label>Bdjobs Page</label>
                  <select className="finput" value={pageFilter} onChange={e => setPageFilter(e.target.value)}>
                    {pageOptions.map(p => (
                      <option key={p} value={p}>
                        {p === 'All' ? 'All Pages' : `Page ${p} (SL ${(p-1)*50+1}–${p*50})`}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="filter-actions">
                <span className="result-count">
                  Showing <strong>{filtered.length}</strong> of <strong>{rows.length}</strong> applicants
                </span>
                <div className="action-btns">
                  <button className="btn btn-ghost"   onClick={resetFilters}>Reset Filters</button>
                  <button className="btn btn-primary" onClick={exportExcel}>Export Excel ↓</button>
                  <button className="btn btn-danger"  onClick={clearAll}>Clear All</button>
                </div>
              </div>
            </div>

            {/* ── Table ── */}
            {filtered.length === 0 ? (
              <div className="empty-state">No applicants match the current filters.</div>
            ) : (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>{COLS.map(c => <th key={c}>{LBL[c]}</th>)}</tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, i) => (
                      <tr key={i}>
                        {COLS.map(c => (
                          <td
                            key={c}
                            className={
                              c === 'email'         ? 'email-cell'   :
                              c === 'sl'            ? 'sl-cell'      :
                              c === 'careerSummary' ? 'summary-cell' :
                              c === 'bdJobsPage'    ? 'page-cell'    :
                              !r[c]                 ? 'empty'        : ''
                            }
                          >
                            {c === 'jobMatch'   ? <MatchBadge val={r[c]} /> :
                             c === 'bdJobsPage' ? <PageBadge page={r[c]} /> :
                             r[c] || '—'}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </div>

      {/* ── Footer ── */}
      <footer className="footer">
        Developed by <strong>Nafiz Ahmed Rhythm</strong>
      </footer>
    </div>
  );
}