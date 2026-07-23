
>     <script>
          /* â”€â”€ ESTADO GLOBAL â”€â”€ */
          const fileUpload = document.getElementById('file-upload');
          const blocksContainer = document.getElementById('blocks-container');
          const sheetContentContainer = document.getElementById('sheet-content-container');
          const reportContainer = document.getElementById('report-view');
          const graphsContainer = document.getElementById('graphs-view');
          const loadingEl = document.getElementById('loading');
          const fileInfoEl = document.getElementById('file-info');
          const fileNameEl = document.getElementById('file-name');
          const clearFileBtn = document.getElementById('clear-file-btn');
          const monthSelect = document.getElementById('month-select');
          const repoStatusBadge = document.getElementById('repo-status-badge');
          const repoStatusMsg = document.getElementById('repo-load-status-msg');
  
          let activeCharts = [];
          let workbookData = null;
  
          if (window.ChartAnnotation) Chart.register(window.ChartAnnotation);
  
          /* â”€â”€ NAVEGACIÃ“N DE PESTAÃ‘AS â”€â”€ */
          document.querySelectorAll('.tab-btn').forEach(btn => {
              btn.addEventListener('click', () => {
                  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
                  btn.classList.add('active');
                  document.getElementById(btn.dataset.tab).classList.add('active');
                  if (btn.dataset.tab === 'report-view' && workbookData) generateWeeklyReport();
                  if (btn.dataset.tab === 'graphs-view' && workbookData) generateGraphsReport();
              });
          });
  
          /* â”€â”€ EVENTOS DE CARGA â”€â”€ */
          fileUpload.addEventListener('change', handleFile);
          clearFileBtn.addEventListener('click', () => clearFile(true));
          monthSelect.addEventListener('change', e => { if (e.target.value) loadExcelByMonth(e.target.value); });
  
          /* â”€â”€ INDEXEDDB â”€â”€ */
          const dbName = "CamasPilotoDB", storeName = "excelFiles";
          function getDB() {
              return new Promise((resolve, reject) => {
                  const req = indexedDB.open(dbName, 1);
                  req.onupgradeneeded = e => { const db = e.target.result; if 
(!db.objectStoreNames.contains(storeName)) db.createObjectStore(storeName); };
                  req.onsuccess = e => resolve(e.target.result);
                  req.onerror = e => reject(e.target.error);
              });
          }
          function saveFileToDB(filename, arrayBuffer) {
              return getDB().then(db => new Promise((resolve, reject) => {
                  const tx = db.transaction(storeName, "readwrite");
                  const store = tx.objectStore(storeName);
                  const req = store.put({ filename, data: arrayBuffer }, "activeFile");
                  req.onsuccess = () => resolve();
                  req.onerror = e => reject(e.target.error);
              })).catch(err => console.warn("IndexedDB save error:", err));
          }
          function getFileFromDB() {
              return getDB().then(db => new Promise((resolve, reject) => {
                  const tx = db.transaction(storeName, "readonly");
                  const req = tx.objectStore(storeName).get("activeFile");
                  req.onsuccess = e => resolve(e.target.result);
                  req.onerror = e => reject(e.target.error);
              })).catch(() => null);
          }
          function deleteFileFromDB() {
              return getDB().then(db => new Promise((resolve, reject) => {
                  const tx = db.transaction(storeName, "readwrite");
                  const req = tx.objectStore(storeName).delete("activeFile");
                  req.onsuccess = () => resolve();
                  req.onerror = e => reject(e.target.error);
              })).catch(() => {});
          }
  
          /* â”€â”€ CARGA DESDE EL SERVIDOR (GITHUB PAGES) â”€â”€ */
          async function loadExcelByMonth(month) {
              repoStatusBadge.textContent = 'Cargandoâ€¦';
              repoStatusBadge.classList.remove('loaded');
              if(repoStatusMsg) repoStatusMsg.textContent = `Buscando archivo de ${month}â€¦`;
              clearFile(false);
              loadingEl.style.display = 'flex';
  
              const tryFetch = async (ext) => {
                  const fileName = `${month}${ext}`;
                  const r = await fetch(`./${fileName}?t=${new Date().getTime()}`); // Cache buster
                  if (!r.ok) throw new Error(`HTTP ${r.status}`);
                  return { buf: await r.arrayBuffer(), fileName };
              };
  
              try {
                  let result;
                  try {
                      result = await tryFetch('.xlsx');
                  } catch (e) {
                      try {
                          result = await tryFetch('.xlsm');
                      } catch (e2) {
                          result = await tryFetch('.xls');
                      }
                  }
                  
                  processAndDisplayWorkbook(result.buf, result.fileName);
                  saveFileToDB(result.fileName, result.buf);
                  repoStatusBadge.textContent = 'âœ“ Cargado';
                  repoStatusBadge.classList.add('loaded');
                  if(repoStatusMsg) repoStatusMsg.textContent = `Archivo ${result.fileName} listo.`;
              } catch (err) {
                  repoStatusBadge.textContent = 'Error';
                  if(repoStatusMsg) repoStatusMsg.textContent = `No se encontrÃ³ el archivo ${month} (.xlsx/.xlsm) en 
GitHub. Verifica haberlo subido.`;
              } finally {
                  loadingEl.style.display = 'none';
              }
          }
  
          /* â”€â”€ MANEJO DE ARCHIVOS LOCALES â”€â”€ */
          function handleFile(e) {
              const file = e.target.files[0];
              if (!file) return;
              clearFile(false);
              loadingEl.style.display = 'flex';
              const reader = new FileReader();
              reader.onload = ev => {
                  try {
                      const buf = ev.target.result;
                      processAndDisplayWorkbook(buf, file.name);
                      saveFileToDB(file.name, buf);
                  } catch (err) {
                      console.error(err);
                      showInlineError(sheetContentContainer, 'Error al procesar el archivo. AsegÃºrate de que es un 
Excel vÃ¡lido.');
                  } finally {
                      loadingEl.style.display = 'none';
                  }
              };
              reader.onerror = () => { loadingEl.style.display = 'none'; showInlineError(sheetContentContainer, 'Error 
al leer el archivo.'); };
              reader.readAsArrayBuffer(file);
          }
  
          function clearFile(deleteFromDB = true) {
              workbookData = null;
              blocksContainer.innerHTML = '';
              sheetContentContainer.innerHTML = '';
              activeCharts.forEach(c => c.destroy());
              activeCharts = [];
              fileInfoEl.style.display = 'none';
              fileUpload.value = '';
              reportContainer.innerHTML = `<div class="empty-state"><p>Carga un archivo Excel para generar el reporte 
semanal.</p></div>`;
              graphsContainer.innerHTML = `<div class="empty-state"><p>Carga un archivo Excel para ver las grÃ¡ficas 
comparativas.</p></div>`;
              if (deleteFromDB) deleteFileFromDB();
          }
  
          function showInlineError(container, msg) {
              container.innerHTML = `<div class="inline-alert err" style="margin:16px;">${msg}</div>`;
          }
  
          /* â”€â”€ HELPERS DE DATOS â”€â”€ */
          function getLimitsForValve(v) {
              if (!v) return { ce: null, no3: null };
              v = parseInt(v);
              if ([1,2,3,4,5,6,7,8,9,10,11,21,22,23,24,25,26,27,28,29,30,31,32].includes(v)) return { ce: [2.4, 3], 
no3: [180, 220] };
              if ([12,13,14].includes(v)) return { ce: [1.8, 2.4], no3: [160, 200] };
              if ([15,16,17,36,37].includes(v)) return { ce: [1.7, 2.3], no3: [180, 220] };
              if ([18,47].includes(v)) return { ce: [1.8, 2.4], no3: [180, 220] };
              if ([19,52].includes(v)) return { ce: [1.8, 2.5], no3: [180, 220] };
              if ([20,51].includes(v)) return { ce: [1.8, 2.4], no3: [160, 200] };
              if ([33,34,35,50].includes(v)) return { ce: [2.4, 3], no3: [180, 220] };
              if ([38].includes(v)) return { ce: [1.7, 2.3], no3: [180, 220] };
              if ([39,40,41].includes(v)) return { ce: [1.7, 2.1], no3: [180, 220] };
              if ([42,43,44,53].includes(v)) return { ce: [1.8, 2.4], no3: [160, 200] };
              if ([45].includes(v)) return { ce: [1.8, 2.4], no3: [160, 200] };
              if ([46].includes(v)) return { ce: [1.8, 2.4], no3: [180, 220] };
              if ([48,49].includes(v)) return { ce: [2.4, 3], no3: [180, 220] };
              return { ce: null, no3: null };
          }
  
          function getVarietyForValve(v) {
              if (!v) return '';
              v = parseInt(v);
              if ([1,2,3,4,5,6,7,8,9,10,11,21,22,23,24,25,26,27,28,29,30,31,32].includes(v)) return 'Freedom';
              if ([12,13,14].includes(v)) return 'Alive';
              if ([15,16,17,36,37].includes(v)) return 'Amarela';
              if ([18,47].includes(v)) return 'Vainilla';
              if ([19,52].includes(v)) return 'Kendal';
              if ([20,51].includes(v)) return 'Ameli';
              if ([33,34,35,50].includes(v)) return 'Jacaranda';
              if ([38].includes(v)) return 'Texas';
              if ([39,40,41].includes(v)) return 'Blessing';
              if ([42,43,44,53].includes(v)) return 'Aperol';
              if ([45].includes(v)) return 'Suggar Doll';
              if ([46].includes(v)) return 'Dakar';
              if ([48,49].includes(v)) return 'Brave in Dark';
              return 'Desc.';
          }
  
          const parseNumericValue = v => {
              if (v === null || v === undefined || v === '') return NaN;
              if (typeof v === 'number') return v;
              let s = String(v).replace(/%/g, '').trim();
              if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
              return (isFinite(s) && s !== '') ? parseFloat(s) : NaN;
          };
          const cleanSumaText = v => typeof v === 'string' ? v.replace(/^[Ss][uÃº]ma\s+de\s+/i, 
'').replace(/^[Ss][uÃº]ma\s+/i, '') : v;
          const cleanSheetData = d => Array.isArray(d) ? d.map(r => Array.isArray(r) ? r.map(cleanSumaText) : r) : d;
  
          function getSheetJson(ws) {
              if (ws && ws['!ref']) {
                  try {
                      const range = XLSX.utils.decode_range(ws['!ref']);
                      range.e.c = Math.max(range.e.c, 150); // Forzar lectura de columnas adicionales
                      ws['!ref'] = XLSX.utils.encode_range(range);
                  } catch (e) {}
              }
              return cleanSheetData(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }));
          }
  
          function getBlockForValve(v) {
              const m = { 1:'Bloque 50',2:'Bloque 50',3:'Bloque 50',12:'Bloque 51',13:'Bloque 51',14:'Bloque 
51',4:'Bloque 52',5:'Bloque 52',6:'Bloque 52',7:'Bloque 52',8:'Bloque 52',15:'Bloque 53',16:'Bloque 53',17:'Bloque 
53',9:'Bloque 54',10:'Bloque 54',11:'Bloque 54',18:'Bloque 55',19:'Bloque 55',20:'Bloque 55',21:'Bloque 56',22:'Bloque 
56',23:'Bloque 56',33:'Bloque 57',34:'Bloque 57',35:'Bloque 57',24:'Bloque 58',25:'Bloque 58',26:'Bloque 
58',36:'Bloque 59',37:'Bloque 59',38:'Bloque 59',27:'Bloque 60',28:'Bloque 60',29:'Bloque 60',39:'Bloque 
61',40:'Bloque 61',41:'Bloque 61',30:'Bloque 62',31:'Bloque 62',32:'Bloque 62',42:'Bloque 63',43:'Bloque 
63',44:'Bloque 63',45:'Bloque 64',46:'Bloque 64',47:'Bloque 64',48:'Bloque 65',49:'Bloque 65',50:'Bloque 
65',51:'Bloque 66',52:'Bloque 66',53:'Bloque 66' };
              return m[v] ? [m[v]] : [];
          }
  
          function sortBlocks(arr) {
              return arr.sort((a, b) => {
                  if (a === 'Otros') return 1; if (b === 'Otros') return -1;
                  const rx = /Bloque (\d+)(?:\.(\d+))?([A-Z])?/;
                  const ma = a.match(rx), mb = b.match(rx);
                  if (ma && mb) {
                      const d = parseInt(ma[1]) - parseInt(mb[1]); if (d) return d;
                      const e = (ma[2]||0) - (mb[2]||0); if (e) return e;
                      return (ma[3]||'').localeCompare(mb[3]||'');
                  }
                  return a.localeCompare(b);
              });
          }
  
          /* â”€â”€ PROCESO Y VISUALIZACIÃ“N DEL WORKBOOK â”€â”€ */
          function processAndDisplayWorkbook(arrayBuffer, filename) {
              fileNameEl.textContent = filename;
              fileInfoEl.style.display = 'flex';
              try {
                  const data = new Uint8Array(arrayBuffer);
                  workbookData = XLSX.read(data, { type: 'array' });
                  const valvesByBlock = {};
                  workbookData.SheetNames.forEach(name => {
                      const m = name.match(/\((\d+)\)/);
                      if (!m) return;
                      const vn = parseInt(m[1]);
                      getBlockForValve(vn).forEach(bn => {
                          if (!valvesByBlock[bn]) valvesByBlock[bn] = [];
                          valvesByBlock[bn].push({ sheetName: name, valveNumber: vn });
                      });
                  });
                  blocksContainer.innerHTML = '';
                  sortBlocks(Object.keys(valvesByBlock)).forEach(bn => {
                      const card = document.createElement('div');
                      card.className = 'block-card';
                      const hdr = document.createElement('div');
                      hdr.className = 'block-header';
                      hdr.textContent = bn;
                      const body = document.createElement('div');
                      body.className = 'block-body';
                      valvesByBlock[bn].sort((a, b) => a.valveNumber - b.valveNumber).forEach(({ sheetName, 
valveNumber }) => {
                          const btn = document.createElement('button');
                          btn.className = 'valve-btn';
                          btn.textContent = `VÃ¡lvula ${valveNumber} (${getVarietyForValve(valveNumber)})`;
                          btn.dataset.sheetName = sheetName;
                          btn.addEventListener('click', () => {
                              document.querySelectorAll('.valve-btn').forEach(b => b.classList.remove('active'));
                              btn.classList.add('active');
                              displaySheetContent(sheetName);
                          });
                          body.appendChild(btn);
                      });
                      card.appendChild(hdr);
                      card.appendChild(body);
                      blocksContainer.appendChild(card);
                  });
                  const first = blocksContainer.querySelector('.valve-btn');
                  if (first) first.click();
                  
                  renderWeeklyCEChart();
              } catch (err) {
                  console.error(err);
                  showInlineError(sheetContentContainer, 'No se pudo leer el archivo. AsegÃºrate de que es un Excel 
vÃ¡lido.');
                  clearFile();
              }
          }
  
          /* â”€â”€ CONTENIDO DE HOJA â”€â”€ */
  /* â”€â”€ HELPER DE BÃšSQUEDA DE FILAS MEJORADO â”€â”€ */
  function findRowIndices(data) {
      const findRowIdx = (...searchTerms) => data.findIndex(row => {
          if (!row || !row[0]) return false;
          // Normaliza eliminando puntos, espacios mÃºltiples y convirtiendo a minÃºsculas
          const label = String(row[0]).toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ').trim();
          return searchTerms.some(term => {
              const cleanTerm = term.toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ').trim();
              return label.includes(cleanTerm);
          });
      });
  
      return {
          ceEsp: findRowIdx('ce esperada', 'ce esperado'),
          ceDre: findRowIdx('ce drenaje', 'ce emisor'),
          phEsp: findRowIdx('ph esperada', 'ph esperado'),
          phDre: findRowIdx('ph drenaje', 'ph emisor'),
          ejecucion: findRowIdx('% ejecucion', 'ejecucion'),
          no3: findRowIdx('n-no3', 'no3'),
          vol: findRowIdx('volumen ejecutado', 'volumen de riego'),
          
          // BÃºsqueda flexible que tolera "C.E. Drenaje", "C.E Drenaje", "CE Drenaje", etc.
          ceDrenajeReal: findRowIdx('ce drenaje', 'c e drenaje'),
          ceEmisorReal: findRowIdx('ce emisor', 'c e emisor'),
          phDrenajeReal: findRowIdx('ph drenaje'),
          phEmisorReal: findRowIdx('ph emisor'),
  
          pctDrenaje: findRowIdx('% drenaje', 'porcentaje de drenaje')
      };
  }
  
  /* â”€â”€ CÃLCULO DE SEMANA ISO ROBUSTO â”€â”€ */
  function getWeekFromDateStr(dateStr, monthSelected) {
      if (dateStr === null || dateStr === undefined) return null;
      dateStr = String(dateStr).trim();
      if (!dateStr) return null;
  
      let date = null;
      const currentYear = new Date().getFullYear();
  
      // 1. Si es un nÃºmero puro
      if (!isNaN(dateStr) && !dateStr.includes('/') && !dateStr.includes('-')) {
          const num = parseFloat(dateStr);
          if (num > 35000) {
              // Serial de Excel real (ej. 45474 = Julio)
              date = new Date((num - 25569) * 86400 * 1000);
          } else if (num >= 1 && num <= 31) {
              // DÃ­a del mes (1..31) -> Usa el mes seleccionado en la cabecera
              const monthMap = {
                  'enero':0, 'febrero':1, 'marzo':2, 'abril':3, 'mayo':4, 'junio':5,
                  'julio':6, 'agosto':7, 'septiembre':8, 'octubre':9, 'noviembre':10, 'diciembre':11
              };
              const mIdx = monthMap[monthSelected?.toLowerCase()] ?? new Date().getMonth();
              date = new Date(currentYear, mIdx, num);
          }
      } 
      // 2. Si viene como formato fecha DD/MM/YYYY o DD/MM
      else if (dateStr.includes('/')) {
          const parts = dateStr.split('/');
          if (parts.length >= 2) {
              const d = parseInt(parts[0], 10);
              const m = parseInt(parts[1], 10) - 1;
              let y = parts.length === 3 ? parseInt(parts[2], 10) : currentYear;
              if (y < 100) y += 2000;
              date = new Date(y, m, d);
          }
      }
      // 3. Formato ISO YYYY-MM-DD
      else if (dateStr.includes('-')) {
          const parts = dateStr.split('-');
          if (parts.length === 3) {
              date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          }
      }
  
      if (!date || isNaN(date.getTime())) return null;
  
      // Algoritmo oficial ISO-8601 para nÃºmero de semana
      const target = new Date(date.valueOf());
      const dayNr = (date.getDay() + 6) % 7;
      target.setDate(target.getDate() - dayNr + 3);
      const firstThursday = target.valueOf();
      target.setMonth(0, 1);
      if (target.getDay() !== 4) target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
      return 1 + Math.ceil((firstThursday - target) / 604800000);
  }
  
  /* â”€â”€ RENDERIZADO DE GRÃFICA SEMANAL CE â”€â”€ */
  function renderWeeklyCEChart() {
      if (!workbookData) return;
      
      const selSemana = document.getElementById('filter-semana');
      let fSemana = selSemana.value;
      const fVariedad = document.getElementById('filter-variedad').value;
      const fBloque = document.getElementById('filter-bloque').value;
      const fValvula = document.getElementById('filter-valvula').value;
      const activeMonth = document.getElementById('month-select')?.value || '';
      
      const emptyState = document.getElementById('ce-weekly-empty');
      const contentState = document.getElementById('ce-weekly-content');
      
      // Recopilar todas las semanas de las pestaÃ±as
      if (selSemana.dataset.loaded !== workbookData.SheetNames[0]) {
          const globalWeeks = new Set();
          workbookData.SheetNames.forEach(sheetName => {
              const ws = workbookData.Sheets[sheetName];
              const json = getSheetJson(ws);
              if (json.length === 0) return;
              const headerRow = json[0];
              for (let i = 1; i < headerRow.length; i++) {
                  const dateStr = String(headerRow[i]).trim();
                  const week = getWeekFromDateStr(dateStr, activeMonth);
                  if (week) globalWeeks.add(week);
              }
          });
          
          selSemana.innerHTML = '<option value="">Todas</option>';
          const sortedWeeks = Array.from(globalWeeks).sort((a,b)=>a-b);
          sortedWeeks.forEach(w => {
              const op = document.createElement('option');
              op.value = w; op.textContent = `Semana ${w}`;
              selSemana.appendChild(op);
          });
          
          if (sortedWeeks.length > 0) {
              selSemana.value = sortedWeeks[0]; 
              fSemana = String(sortedWeeks[0]);
          }
          selSemana.dataset.loaded = workbookData.SheetNames[0];
      }
  
      // Filtrar vÃ¡lvulas correspondientes
      const matchingSheets = workbookData.SheetNames.filter(name => {
          const m = name.match(/\((\d+)\)/);
          if (!m) return false;
          const v = parseInt(m[1]);
          if (fValvula && v !== parseInt(fValvula)) return false;
          if (fBloque && !getBlockForValve(v).includes(fBloque)) return false;
          if (fVariedad && getVarietyForValve(v) !== fVariedad) return false;
          return true;
      });
      
      if (matchingSheets.length === 0) {
          emptyState.style.display = 'flex';
          emptyState.innerHTML = '<p>No se encontraron vÃ¡lvulas para los filtros seleccionados.</p>';
          contentState.style.display = 'none';
          return;
      }
      
      const valveStats = {};
      
      matchingSheets.forEach(sheetName => {
          const m = sheetName.match(/\((\d+)\)/);
          const vNum = parseInt(m[1]);
          const ws = workbookData.Sheets[sheetName];
          const json = getSheetJson(ws);
          if (json.length === 0) return;
          
          const indices = findRowIndices(json);
          if (indices.ceDrenajeReal === -1 || indices.ceEmisorReal === -1) return;
          
          const headerRow = json[0];
          const rowDre = json[indices.ceDrenajeReal];
          const rowEmi = json[indices.ceEmisorReal];
          
          if (!valveStats[vNum]) {
              valveStats[vNum] = { dreSum: 0, dreCount: 0, emiSum: 0, emiCount: 0 };
          }
          
          const maxCols = Math.max(headerRow.length, rowDre.length, rowEmi.length);
          for (let i = 1; i < maxCols; i++) {
              const dateStr = headerRow[i] ? String(headerRow[i]).trim() : '';
              const week = getWeekFromDateStr(dateStr, activeMonth);
              
              if (fSemana && week !== parseInt(fSemana)) continue;
              
              const dreVal = parseNumericValue(rowDre[i]);
              const emiVal = parseNumericValue(rowEmi[i]);
              
              if (!isNaN(dreVal) && dreVal > 0) { valveStats[vNum].dreSum += dreVal; valveStats[vNum].dreCount++; }
              if (!isNaN(emiVal) && emiVal > 0) { valveStats[vNum].emiSum += emiVal; valveStats[vNum].emiCount++; }
          }
      });
      
      const valves = Object.keys(valveStats).map(Number).sort((a, b) => a - b);
      
      if (valves.length === 0) {
          emptyState.style.display = 'flex';
          emptyState.innerHTML = '<p>No hay datos de Conductividad para la semana o filtros seleccionados.</p>';
          contentState.style.display = 'none';
          return;
      }
      
      emptyState.style.display = 'none';
      contentState.style.display = 'block';
      
      const labels = valves.map(v => {
          const variety = getVarietyForValve(v);
          const blocks = getBlockForValve(v);
          const blockNum = blocks.length > 0 ? blocks[0].replace('Bloque ', '') : '';
          return [variety, `V${v}`, `B${blockNum}`];
      });
  
      const dataDre = valves.map(v => valveStats[v].dreCount > 0 ? (valveStats[v].dreSum / 
valveStats[v].dreCount).toFixed(2) : null);
      const dataEmi = valves.map(v => valveStats[v].emiCount > 0 ? (valveStats[v].emiSum / 
valveStats[v].emiCount).toFixed(2) : null);
      
      const ctx = document.getElementById('ce-weekly-canvas').getContext('2d');
      if (weeklyCEChart) weeklyCEChart.destroy();
      
      const limits = getLimitsForValve(valves[0]);
      let yMinGreen = limits?.ce ? limits.ce[0] : 2.4; 
      let yMaxGreen = limits?.ce ? limits.ce[1] : 3.0;
  
      const chartTitle = fVariedad ? `C.E Drenaje y Emisor - ${fVariedad}` : `Resumen C.E Drenaje y Emisor por 
VÃ¡lvula`;
  
      weeklyCEChart = new Chart(ctx, {
          type: 'bar',
          data: {
              labels: labels,
              datasets: [
                  { 
                      label: 'C.E Drenaje', 
                      data: dataDre, 
                      backgroundColor: '#b45f06', 
                      borderColor: '#8e4a04', 
                      borderWidth: 1, 
                      grouped: false,
                      barPercentage: 0.7,
                      categoryPercentage: 0.8
                  },
                  { 
                      label: 'C.E Emisor', 
                      data: dataEmi, 
                      backgroundColor: '#ff0000', 
                      borderColor: '#cc0000', 
                      borderWidth: 1, 
                      grouped: false,
                      barPercentage: 0.35,
                      categoryPercentage: 0.8
                  }
              ]
          },
          options: {
              responsive: true,
              maintainAspectRatio: false,
              scales: { 
                  y: { 
                      beginAtZero: true, 
                      suggestedMax: 4.0, 
                      title: { display: true, text: 'Conductividad ElÃ©ctrica (mS/cm)', font: { size: 12, weight: 
'bold' } } 
                  },
                  x: { 
                      ticks: { maxRotation: 0, minRotation: 0, font: { size: 10 } },
                      grid: { display: false }
                  }
              },
              plugins: { 
                  legend: { display: true, position: 'bottom' }, 
                  title: { display: true, text: chartTitle, font: { size: 15, weight: 'bold' } },
                  annotation: {
                      annotations: {
                          greenBand: {
                              type: 'box',
                              yMin: yMinGreen,
                              yMax: yMaxGreen,
                              backgroundColor: 'rgba(34, 197, 94, 0.15)',
                              borderColor: 'rgba(34, 197, 94, 0.4)',
                              borderWidth: 1,
                              drawTime: 'beforeDatasetsDraw'
                          }
                      }
                  }
              }
          }
      });
  } // Fin de la funciÃ³n renderWeeklyCEChart
  
          // === AGREGAR EVENT LISTENERS PARA QUE LOS FILTROS FUNCIONEN ===
          document.addEventListener('DOMContentLoaded', () => {
              document.getElementById('filter-semana').addEventListener('change', renderWeeklyCEChart);
              document.getElementById('filter-variedad').addEventListener('change', renderWeeklyCEChart);
              document.getElementById('filter-bloque').addEventListener('change', renderWeeklyCEChart);
              document.getElementById('filter-valvula').addEventListener('change', renderWeeklyCEChart);
          });
      </script>
  </body>
  </html>


