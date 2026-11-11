// Cambia únicamente este valor para definir la contraseña de acceso.
    const APP_PASSWORD = 'CANO';
    const authenticated = sessionStorage.getItem('survey-analyzer-authenticated') === 'true';
    if (authenticated) document.body.classList.add('authenticated');
    const get = id => document.getElementById(id);
    function logError(context, error) {
      const detail = error instanceof Error ? error.stack || error.message : String(error);
      console.error(`[Analizador de encuestas] ${context}: ${detail}`, error);
    }
    window.addEventListener('error', event => logError('Error no controlado', event.error || event.message));
    window.addEventListener('unhandledrejection', event => logError('Promesa rechazada sin controlar', event.reason));
    get('loginForm').addEventListener('submit', event => {
      event.preventDefault();
      const password = get('password').value;
      if (password === APP_PASSWORD) {
        sessionStorage.setItem('survey-analyzer-authenticated', 'true');
        document.body.classList.add('authenticated');
        get('password').value = '';
        get('loginError').textContent = '';
      } else {
        get('loginError').textContent = 'Contraseña incorrecta.';
        get('password').select();
      }
    });

    const state = { rows: [], sourceRows: [], filtered: [], previousRows: [], previousSourceRows: [], charts: {}, fileName: '', previousFileName: '', sheetName: '', previousSheetName: '', validation: null, previousValidation: null };
    const $ = id => document.getElementById(id);
    const clean = value => String(value ?? '').replace(/\s+/g, ' ').trim();
    const norm = value => clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const number = value => {
      if (typeof value === 'number') return Number.isFinite(value) ? value : null;
      const text = clean(value);
      if (!text) return null;
      const normalized = text.replace(',', '.').replace(/[^\d.-]/g, '');
      if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
      const n = Number(normalized);
      return Number.isFinite(n) ? n : null;
    };
    const money = value => new Intl.NumberFormat('es-CO', { maximumFractionDigits: 2 }).format(value);
    const pct = value => `${(value * 100).toFixed(1).replace('.', ',')} %`;
    const avg = values => values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
    const esc = text => String(text ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
    const ratingFields = [
      { key: 'solution', label: 'Solución o respuesta al requerimiento' },
      { key: 'timeliness', label: 'Oportunidad del servicio' },
      { key: 'attitude', label: 'Actitud y disposición del personal' },
      { key: 'clarity', label: 'Claridad en las indicaciones' }
    ];
    const aliases = {
      service: ['Servicio evaluado', 'Servicio'],
      dependency: ['Unidad Administrativa', 'Dependencia', 'Unidad administrativa'],
      name: ['Nombre2', 'Nombre'],
      start: ['Hora de inicio', 'Fecha de inicio', 'Fecha y hora', 'Marca temporal', 'Timestamp'],
      suggestion: ['Sugerencias y/o recomendaciones'],
      solution: ['¿Se dio solución o respuesta a sus requerimientos?'],
      timeliness: ['¿Los servicios fueron prestados de manera oportuna?'],
      attitude: ['¿Cómo califica la actitud y disposición del personal que atendió sus requerimientos?'],
      clarity: ['¿Las indicaciones del personal que atendió sus requerimientos fueron claras?']
    };

    function findKey(headers, possible) {
      const wanted = possible.map(norm);
      return headers.find(h => wanted.includes(norm(h))) || headers.find(h => wanted.some(w => norm(h).includes(w)));
    }
    function headerInfo(raw) {
      let best = { index: -1, headers: [], keys: {}, score: -1 };
      const limit = Math.min(raw.length, 20);
      for (let index = 0; index < limit; index += 1) {
        const headers = (raw[index] || []).map(clean);
        const keys = {};
        for (const [semantic, options] of Object.entries(aliases)) keys[semantic] = findKey(headers, options);
        const score = Object.values(keys).filter(Boolean).length;
        if (score > best.score) best = { index, headers, keys, score };
      }
      return best;
    }
    function dateFromParts(year, month, day) {
      const date = new Date(year, month, day);
      return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day ? date : null;
    }
    function parseDate(value) {
      if (value instanceof Date && !isNaN(value)) {
        const year = value.getFullYear();
        const month = value.getMonth();
        const day = value.getDate();
        return dateFromParts(year, month, day);
      }
      if (typeof value === 'number' && window.XLSX?.SSF?.parse_date_code) {
        const parts = XLSX.SSF.parse_date_code(value);
        if (parts && parts.y && parts.m && parts.d) return dateFromParts(parts.y, parts.m - 1, parts.d);
      }
      const text = clean(value);
      const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (isoDate) return dateFromParts(Number(isoDate[1]), Number(isoDate[2]) - 1, Number(isoDate[3]));
      const localDate = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);
      if (localDate) return dateFromParts(Number(localDate[3]), Number(localDate[2]) - 1, Number(localDate[1]));
      const textDate = text.match(/^(\d{1,2})\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{4})$/i);
      if (textDate) {
        const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
        const month = months.indexOf(norm(textDate[2]));
        if (month >= 0) return dateFromParts(Number(textDate[3]), month, Number(textDate[1]));
      }
      return null;
    }
    function mapRows(raw) {
      const header = headerInfo(raw);
      const { headers, keys, index: headerIndex } = header;
      const missing = [
        !keys.service && 'Servicio evaluado',
        !keys.dependency && 'Unidad Administrativa o dependencia',
        !keys.start && 'Hora de inicio o fecha'
      ].filter(Boolean);
      if (missing.length) {
        const available = headers.filter(Boolean).slice(0, 20).join(', ') || 'ninguna';
        throw new Error(`No se encontraron las columnas mínimas: ${missing.join('; ')}. Encabezados detectados: ${available}.`);
      }
      return raw.slice(headerIndex + 1).map((cells, index) => {
        if (!cells.some(value => clean(value))) return null;
        const obj = Object.fromEntries(headers.map((h, i) => [h, cells[i]]));
        const date = parseDate(obj[keys.start]);
        const ratings = {};
        for (const item of ratingFields) ratings[item.key] = number(obj[keys[item.key]]);
        const suggestion = clean(obj[keys.suggestion]);
        const finalAverage = avg(Object.values(ratings).filter(v => v !== null));
        const sourceRow = headerIndex + index + 2;
        const sourceId = obj.ID ?? obj.Id;
        const id = clean(sourceId) || `fila-${sourceRow}`;
        const errors = [];
        if (!date) errors.push('Fecha inválida o ausente');
        if (!clean(obj[keys.service])) errors.push('Servicio ausente');
        if (!clean(obj[keys.dependency])) errors.push('Dependencia ausente');
        return {
          row: sourceRow,
          sourceRow,
          id,
          explicitId: Boolean(clean(sourceId)),
          service: clean(obj[keys.service]),
          dependency: clean(obj[keys.dependency]),
          respondent: clean(obj[keys.name]),
          suggestion,
          date,
          month: date ? date.getMonth() + 1 : null,
          year: date ? date.getFullYear() : null,
          ratings,
          average: finalAverage,
          raw: cells.slice(),
          miguel: /miguel\s*[aá]ngel\s*correa|acompa[nñ]ada por miguel/i.test(suggestion),
          validationErrors: errors,
          valid: errors.length === 0,
          recordKey: `${clean(sourceId) || 'sin-id'}|${date ? date.toISOString().slice(0, 10) : 'sin-fecha'}|${norm(obj[keys.service])}|${norm(obj[keys.dependency])}|${Object.values(ratings).join('|')}|${norm(suggestion)}`
        };
      }).filter(Boolean);
    }
    function parseWorkbook(buffer) {
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const errors = [];
      for (const sheetName of workbook.SheetNames) {
        try {
          const sheet = workbook.Sheets[sheetName];
          const sourceRows = mapRows(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true }));
          return { rows: sourceRows.filter(row => row.valid), sourceRows, sheetName };
        } catch (error) {
          errors.push(`${sheetName}: ${error.message}`);
        }
      }
      throw new Error(`No fue posible identificar una hoja válida. ${errors.join(' | ')}`);
    }
    function periodKey(row) {
      return row.date ? `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, '0')}` : '';
    }
    function splitPeriods(rows) {
      const periods = [...new Set(rows.map(periodKey).filter(Boolean))].sort();
      if (periods.length < 2) return { current: rows, previous: [] };
      const currentPeriod = periods[periods.length - 1];
      const previousPeriod = periods[periods.length - 2];
      return {
        current: rows.filter(row => periodKey(row) === currentPeriod),
        previous: rows.filter(row => periodKey(row) === previousPeriod)
      };
    }
    function options() {
      const years = [...new Set(state.rows.map(r => r.year).filter(Boolean))].sort((a,b) => a-b);
      const months = [...new Set(state.rows.map(r => r.month).filter(Boolean))].sort((a,b) => a-b);
      $('year').innerHTML = '<option value="all">Todos</option>' + years.map(year => `<option value="${year}">${year}</option>`).join('');
      $('month').innerHTML = '<option value="all">Todos</option>' + months.map(m => `<option value="${m}">${new Date(2020, m-1, 1).toLocaleString('es-CO', {month:'long'})}</option>`).join('');
      const deps = [...new Set(state.rows.map(r => r.dependency).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'es'));
      $('dependency').innerHTML = '<option value="all">Todas</option>' + deps.map(d => `<option value="${esc(d)}">${esc(d)}</option>`).join('');
    }
    function duplicateGroups(rows, keySelector) {
      const groups = new Map();
      rows.forEach(row => {
        const key = keySelector(row);
        if (!key) return;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(row);
      });
      return [...groups.entries()]
        .filter(([, values]) => values.length > 1)
        .map(([key, values]) => ({ key, rows: values.map(row => ({ sourceRow: row.sourceRow, id: row.id, period: periodKey(row) })) }));
    }
    function validateDataset(rows, sourceRows, filterError = '', processedRows = rows) {
      const invalidSource = sourceRows.filter(row => !row.valid);
      const sourceRowsByNumber = new Set(sourceRows.map(row => row.sourceRow));
      const processedRowsByNumber = new Set(processedRows.map(row => row.sourceRow));
      const omittedRows = processedRows.filter(row => !sourceRowsByNumber.has(row.sourceRow));
      const duplicateGroupsFound = duplicateGroups(processedRows, row => row.recordKey);
      const duplicateIdGroups = duplicateGroups(processedRows.filter(row => row.explicitId), row => `${row.id}|${periodKey(row)}`);
      const duplicates = duplicateGroupsFound.map(group => group.key);
      const duplicateIdsByPeriod = duplicateIdGroups.map(group => group.key);
      const periodMismatches = processedRows.filter(row => !row.date || row.year !== row.date.getFullYear() || row.month !== row.date.getMonth() + 1);
      const invalidDetails = invalidSource.map(row => `${row.sourceRow} (ID ${row.id}): ${row.validationErrors.join(', ')}`);
      const duplicateDetails = [...duplicateGroupsFound, ...duplicateIdGroups]
        .map(group => `filas ${group.rows.map(row => row.sourceRow).join(', ')}`)
        .filter((detail, index, list) => list.indexOf(detail) === index);
      const errors = [];
      if (invalidSource.length) errors.push(`${invalidSource.length} registro(s) de origen no son válidos. Filas: ${invalidDetails.slice(0, 8).join(' · ')}`);
      if (omittedRows.length || processedRowsByNumber.size !== processedRows.length) errors.push('Se detectaron registros omitidos o repetidos durante el procesamiento.');
      if (duplicates.length || duplicateIdsByPeriod.length) errors.push(`Se detectaron registros duplicados dentro de un mismo periodo. ${duplicateDetails.slice(0, 8).join(' · ')}`);
      if (periodMismatches.length) errors.push('Se detectaron registros asignados a un año o mes diferente de su fecha.');
      if (filterError) errors.push(filterError);
      return {
        ok: errors.length === 0,
        errors,
        sourceTotal: sourceRows.length,
        processedTotal: processedRows.length,
        filteredTotal: rows.length,
        invalidTotal: invalidSource.length,
        invalidDetails,
        duplicateKeys: duplicates,
        duplicateIdsByPeriod,
        duplicateDetails,
        periodMismatches: periodMismatches.length,
        omittedTotal: omittedRows.length
      };
    }
    function applyFilter() {
      const mode = $('period').value;
      const year = $('year').value;
      const month = $('month').value;
      const quarter = $('quarter').value;
      const dependency = $('dependency').value;
      const years = [...new Set(state.rows.map(row => row.year).filter(Boolean))];
      const needsYear = year === 'all' && years.length > 1 && (mode === 'quarter' || (mode === 'month' && month !== 'all'));
      const filterErrors = [];
      if (needsYear) filterErrors.push('Seleccione un año para analizar el periodo sin mezclar meses de años diferentes.');
      if (state.previousValidation && !state.previousValidation.ok) filterErrors.push(`El periodo anterior tiene inconsistencias: ${state.previousValidation.errors.join(' ')}`);
      const filterError = filterErrors.join(' ');
      state.filtered = state.rows.filter(r => {
        const yearOk = year === 'all' || String(r.year) === year;
        const periodOk = mode === 'all'
          || (mode === 'month' && (month === 'all' || String(r.month) === month))
          || (mode === 'quarter' && (quarter === 'all' || String(Math.ceil(r.month / 3)) === quarter));
        return yearOk && periodOk && (dependency === 'all' || r.dependency === dependency);
      });
      state.validation = validateDataset(state.filtered, state.sourceRows, filterError, state.rows);
      updatePeriodLabel();
      render();
    }
    function currentReportRows() {
      const dependency = $('dependency').value;
      const periods = [...new Set(state.rows.map(periodKey).filter(Boolean))].sort();
      if ($('period').value === 'all' && periods.length > 1) {
        const latestPeriod = periods[periods.length - 1];
        return state.rows.filter(row => periodKey(row) === latestPeriod && (dependency === 'all' || row.dependency === dependency));
      }
      return state.filtered;
    }
    function previousReportRows(current) {
      const dependency = $('dependency').value;
      const currentPeriod = [...new Set(current.map(periodKey).filter(Boolean))].sort().pop();
      const candidates = [...state.previousRows, ...state.rows];
      const unique = new Map();
      candidates.forEach(row => {
        const key = `${row.id}|${periodKey(row)}`;
        if (!unique.has(key)) unique.set(key, row);
      });
      const periods = [...new Set([...unique.values()].map(periodKey).filter(period => period && (!currentPeriod || period < currentPeriod)))].sort();
      const previousPeriod = periods[periods.length - 1];
      return [...unique.values()].filter(row => periodKey(row) === previousPeriod && (dependency === 'all' || row.dependency === dependency));
    }
    function updatePeriodLabel() {
      const mode = $('period').value;
      const year = $('year').value;
      const month = $('month').value;
      const quarter = $('quarter').value;
      const dependency = $('dependency').value;
      let label = 'Todos los registros';
      if (mode === 'month' && month !== 'all') {
        const monthName = new Date(2020, Number(month) - 1, 1).toLocaleString('es-CO', { month: 'long' });
        const years = [...new Set(state.filtered.map(row => row.year).filter(Boolean))].sort();
        label = `${monthName}${years.length === 1 ? ` de ${years[0]}` : ''}`;
      } else if (mode === 'quarter' && quarter !== 'all') {
        const quarterNames = {
          1: 'Primer trimestre',
          2: 'Segundo trimestre',
          3: 'Tercer trimestre',
          4: 'Cuarto trimestre'
        };
        label = `${quarterNames[quarter]}${year !== 'all' ? ` de ${year}` : ''}`;
      } else if (mode === 'month') {
        label = 'Todos los meses';
      } else if (mode === 'quarter') {
        label = 'Todos los trimestres';
      }
      if (dependency !== 'all') label += ` · Dependencia: ${dependency}`;
      $('periodLabel').textContent = `Periodo analizado: ${label}`;
    }
    function group(rows, field) {
      const map = new Map();
      rows.forEach(r => {
        const key = r[field] || 'Sin dato';
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(r);
      });
      return [...map.entries()].map(([key, values]) => ({ key, values, count: values.length, average: avg(values.map(v => v.average)), miguel: values.filter(v => v.miguel).length })).sort((a,b) => b.count - a.count);
    }
    function destroyChart(key) { if (state.charts[key]) state.charts[key].destroy(); }
    function chart(key, canvas, config) { destroyChart(key); state.charts[key] = new Chart($(canvas), config); }
    function baseScales() { return { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: '#e2e8f0' } } }; }
    function render() {
      const rows = state.filtered;
      const total = rows.length;
      const overall = avg(rows.map(r => r.average));
      const below = rows.filter(r => r.average < 5).sort((a,b) => a.average - b.average);
      const miguel = rows.filter(r => r.miguel).length;
      const services = group(rows, 'service');
      const deps = group(rows, 'dependency');
      $('mTotal').textContent = total.toLocaleString('es-CO');
      $('mAverage').textContent = total ? overall.toFixed(2).replace('.', ',') : '—';
      $('mBelow').textContent = total ? `${below.length} (${pct(below.length / total)})` : '—';
      $('mMiguel').textContent = total ? `${miguel} (${pct(miguel / total)})` : '—';
      $('mService').textContent = services[0]?.key || '—';
      $('mServiceHint').textContent = services[0] ? `${services[0].count} encuestas · ${pct(services[0].count / total)}` : '';
      $('mDependency').textContent = deps[0]?.key || '—';
      $('mDependencyHint').textContent = deps[0] ? `${deps[0].count} encuestas · ${pct(deps[0].count / total)}` : '';
      renderCharts(rows, services, deps);
      renderTables(rows, services, deps, below);
      renderQuality(rows);
      const reportRows = currentReportRows();
      $('report').textContent = buildReport(reportRows, group(reportRows, 'service'), group(reportRows, 'dependency'), reportRows.filter(row => row.average < 5).sort((a, b) => a.average - b.average), reportRows.filter(row => row.miguel).length);
      if (state.validation && !state.validation.ok) {
        $('status').className = 'status error';
        $('status').textContent = `Validación detenida: ${state.validation.errors.join(' ')}`;
      }
      ['exportWord', 'exportData', 'exportCsv'].forEach(id => { $(id).disabled = !total || Boolean(state.validation && !state.validation.ok); });
    }
    function renderQuality(rows) {
      const duplicateIds = state.validation?.duplicateIdsByPeriod || [];
      const missing = {
        'Fecha de inicio': rows.filter(r => !r.date).length,
        'Servicio evaluado': rows.filter(r => !r.service).length,
        'Dependencia': rows.filter(r => !r.dependency).length,
        'Criterios incompletos': rows.filter(r => Object.values(r.ratings).some(v => v === null)).length
      };
      const rowsHtml = [
        ['Registros analizados', rows.length, 'OK'],
        ['Registros originales', state.validation?.sourceTotal ?? rows.length, state.validation?.sourceTotal === state.validation?.processedTotal ? 'OK' : 'Revisar'],
        ['Registros procesados', state.validation?.processedTotal ?? rows.length, state.validation?.invalidTotal ? `Revisar (${state.validation.invalidTotal} inválidos)` : 'OK'],
        ['Registros omitidos', state.validation?.omittedTotal ?? 0, state.validation?.omittedTotal ? 'Revisar' : 'OK'],
        ['Participación por servicio', `${group(rows, 'service').reduce((a, b) => a + b.count, 0)}/${rows.length}`, '100 %'],
        ['Participación por dependencia', `${group(rows, 'dependency').reduce((a, b) => a + b.count, 0)}/${rows.length}`, '100 %'],
        ['Registros duplicados', state.validation?.duplicateKeys?.length || 0, state.validation?.duplicateKeys?.length ? 'Revisar' : 'No detectados'],
        ['IDs duplicados por periodo', duplicateIds.length, duplicateIds.length ? `Revisar: ${duplicateIds.slice(0, 8).join(', ')}` : 'No detectados'],
        ['Asignación temporal', state.validation?.periodMismatches || 0, state.validation?.periodMismatches ? 'Revisar' : 'OK'],
        ['Detalle de inconsistencias', state.validation?.errors?.join(' ') || 'Sin inconsistencias', state.validation?.ok ? 'OK' : 'Revisar'],
        ...Object.entries(missing).map(([label, count]) => [label, count, count ? 'Revisar' : 'OK'])
      ];
      $('qualityTable').innerHTML = `<table><thead><tr><th>Validación</th><th class="num">Resultado</th><th>Estado</th></tr></thead><tbody>${rowsHtml.map(row => `<tr><td>${esc(row[0])}</td><td class="num">${esc(row[1])}</td><td class="${String(row[2]).startsWith('OK') || row[2] === '100 %' || row[2] === 'No detectados' ? 'good' : 'risk'}">${esc(row[2])}</td></tr>`).join('')}</tbody></table>`;
    }
    function renderCharts(rows, services, deps) {
      chart('service', 'serviceChart', { type: 'doughnut', data: { labels: services.map(x => x.key), datasets: [{ data: services.map(x => x.count), backgroundColor: ['#2563eb','#0f766e','#f59e0b','#9333ea','#ef4444','#64748b'] }] }, options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } } });
      const topDeps = deps.slice(0, 10).reverse();
      chart('dependency', 'dependencyChart', { type: 'bar', data: { labels: topDeps.map(x => x.key), datasets: [{ label: 'Encuestas', data: topDeps.map(x => x.count), backgroundColor: '#0f766e', borderRadius: 5 }] }, options: { maintainAspectRatio: false, indexAxis: 'y', scales: baseScales(), plugins: { legend: { display: false } } } });
      const periods = [...new Set(rows.map(periodKey).filter(Boolean))].sort();
      const monthly = periods.map(period => rows.filter(r => periodKey(r) === period).length);
      chart('monthly', 'monthlyChart', { type: 'line', data: { labels: periods.map(period => {
        const [year, month] = period.split('-').map(Number);
        return `${new Date(year, month - 1, 1).toLocaleString('es-CO', { month: 'short' })} ${year}`;
      }), datasets: [{ label: 'Encuestas', data: monthly, borderColor: '#2563eb', backgroundColor: '#dbeafe', fill: true, tension: .3 }] }, options: { maintainAspectRatio: false, scales: baseScales() } });
      const criteria = ratingFields.map(item => avg(rows.map(r => r.ratings[item.key]).filter(v => v !== null)));
      chart('criteria', 'criteriaChart', { type: 'bar', data: { labels: ['Solución','Oportunidad','Actitud','Claridad'], datasets: [{ label: 'Promedio', data: criteria, backgroundColor: ['#2563eb','#0f766e','#f59e0b','#9333ea'], borderRadius: 5 }] }, options: { maintainAspectRatio: false, scales: { x: { grid: { display: false } }, y: { min: 0, max: 5, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } } });
    }
    function renderTables(rows, services, deps, below) {
      $('serviceTable').innerHTML = services.length ? `<table><thead><tr><th>Servicio</th><th class="num">Encuestas</th><th class="num">Participación</th><th class="num">Promedio</th><th class="num">Miguel Ángel</th></tr></thead><tbody>${services.map(x => `<tr><td>${esc(x.key)}</td><td class="num">${x.count}</td><td class="num">${pct(x.count / rows.length)}</td><td class="num">${x.average.toFixed(2)}</td><td class="num">${x.miguel} (${pct(x.miguel / x.count)})</td></tr>`).join('')}</tbody></table>` : '<div class="empty">No hay datos para el filtro seleccionado.</div>';
      $('dependencyTable').innerHTML = deps.length ? `<table><thead><tr><th>Dependencia</th><th class="num">Encuestas</th><th class="num">Participación</th><th class="num">Promedio</th></tr></thead><tbody>${deps.map(x => `<tr><td>${esc(x.key)}</td><td class="num">${x.count}</td><td class="num">${pct(x.count / rows.length)}</td><td class="num">${x.average.toFixed(2)}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">No hay datos para el filtro seleccionado.</div>';
      $('casesTable').innerHTML = below.length ? `<table><thead><tr><th>ID / fila</th><th>Fecha</th><th>Dependencia</th><th>Servicio</th><th class="num">Promedio</th><th>Observación</th></tr></thead><tbody>${below.map(r => `<tr><td>${esc(r.id)} / ${r.row}</td><td>${r.date ? r.date.toLocaleDateString('es-CO') : 'Sin fecha'}</td><td>${esc(r.dependency)}</td><td>${esc(r.service)}</td><td class="num ${r.average <= 3 ? 'risk' : ''}">${r.average.toFixed(2)}</td><td>${esc(r.suggestion || 'Sin observación')}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">No se identifican calificaciones inferiores a 5,00.</div>';
    }
    function buildReport(rows, services, deps, below, miguel) {
      if (!rows.length) return 'No hay registros para el filtro seleccionado.';
      const total = rows.length;
      const overall = avg(rows.map(r => r.average));
      const criteria = ratingFields.map(item => ({ label: item.label, value: avg(rows.map(r => r.ratings[item.key]).filter(v => v !== null)) }));
      const min = [...criteria].sort((a,b) => a.value - b.value)[0];
      const max = [...criteria].sort((a,b) => b.value - a.value)[0];
      const lowNames = below.slice(0, 3).map(r => `${r.id} (${r.average.toFixed(2)})`).join(', ') || 'ninguno';
      const periodLabel = $('periodLabel').textContent.replace('Periodo analizado: ', '');
      return `ANÁLISIS EJECUTIVO\n\nPeriodo analizado: ${periodLabel}.\n\nSe analizaron ${total.toLocaleString('es-CO')} encuestas de satisfacción. La calificación promedio general fue de ${overall.toFixed(2).replace('.', ',')} sobre 5,00. ${below.length ? `Se identificaron ${below.length} encuestas con promedio inferior a 5,00 (${pct(below.length / total)}), que requieren revisión puntual.` : 'No se identificaron encuestas con promedio inferior a 5,00.'}\n\nLa mayor participación correspondió al servicio ${services[0].key}, con ${services[0].count} encuestas (${pct(services[0].count / total)}). La dependencia con mayor participación fue ${deps[0].key}, con ${deps[0].count} encuestas (${pct(deps[0].count / total)}).\n\nPor criterio, el promedio más bajo fue ${min.label} (${min.value.toFixed(2)}) y el más alto fue ${max.label} (${max.value.toFixed(2)}). Se identificaron ${miguel} encuestas asociadas al acompañamiento de Miguel Ángel Correa (${pct(miguel / total)}).\n\nLos casos con menor valoración fueron: ${lowNames}. Se recomienda verificar la solución efectiva antes del cierre, fortalecer la comunicación con el usuario y documentar las causas de los casos inferiores para determinar si corresponden a situaciones aisladas o repetitivas.\n\nVALIDACIONES AUTOMÁTICAS\n• Participación por servicio: ${services.reduce((a,b) => a + b.count, 0) === total ? '100 %' : 'Revisar'}.\n• Participación por dependencia: ${deps.reduce((a,b) => a + b.count, 0) === total ? '100 %' : 'Revisar'}.\n• Promedio general: calculado como promedio de los promedios individuales disponibles.\n• La identificación de Miguel Ángel se basa en el texto de observaciones; debe confirmarse contra la regla institucional de marcación.`;
    }
    function toggleControls() {
      const mode = $('period').value;
      $('month').disabled = mode !== 'month';
      $('quarter').disabled = mode !== 'quarter';
    }
    $('file').addEventListener('change', event => {
      const file = event.target.files[0];
      if (!file) return;
      state.fileName = file.name;
      state.previousRows = [];
      state.previousSourceRows = [];
      state.previousFileName = '';
      state.previousSheetName = '';
      state.previousValidation = null;
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const parsed = parseWorkbook(e.target.result);
          const periods = splitPeriods(parsed.rows);
          state.rows = parsed.rows;
          state.sourceRows = parsed.sourceRows;
          if (periods.previous.length) {
            state.previousRows = periods.previous;
            state.previousSourceRows = parsed.sourceRows.filter(row => periodKey(row) === periodKey(periods.previous[0]));
            state.previousFileName = `${file.name} · periodo anterior`;
          }
          state.sheetName = parsed.sheetName;
          options();
          $('fileName').textContent = file.name;
          $('fileMeta').textContent = `${state.rows.length.toLocaleString('es-CO')} registros válidos de ${state.sourceRows.length.toLocaleString('es-CO')} filas con información · hoja: ${state.sheetName}`;
          $('status').className = 'status';
          $('status').textContent = periods.previous.length
            ? `Archivo cargado: ${state.rows.length.toLocaleString('es-CO')} registros válidos. Se detectaron automáticamente el periodo actual y el anterior para el comparativo.`
            : 'Archivo cargado y validado localmente. Cargue el periodo anterior para completar el comparativo.';
          applyFilter();
        } catch (error) {
          logError(`Error al cargar el consolidado "${file.name}"`, error);
          $('status').className = 'status error';
          $('status').textContent = error.message || 'No fue posible leer el archivo.';
        }
      };
      reader.readAsArrayBuffer(file);
    });
    $('previousFile').addEventListener('change', event => {
      const file = event.target.files[0];
      if (!file) return;
      state.previousFileName = file.name;
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const parsed = parseWorkbook(e.target.result);
          state.previousRows = parsed.rows;
          state.previousSourceRows = parsed.sourceRows;
          state.previousSheetName = parsed.sheetName;
          state.previousValidation = validateDataset(parsed.rows, parsed.sourceRows, '', parsed.rows);
          $('status').className = 'status';
          $('status').textContent = state.previousValidation.ok
            ? `Periodo anterior cargado: ${state.previousRows.length.toLocaleString('es-CO')} registros.`
            : `Validación detenida en el periodo anterior: ${state.previousValidation.errors.join(' ')}`;
          applyFilter();
        } catch (error) {
          logError(`Error al cargar el periodo anterior "${file.name}"`, error);
          state.previousRows = [];
          state.previousSourceRows = [];
          state.previousValidation = null;
          $('status').className = 'status error';
          $('status').textContent = error.message || 'No fue posible leer el periodo anterior.';
        }
      };
      reader.readAsArrayBuffer(file);
    });
    $('period').addEventListener('change', () => { toggleControls(); applyFilter(); });
    $('year').addEventListener('change', applyFilter);
    $('month').addEventListener('change', applyFilter);
    $('quarter').addEventListener('change', applyFilter);
    $('dependency').addEventListener('change', applyFilter);
    function downloadBlob(blob, fileName) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function monthYearLabel(rows) {
      const dates = rows.map(row => row.date).filter(Boolean);
      if (!dates.length) return $('periodLabel').textContent.replace('Periodo analizado: ', '').toUpperCase();
      const first = dates.slice().sort((a, b) => a - b)[0];
      const last = dates.slice().sort((a, b) => b - a)[0];
      const firstLabel = first.toLocaleString('es-CO', { month: 'long', year: 'numeric' });
      const lastLabel = last.toLocaleString('es-CO', { month: 'long', year: 'numeric' });
      return firstLabel === lastLabel ? firstLabel.toUpperCase() : `${firstLabel} - ${lastLabel}`.toUpperCase();
    }

    function dateRangeLabel(rows) {
      const dates = rows.map(row => row.date).filter(Boolean).sort((a, b) => a - b);
      if (!dates.length) return monthYearLabel(rows);
      const first = dates[0];
      const last = dates[dates.length - 1];
      const month = first.toLocaleString('es-CO', { month: 'long' });
      const year = first.getFullYear();
      return first.getMonth() === last.getMonth() && first.getFullYear() === last.getFullYear()
        ? `${String(first.getDate()).padStart(2, '0')} al ${String(last.getDate()).padStart(2, '0')} de ${month} de ${year}`
        : `${first.toLocaleDateString('es-CO')} al ${last.toLocaleDateString('es-CO')}`;
    }

    function metrics(rows) {
      const services = group(rows, 'service');
      const deps = group(rows, 'dependency');
      const below = rows.filter(row => row.average < 5).sort((a, b) => a.average - b.average);
      const criteria = ratingFields.map(item => ({
        key: item.key,
        label: item.label,
        shortLabel: { solution: 'Solución', timeliness: 'Oportunidad', attitude: 'Actitud', clarity: 'Claridad' }[item.key],
        average: avg(rows.map(row => row.ratings[item.key]).filter(value => value !== null))
      }));
      const serviceNames = [...new Set(rows.map(row => row.service).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
      const matrix = [...new Set(rows.map(row => row.dependency).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es')).map(dependency => ({
        dependency,
        values: serviceNames.map(service => rows.filter(row => row.dependency === dependency && row.service === service).length),
        total: rows.filter(row => row.dependency === dependency).length
      }));
      return {
        rows,
        total: rows.length,
        average: avg(rows.map(row => row.average)),
        below,
        perfect: rows.filter(row => row.average === 5).length,
        miguel: rows.filter(row => row.miguel).length,
        services,
        deps,
        criteria,
        serviceNames,
        matrix,
        quality: {
          analyzed: rows.length,
          missingDate: rows.filter(row => !row.date).length,
          missingService: rows.filter(row => !row.service).length,
          missingDependency: rows.filter(row => !row.dependency).length,
          incompleteCriteria: rows.filter(row => Object.values(row.ratings).some(value => value === null)).length,
          duplicateIds: [...new Set(rows.map(row => String(row.id)).filter((id, index, list) => list.indexOf(id) !== index))]
        }
      };
    }

    function comparison(current, previous) {
      if (!previous.length) return null;
      const currentMetrics = metrics(current);
      const previousMetrics = metrics(previous);
      const currentService = currentMetrics.services[0];
      const previousService = previousMetrics.services[0];
      const currentDependency = currentMetrics.deps[0];
      const previousDependency = previousMetrics.deps[0];
      return {
        previousPeriod: monthYearLabel(previous),
        total: { current: currentMetrics.total, previous: previousMetrics.total, difference: currentMetrics.total - previousMetrics.total },
        average: { current: currentMetrics.average, previous: previousMetrics.average, difference: currentMetrics.average - previousMetrics.average },
        belowFive: { current: currentMetrics.below.length, previous: previousMetrics.below.length, difference: currentMetrics.below.length - previousMetrics.below.length },
        dependency: { current: currentDependency?.key || 'Sin dato', currentCount: currentDependency?.count || 0, previous: previousDependency?.key || 'Sin dato', previousCount: previousDependency?.count || 0 },
        service: { current: currentService?.key || 'Sin dato', currentCount: currentService?.count || 0, previous: previousService?.key || 'Sin dato', previousCount: previousService?.count || 0 }
      };
    }

    function historicalComparison(rows) {
      const buckets = new Map();
      rows.forEach(row => {
        const key = periodKey(row);
        if (!key) return;
        if (!buckets.has(key)) buckets.set(key, []);
        buckets.get(key).push(row);
      });
      return [...buckets.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([key, values], index, all) => {
        const current = metrics(values);
        const priorValues = index ? all[index - 1][1] : [];
        const prior = priorValues.length ? metrics(priorValues) : null;
        return {
          period: monthYearLabel(values),
          total: current.total,
          average: current.average,
          belowFive: current.below.length,
          topDependency: current.deps[0]?.key || 'Sin dato',
          topDependencyCount: current.deps[0]?.count || 0,
          topService: current.services[0]?.key || 'Sin dato',
          topServiceCount: current.services[0]?.count || 0,
          variation: prior ? {
            total: current.total - prior.total,
            average: current.average - prior.average,
            belowFive: current.below.length - prior.below.length
          } : null
        };
      });
    }

    function caseAnalysis(row) {
      const lowest = ratingFields
        .filter(item => row.ratings[item.key] !== null)
        .sort((a, b) => row.ratings[a.key] - row.ratings[b.key]);
      const lowestLabels = lowest.filter(item => row.ratings[item.key] === row.ratings[lowest[0]?.key]).map(item => item.label.toLowerCase());
      const observation = row.suggestion ? ` La observación registrada indica: “${row.suggestion}”.` : '';
      return `El registro ${row.id}, correspondiente a ${row.dependency || 'una dependencia no identificada'}, obtuvo un promedio de ${row.average.toFixed(2)}. ${lowestLabels.length ? `La principal oportunidad se relaciona con ${lowestLabels.join(' y ')}.` : ''}${observation}`;
    }

    function reportNarratives(current, previous) {
      const data = metrics(current);
      const comparisonData = comparison(current, previous);
      const historySource = [...current];
      previous.forEach(row => {
        if (!historySource.some(existing => String(existing.id) === String(row.id) && periodKey(existing) === periodKey(row))) historySource.push(row);
      });
      const history = historicalComparison(historySource);
      const strongestCriteria = data.criteria.filter(item => item.average >= 4.99).map(item => item.shortLabel.toLowerCase());
      const lowestCriterion = data.criteria.slice().sort((a, b) => a.average - b.average)[0];
      const strengths = [
        `Se analizaron ${data.total.toLocaleString('es-CO')} encuestas con una calificación promedio general de ${data.average.toFixed(2).replace('.', ',')} sobre 5,00.`,
        `${data.perfect.toLocaleString('es-CO')} encuestas (${data.total ? pct(data.perfect / data.total) : '0,0 %'}) obtuvieron un promedio individual de 5,00.`,
        `El servicio con mayor participación fue ${data.services[0]?.key || 'Sin dato'}, con ${data.services[0]?.count || 0} encuestas (${data.total ? pct((data.services[0]?.count || 0) / data.total) : '0,0 %'}).`,
        strongestCriteria.length ? `Los criterios con resultados máximos o prácticamente máximos fueron: ${strongestCriteria.join(', ')}.` : `Los cuatro criterios evaluados presentaron resultados superiores a ${Math.max(0, lowestCriterion.average - 0.02).toFixed(2)}.`
      ];
      const opportunities = [
        data.below.length ? `Se requiere revisar ${data.below.length} encuesta(s) con promedio inferior a 5,00, especialmente el registro ${data.below[0].id}.` : 'No se identificaron casos con promedio inferior a 5,00.',
        `El criterio con menor promedio fue ${lowestCriterion.label.toLowerCase()} (${lowestCriterion.average.toFixed(2)}).`,
        'Mantener la verificación de la solución antes del cierre y reforzar la claridad de las indicaciones cuando el usuario requiera acompañamiento adicional.',
        data.quality.incompleteCriteria ? `Antes de aprobar el informe deben revisarse ${data.quality.incompleteCriteria} registro(s) con criterios incompletos.` : 'No se detectaron criterios incompletos en los registros analizados.'
      ];
      if (comparisonData) {
        opportunities.push(`Frente a ${comparisonData.previousPeriod.toLowerCase()}, el volumen de encuestas varió en ${comparisonData.total.difference > 0 ? 'un aumento' : 'una reducción'} de ${Math.abs(comparisonData.total.difference)} registro(s).`);
      }
      const conclusion = `Durante ${monthYearLabel(current).toLowerCase()} se analizaron ${data.total} encuestas de satisfacción, obteniendo una calificación promedio general de ${data.average.toFixed(2).replace('.', ',')} sobre 5,00. Del total, ${data.perfect} encuestas (${pct(data.perfect / data.total)}) alcanzaron un promedio individual de 5,00 y ${data.below.length} presentaron resultados inferiores a 5,00, por lo que deben ser revisadas según el criterio con menor valoración y la observación registrada. El servicio con mayor participación fue ${data.services[0]?.key || 'Sin dato'}, mientras que la dependencia con mayor número de respuestas fue ${data.deps[0]?.key || 'Sin dato'}. ${comparisonData ? `En comparación con ${comparisonData.previousPeriod.toLowerCase()}, el total de encuestas ${comparisonData.total.difference === 0 ? 'se mantuvo sin variación' : comparisonData.total.difference > 0 ? `aumentó en ${comparisonData.total.difference}` : `disminuyó en ${Math.abs(comparisonData.total.difference)}`}; el promedio general ${Math.abs(comparisonData.average.difference) < 0.005 ? 'se mantuvo estable' : comparisonData.average.difference > 0 ? `aumentó en ${comparisonData.average.difference.toFixed(2)}` : `disminuyó en ${Math.abs(comparisonData.average.difference).toFixed(2)}`}; y los casos inferiores a 5,00 ${comparisonData.belowFive.difference === 0 ? 'no presentaron variación' : comparisonData.belowFive.difference > 0 ? `aumentaron en ${comparisonData.belowFive.difference}` : `disminuyeron en ${Math.abs(comparisonData.belowFive.difference)}`}.` : 'El comparativo con el mes anterior no fue calculado porque no se proporcionó una base histórica.'} En consecuencia, se recomienda mantener las fortalezas identificadas, verificar la solución antes del cierre, mejorar la oportunidad y claridad cuando corresponda y hacer seguimiento prioritario a los casos con menor calificación.`;
      const management = [
        `Se identificaron ${data.miguel} encuestas asociadas al acompañamiento de Miguel Ángel Correa Cano (${pct(data.miguel / data.total)} del total).`,
        'Optimizar la estrategia de aplicación de encuestas, buscando mayor efectividad en la comunicación con los usuarios.',
        'Mejorar la organización semanal de las encuestas para contribuir al cumplimiento oportuno de la meta establecida.',
        'Continuar con la revisión y depuración de las bases antes de consolidar los resultados del informe mensual.',
        'Verificar que las observaciones permitan distinguir claramente entre encuestas acompañadas y encuestas complementarias.'
      ];
      return { strengths, opportunities, management, conclusion, cases: data.below.map(caseAnalysis), comparison: comparisonData, history };
    }

    function reportSnapshot() {
      const reportRows = currentReportRows();
      const previousRows = previousReportRows(reportRows);
      const data = metrics(reportRows);
      const narratives = reportNarratives(reportRows, previousRows);
      return {
        metadata: {
          generatedAt: new Date().toISOString(),
          sourceFile: state.fileName,
          previousFile: state.previousFileName,
          sheet: state.sheetName,
          previousSheet: state.previousSheetName,
          period: monthYearLabel(reportRows),
          dateRange: dateRangeLabel(reportRows),
          filter: { period: $('period').value, year: $('year').value, month: $('month').value, quarter: $('quarter').value, dependency: $('dependency').value }
        },
        summary: { total: data.total, average: data.average, belowFive: data.below.length, perfect: data.perfect, miguelAngel: data.miguel },
        byService: data.services.map(item => ({ service: item.key, count: item.count, participation: item.count / data.total, average: item.average, miguelAngel: item.miguel })),
        byDependency: data.deps.map(item => ({ dependency: item.key, count: item.count, participation: item.count / data.total, average: item.average })),
        serviceDependencyMatrix: { services: data.serviceNames, rows: data.matrix },
        byCriterion: data.criteria.map(item => ({ criterion: item.label, average: item.average })),
        scoreDistribution: [{ label: '5,00', count: data.perfect, participation: data.perfect / data.total }, { label: 'Inferior a 5,00', count: data.below.length, participation: data.below.length / data.total }],
        quality: data.quality,
        validation: state.validation,
        cases: data.below.map(row => ({ id: row.id, sourceRow: row.row, date: row.date ? row.date.toISOString() : null, dependency: row.dependency, service: row.service, ratings: row.ratings, average: row.average, suggestion: row.suggestion })),
        narratives,
        comparison: narratives.comparison,
        historicalComparison: narratives.history,
        records: reportRows.map(row => ({ id: row.id, sourceRow: row.sourceRow, recordKey: row.recordKey, date: row.date ? row.date.toISOString() : null, year: row.year, month: row.month, service: row.service, dependency: row.dependency, respondent: row.respondent, ratings: row.ratings, average: row.average, miguelAngel: row.miguel, suggestion: row.suggestion }))
      };
    }

    function csvCell(value) {
      const text = value === null || value === undefined ? '' : String(value);
      return `"${text.replace(/"/g, '""')}"`;
    }

    $('exportData').addEventListener('click', () => {
      downloadBlob(new Blob(['\ufeff', JSON.stringify(reportSnapshot(), null, 2)], { type: 'application/json;charset=utf-8' }), 'analisis-encuestas.json');
    });

    $('exportCsv').addEventListener('click', () => {
      const columns = ['id', 'sourceRow', 'date', 'service', 'dependency', 'respondent', 'solution', 'timeliness', 'attitude', 'clarity', 'average', 'miguelAngel', 'suggestion'];
      const lines = [columns.map(csvCell).join(';')];
      reportSnapshot().records.forEach(record => lines.push([
        record.id, record.sourceRow, record.date, record.service, record.dependency, record.respondent,
        record.ratings.solution, record.ratings.timeliness, record.ratings.attitude, record.ratings.clarity,
        record.average, record.miguelAngel, record.suggestion
      ].map(csvCell).join(';')));
      downloadBlob(new Blob(['\ufeff', lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'detalle-encuestas.csv');
    });

    $('exportWord').addEventListener('click', async () => {
      try {
        const snapshot = reportSnapshot();
        if (!snapshot.records.length || !window.docx) return;
      const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ImageRun, AlignmentType } = window.docx;
      const cell = (value, bold = false, header = false) => new TableCell({
        shading: header ? { fill: 'A9C9D8' } : undefined,
        children: [new Paragraph({ children: [new TextRun({ text: String(value ?? ''), bold, color: '000000', font: 'Arial', size: 18 })] })]
      });
      const headerRow = labels => new TableRow({ children: labels.map(label => cell(label, true, true)) });
      const table = (headers, rows) => new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow(headers), ...rows.map(row => new TableRow({ children: row.map(value => cell(value)) }))] });
      const sectionTitle = text => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, bold: true, font: 'Arial', size: 22 })] });
      const paragraph = text => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 20 })], spacing: { after: 120 } });
      const logoBytes = window.ALCALDIA_LOGO_BASE64 ? Uint8Array.from(atob(window.ALCALDIA_LOGO_BASE64), char => char.charCodeAt(0)) : null;
      const reportRows = currentReportRows();
      const data = metrics(reportRows);
      const serviceTableRows = data.services.map(item => [item.key, item.count, pct(item.count / data.total), `${item.average.toFixed(2)} / 5,00`, item.miguel]);
      serviceTableRows.push(['Total', data.total, '100,0 %', `${data.average.toFixed(2)} / 5,00`, data.miguel]);
      const dependencyTableRows = data.deps.map(item => [item.key, item.count, pct(item.count / data.total), item.average.toFixed(2)]);
      dependencyTableRows.push(['Total', data.total, '100,0 %', data.average.toFixed(2)]);
      const matrixRows = data.matrix.map(row => [row.dependency, ...row.values, row.total]);
      matrixRows.push(['Total', ...data.serviceNames.map(service => reportRows.filter(item => item.service === service).length), data.total]);
      const criteriaRows = data.criteria.map(item => [item.label, `${item.average.toFixed(2)} / 5,00`]);
      const scoreRows = [[data.perfect.toFixed(0) === '0' ? '5,00' : '5,00', data.perfect, pct(data.perfect / data.total)], ['Inferior a 5,00', data.below.length, pct(data.below.length / data.total)], ['Total', data.total, '100,0 %']];
      const caseRows = data.below.map(row => [row.id, row.date ? row.date.toLocaleDateString('es-CO') : 'Sin fecha', row.dependency, row.service, row.ratings.solution ?? '', row.ratings.timeliness ?? '', row.ratings.attitude ?? '', row.ratings.clarity ?? '', row.average.toFixed(2)]);
      const comparisonHeaders = snapshot.comparison ? ['Indicador', snapshot.comparison.previousPeriod, snapshot.metadata.period, 'Variación'] : [];
      const comparisonRowsLegacy = snapshot.comparison ? [
        ['Total de encuestas', snapshot.comparison.total.previous, snapshot.comparison.total.current, snapshot.comparison.total.difference === 0 ? 'Se mantiene' : `${snapshot.comparison.total.difference > 0 ? '+' : ''}${snapshot.comparison.total.difference}`],
        ['Calificación promedio general', `${snapshot.comparison.average.previous.toFixed(2).replace('.', ',')} / 5,00`, `${snapshot.comparison.average.current.toFixed(2).replace('.', ',')} / 5,00`, Math.abs(snapshot.comparison.average.difference) < 0.005 ? 'Se mantiene' : `${snapshot.comparison.average.difference > 0 ? '+' : ''}${snapshot.comparison.average.difference.toFixed(2).replace('.', ',')}`],
        ['Casos con calificación inferior a 5,00', snapshot.comparison.belowFive.previous, snapshot.comparison.belowFive.current, snapshot.comparison.belowFive.difference === 0 ? 'Se mantiene' : `${snapshot.comparison.belowFive.difference > 0 ? '+' : ''}${snapshot.comparison.belowFive.difference} casos`],
        ['Dependencia con mayor participación', `${snapshot.comparison.dependency.previous} (${snapshot.comparison.dependency.previousCount})`, `${snapshot.comparison.dependency.current} (${snapshot.comparison.dependency.currentCount})`, snapshot.comparison.dependency.currentCount - snapshot.comparison.dependency.previousCount === 0 ? 'Se mantiene' : `${snapshot.comparison.dependency.currentCount - snapshot.comparison.dependency.previousCount} encuestas`],
        ['Servicio con mayor participación', `${snapshot.comparison.service.previous} (${snapshot.comparison.service.previousCount})`, `${snapshot.comparison.service.current} (${snapshot.comparison.service.currentCount})`, snapshot.comparison.service.currentCount === snapshot.comparison.service.previousCount ? 'Se mantiene' : snapshot.comparison.service.currentCount > snapshot.comparison.service.previousCount ? 'Aumentó' : 'Disminuyó']
      ] : [];
      const children = [];
      if (logoBytes) children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: logoBytes, transformation: { width: 185, height: 80 } })] }));
      children.push(
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'INFORME GENERAL DE ENCUESTAS DE SATISFACCIÓN', bold: true, font: 'Arial', size: 28 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `MES DE ${snapshot.metadata.period}`, bold: true, font: 'Arial', size: 22 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Elaborado por: Miguel Ángel Correa Cano', bold: true, font: 'Arial', size: 20 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Periodo evaluado: ${snapshot.metadata.dateRange}`, font: 'Arial', size: 20 })] }),
        paragraph(`Los resultados presentados corresponden al consolidado de encuestas de satisfacción del periodo ${snapshot.metadata.period.toLowerCase()}.`),
        table(['Indicador', 'Valor'], [
          ['Total de encuestas analizadas', data.total],
          ['Calificación promedio general', `${data.average.toFixed(2)} / 5,00`],
          ['Pregunta con menor promedio', `${data.criteria.slice().sort((a, b) => a.average - b.average)[0].label} (${data.criteria.slice().sort((a, b) => a.average - b.average)[0].average.toFixed(2)})`],
          ['Servicio con mayor participación', `${data.services[0]?.key || 'Sin dato'} - ${data.services[0]?.count || 0} encuestas (${pct((data.services[0]?.count || 0) / data.total)})`],
          ['Dependencia con mayor participación', `${data.deps[0]?.key || 'Sin dato'} - ${data.deps[0]?.count || 0} encuestas (${pct((data.deps[0]?.count || 0) / data.total)})`],
          ['Encuestas con acompañamiento', `${data.miguel} (${pct(data.miguel / data.total)})`],
          ['Encuestas complementarias', `${data.total - data.miguel} (${pct((data.total - data.miguel) / data.total)})`]
        ]),
        sectionTitle('1. Total de encuestas del mes'),
        paragraph(`Durante ${snapshot.metadata.period.toLowerCase()} se registraron ${data.total} encuestas de satisfacción correspondientes a los diferentes servicios atendidos.`),
        paragraph(`El resultado general fue de ${data.average.toFixed(2).replace('.', ',')} sobre 5,00, lo que refleja una percepción ${data.average >= 4.8 ? 'ampliamente favorable' : 'que requiere seguimiento'} por parte de los usuarios encuestados.`),
        paragraph(`De las ${data.total} encuestas, ${data.miguel} fueron identificadas con acompañamiento de Miguel Ángel Correa, equivalentes al ${pct(data.miguel / data.total)}; las ${data.total - data.miguel} restantes corresponden a encuestas complementarias.`),
        sectionTitle('2. Distribución de encuestas por servicio evaluado'),
        paragraph(`El servicio con mayor participación fue ${data.services[0]?.key || 'Sin dato'}.`),
        table(['Servicio evaluado', 'Encuestas', 'Participación', 'Calificación promedio', 'Acompañamiento'], serviceTableRows),
        sectionTitle('3. Distribución por dependencia'),
        paragraph(`La dependencia con mayor participación fue ${data.deps[0]?.key || 'Sin dato'}, con ${data.deps[0]?.count || 0} encuestas, equivalente al ${pct((data.deps[0]?.count || 0) / data.total)} del total.`),
        table(['Dependencia', 'Encuestas', 'Participación', 'Calificación promedio'], dependencyTableRows),
        sectionTitle('4. Detalle por dependencia y servicio evaluado'),
        paragraph('Cada fila corresponde a una dependencia y cada columna a uno de los servicios evaluados. La columna Total suma los servicios de la dependencia.'),
        table(['Dependencia', ...data.serviceNames, 'Total'], matrixRows),
        sectionTitle('5. Calificación promedio general'),
        table(['Criterio evaluado', 'Promedio'], criteriaRows),
        sectionTitle('6. Análisis de calificaciones obtenidas'),
        paragraph(`De las ${data.total} encuestas registradas, ${data.perfect} (${pct(data.perfect / data.total)}) obtuvieron calificación máxima. Los resultados inferiores a 5,00 fueron ${data.below.length ? 'puntuales y requieren seguimiento específico' : 'inexistentes'}.`),
        table(['Promedio final por encuesta', 'Encuestas', 'Participación'], scoreRows),
        sectionTitle('Casos con calificaciones inferiores en alguno de los criterios:'),
        data.below.length ? table(['ID', 'Fecha', 'Dependencia', 'Servicio', 'Solución', 'Oportunidad', 'Actitud', 'Claridad', 'Promedio'], caseRows) : paragraph('No se identificaron casos con calificaciones inferiores a 5,00.'),
        ...snapshot.narratives.cases.map(text => paragraph(text)),
        sectionTitle('7. Comparativo respecto al mes anterior'),
        snapshot.comparison
          ? table(comparisonHeaders, comparisonRowsLegacy)
          : paragraph('No se encontraron meses anteriores para realizar el comparativo.'),
        sectionTitle('8. Observaciones relevantes'),
        new Paragraph({ children: [new TextRun({ text: 'Fortalezas identificadas:', bold: true, font: 'Arial', size: 20 })] }),
        ...snapshot.narratives.strengths.map(text => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 20 })], bullet: { level: 0 } })),
        new Paragraph({ children: [new TextRun({ text: 'Aspectos destacados para fortalecer:', bold: true, font: 'Arial', size: 20 })] }),
        ...snapshot.narratives.opportunities.map(text => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 20 })], bullet: { level: 0 } })),
        new Paragraph({ children: [new TextRun({ text: 'Aspectos por mejorar en la gestión de aplicación de encuestas a cargo de Miguel Ángel Correa Cano:', bold: true, font: 'Arial', size: 20 })] }),
        ...snapshot.narratives.management.map(text => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 20 })], bullet: { level: 0 } })),
        sectionTitle('9. Conclusión general'),
        paragraph(snapshot.narratives.conclusion),
      );
      const document = new Document({
        styles: {
          default: { document: { run: { font: 'Arial', size: 20 } }, paragraph: { run: { font: 'Arial', size: 20 } } },
          paragraphStyles: [
            { id: 'Normal', name: 'Normal', basedOn: 'Normal', next: 'Normal', run: { font: 'Arial', size: 20 } },
            { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', run: { font: 'Arial', bold: true, size: 22 } }
          ]
        },
        sections: [{ properties: {}, children }]
      });
        const blob = await Packer.toBlob(document);
        downloadBlob(blob, `informe-encuestas-${snapshot.metadata.period.toLowerCase().replace(/\s+/g, '-')}.docx`);
      } catch (error) {
        logError('Error al generar el informe Word', error);
        $('status').className = 'status error';
        $('status').textContent = error.message || 'No fue posible generar el informe Word.';
      }
    });