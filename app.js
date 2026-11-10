// Cambia únicamente este valor para definir la contraseña de acceso.
    const APP_PASSWORD = 'ANGEL';
    const authenticated = sessionStorage.getItem('survey-analyzer-authenticated') === 'true';
    if (authenticated) document.body.classList.add('authenticated');
    const get = id => document.getElementById(id);
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

    const state = { rows: [], filtered: [], charts: {}, fileName: '', sheetName: '' };
    const $ = id => document.getElementById(id);
    const clean = value => String(value ?? '').replace(/\s+/g, ' ').trim();
    const norm = value => clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const number = value => {
      if (typeof value === 'number') return Number.isFinite(value) ? value : null;
      const n = Number(String(value ?? '').replace(',', '.').replace(/[^\d.-]/g, ''));
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
      service: ['Servicio evaluado'],
      dependency: ['Unidad Administrativa'],
      name: ['Nombre2', 'Nombre'],
      start: ['Hora de inicio'],
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
    function parseDate(value) {
      if (value instanceof Date && !isNaN(value)) return value;
      const d = new Date(value);
      return isNaN(d) ? null : d;
    }
    function mapRows(raw) {
      const headers = raw.length ? raw[0].map(clean) : [];
      const keys = {};
      for (const [semantic, options] of Object.entries(aliases)) keys[semantic] = findKey(headers, options);
      if (!keys.service || !keys.dependency || !keys.start) throw new Error('No se encontraron las columnas mínimas de servicio, dependencia y fecha.');
      return raw.slice(1).map((cells, index) => {
        const obj = Object.fromEntries(headers.map((h, i) => [h, cells[i]]));
        const date = parseDate(obj[keys.start]);
        const ratings = {};
        for (const item of ratingFields) ratings[item.key] = number(obj[keys[item.key]]);
        const suggestion = clean(obj[keys.suggestion]);
        const finalAverage = avg(Object.values(ratings).filter(v => v !== null));
        return {
          row: index + 2,
          id: obj.ID ?? obj.Id ?? index + 1,
          service: clean(obj[keys.service]),
          dependency: clean(obj[keys.dependency]),
          respondent: clean(obj[keys.name]),
          suggestion,
          date,
          month: date ? date.getMonth() + 1 : null,
          year: date ? date.getFullYear() : null,
          ratings,
          average: finalAverage,
          miguel: /miguel\s*[aá]ngel\s*correa|acompa[nñ]ada por miguel/i.test(suggestion)
        };
      }).filter(r => r.service || r.dependency || r.date);
    }
    function options() {
      const months = [...new Set(state.rows.map(r => r.month).filter(Boolean))].sort((a,b) => a-b);
      $('month').innerHTML = '<option value="all">Todos</option>' + months.map(m => `<option value="${m}">${new Date(2020, m-1, 1).toLocaleString('es-CO', {month:'long'})}</option>`).join('');
      const deps = [...new Set(state.rows.map(r => r.dependency).filter(Boolean))].sort((a,b) => a.localeCompare(b, 'es'));
      $('dependency').innerHTML = '<option value="all">Todas</option>' + deps.map(d => `<option value="${esc(d)}">${esc(d)}</option>`).join('');
    }
    function applyFilter() {
      const mode = $('period').value;
      const month = $('month').value;
      const quarter = $('quarter').value;
      const dependency = $('dependency').value;
      state.filtered = state.rows.filter(r => {
        const periodOk = mode === 'all' || (mode === 'month' && (month === 'all' || String(r.month) === month)) || (mode === 'quarter' && (quarter === 'all' || String(Math.ceil(r.month / 3)) === quarter));
        return periodOk && (dependency === 'all' || r.dependency === dependency);
      });
      updatePeriodLabel();
      render();
    }
    function updatePeriodLabel() {
      const mode = $('period').value;
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
        label = quarterNames[quarter];
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
      $('report').textContent = buildReport(rows, services, deps, below, miguel);
      ['exportWord', 'exportData', 'exportCsv'].forEach(id => { $(id).disabled = !total; });
    }
    function renderQuality(rows) {
      const ids = rows.map(r => String(r.id)).filter(Boolean);
      const duplicateIds = [...new Set(ids.filter((id, i) => ids.indexOf(id) !== i))];
      const missing = {
        'Fecha de inicio': rows.filter(r => !r.date).length,
        'Servicio evaluado': rows.filter(r => !r.service).length,
        'Dependencia': rows.filter(r => !r.dependency).length,
        'Criterios incompletos': rows.filter(r => Object.values(r.ratings).some(v => v === null)).length
      };
      const rowsHtml = [
        ['Registros analizados', rows.length, 'OK'],
        ['Participación por servicio', `${group(rows, 'service').reduce((a, b) => a + b.count, 0)}/${rows.length}`, '100 %'],
        ['Participación por dependencia', `${group(rows, 'dependency').reduce((a, b) => a + b.count, 0)}/${rows.length}`, '100 %'],
        ['IDs duplicados', duplicateIds.length, duplicateIds.length ? `Revisar: ${duplicateIds.slice(0, 8).join(', ')}` : 'No detectados'],
        ...Object.entries(missing).map(([label, count]) => [label, count, count ? 'Revisar' : 'OK'])
      ];
      $('qualityTable').innerHTML = `<table><thead><tr><th>Validación</th><th class="num">Resultado</th><th>Estado</th></tr></thead><tbody>${rowsHtml.map(row => `<tr><td>${esc(row[0])}</td><td class="num">${esc(row[1])}</td><td class="${String(row[2]).startsWith('OK') || row[2] === '100 %' || row[2] === 'No detectados' ? 'good' : 'risk'}">${esc(row[2])}</td></tr>`).join('')}</tbody></table>`;
    }
    function renderCharts(rows, services, deps) {
      chart('service', 'serviceChart', { type: 'doughnut', data: { labels: services.map(x => x.key), datasets: [{ data: services.map(x => x.count), backgroundColor: ['#2563eb','#0f766e','#f59e0b','#9333ea','#ef4444','#64748b'] }] }, options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } } });
      const topDeps = deps.slice(0, 10).reverse();
      chart('dependency', 'dependencyChart', { type: 'bar', data: { labels: topDeps.map(x => x.key), datasets: [{ label: 'Encuestas', data: topDeps.map(x => x.count), backgroundColor: '#0f766e', borderRadius: 5 }] }, options: { maintainAspectRatio: false, indexAxis: 'y', scales: baseScales(), plugins: { legend: { display: false } } } });
      const months = [...new Set(rows.map(r => r.month).filter(Boolean))].sort((a,b) => a-b);
      const monthly = months.map(m => rows.filter(r => r.month === m).length);
      chart('monthly', 'monthlyChart', { type: 'line', data: { labels: months.map(m => new Date(2020,m-1,1).toLocaleString('es-CO',{month:'short'})), datasets: [{ label: 'Encuestas', data: monthly, borderColor: '#2563eb', backgroundColor: '#dbeafe', fill: true, tension: .3 }] }, options: { maintainAspectRatio: false, scales: baseScales() } });
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
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          state.rows = mapRows(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true }));
          options();
          $('fileName').textContent = file.name;
          state.sheetName = workbook.SheetNames[0];
          $('fileMeta').textContent = `${state.rows.length.toLocaleString('es-CO')} registros detectados · hoja: ${state.sheetName}`;
          $('status').className = 'status';
          $('status').textContent = 'Archivo cargado y validado localmente.';
          applyFilter();
        } catch (error) {
          $('status').className = 'status error';
          $('status').textContent = error.message || 'No fue posible leer el archivo.';
        }
      };
      reader.readAsArrayBuffer(file);
    });
    $('period').addEventListener('change', () => { toggleControls(); applyFilter(); });
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

    function reportSnapshot() {
      const rows = state.filtered;
      const services = group(rows, 'service');
      const deps = group(rows, 'dependency');
      const below = rows.filter(r => r.average < 5).sort((a,b) => a.average - b.average);
      const miguel = rows.filter(r => r.miguel).length;
      return {
        metadata: {
          generatedAt: new Date().toISOString(),
          sourceFile: state.fileName,
          sheet: state.sheetName,
          filter: {
            period: $('period').value,
            month: $('month').value,
            quarter: $('quarter').value,
            dependency: $('dependency').value
          }
        },
        summary: {
          total: rows.length,
          average: avg(rows.map(r => r.average)),
          belowFive: below.length,
          miguelAngel: miguel
        },
        byService: services.map(x => ({ service: x.key, count: x.count, participation: x.count / rows.length, average: x.average, miguelAngel: x.miguel })),
        byDependency: deps.map(x => ({ dependency: x.key, count: x.count, participation: x.count / rows.length, average: x.average })),
        byCriterion: ratingFields.map(item => ({ criterion: item.label, average: avg(rows.map(r => r.ratings[item.key]).filter(v => v !== null)) })),
        quality: {
          analyzed: rows.length,
          missingDate: rows.filter(r => !r.date).length,
          missingService: rows.filter(r => !r.service).length,
          missingDependency: rows.filter(r => !r.dependency).length,
          incompleteCriteria: rows.filter(r => Object.values(r.ratings).some(v => v === null)).length
        },
        cases: below.map(r => ({
          id: r.id,
          sourceRow: r.row,
          date: r.date ? r.date.toISOString() : null,
          dependency: r.dependency,
          service: r.service,
          average: r.average,
          suggestion: r.suggestion
        })),
        records: rows.map(r => ({
          id: r.id,
          sourceRow: r.row,
          date: r.date ? r.date.toISOString() : null,
          year: r.year,
          month: r.month,
          service: r.service,
          dependency: r.dependency,
          respondent: r.respondent,
          ratings: r.ratings,
          average: r.average,
          miguelAngel: r.miguel,
          suggestion: r.suggestion
        }))
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
      reportSnapshot().records.forEach(record => {
        lines.push([
          record.id, record.sourceRow, record.date, record.service, record.dependency, record.respondent,
          record.ratings.solution, record.ratings.timeliness, record.ratings.attitude, record.ratings.clarity,
          record.average, record.miguelAngel, record.suggestion
        ].map(csvCell).join(';'));
      });
      downloadBlob(new Blob(['\ufeff', lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'detalle-encuestas.csv');
    });

    $('exportWord').addEventListener('click', async () => {
      const snapshot = reportSnapshot();
      if (!snapshot.records.length || !window.docx) return;
      const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = window.docx;
      const cell = value => new TableCell({ children: [new Paragraph(String(value ?? ''))] });
      const headerRow = labels => new TableRow({ children: labels.map(label => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: label, bold: true })] })] })) });
      const serviceRows = snapshot.byService.slice(0, 20).map(item => new TableRow({ children: [cell(item.service), cell(item.count), cell(pct(item.participation)), cell(item.average.toFixed(2))] }));
      const dependencyRows = snapshot.byDependency.slice(0, 20).map(item => new TableRow({ children: [cell(item.dependency), cell(item.count), cell(pct(item.participation)), cell(item.average.toFixed(2))] }));
      const caseRows = snapshot.cases.slice(0, 50).map(item => new TableRow({ children: [cell(item.id), cell(item.date ? new Date(item.date).toLocaleDateString('es-CO') : 'Sin fecha'), cell(item.service), cell(item.dependency), cell(item.average.toFixed(2)), cell(item.suggestion || 'Sin observación')] }));
      const reportText = buildReport(state.filtered, group(state.filtered, 'service'), group(state.filtered, 'dependency'), snapshot.cases.map(item => state.filtered.find(r => String(r.id) === String(item.id))).filter(Boolean), snapshot.summary.miguelAngel);
      const document = new Document({
        sections: [{
          children: [
            new Paragraph({ text: 'Informe de análisis de encuestas de satisfacción', heading: HeadingLevel.TITLE }),
            new Paragraph(`Mesa de Servicios · Alcaldía de Itagüí`),
            new Paragraph(`Fuente: ${state.fileName || 'No especificada'} · Hoja: ${state.sheetName || 'No especificada'}`),
            new Paragraph(`Generado: ${new Date().toLocaleString('es-CO')}`),
            new Paragraph({ text: '1. Resumen ejecutivo', heading: HeadingLevel.HEADING_1 }),
            ...reportText.split('\n\n').filter(Boolean).map(text => new Paragraph(text)),
            new Paragraph({ text: '2. Indicadores principales', heading: HeadingLevel.HEADING_1 }),
            new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
              headerRow(['Indicador', 'Resultado']),
              new TableRow({ children: [cell('Encuestas analizadas'), cell(snapshot.summary.total)] }),
              new TableRow({ children: [cell('Promedio general'), cell(snapshot.summary.average.toFixed(2))] }),
              new TableRow({ children: [cell('Casos inferiores a 5,00'), cell(snapshot.summary.belowFive)] }),
              new TableRow({ children: [cell('Casos asociados a Miguel Ángel'), cell(snapshot.summary.miguelAngel)] })
            ] }),
            new Paragraph({ text: '3. Resultados por servicio', heading: HeadingLevel.HEADING_1 }),
            new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow(['Servicio', 'Encuestas', 'Participación', 'Promedio']), ...serviceRows] }),
            new Paragraph({ text: '4. Resultados por dependencia', heading: HeadingLevel.HEADING_1 }),
            new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow(['Dependencia', 'Encuestas', 'Participación', 'Promedio']), ...dependencyRows] }),
            new Paragraph({ text: '5. Criterios evaluados', heading: HeadingLevel.HEADING_1 }),
            new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow(['Criterio', 'Promedio']), ...snapshot.byCriterion.map(item => new TableRow({ children: [cell(item.criterion), cell(item.average.toFixed(2))] }))] }),
            new Paragraph({ text: '6. Casos para seguimiento', heading: HeadingLevel.HEADING_1 }),
            caseRows.length ? new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow(['ID', 'Fecha', 'Servicio', 'Dependencia', 'Promedio', 'Observación']), ...caseRows] }) : new Paragraph('No se identificaron casos con promedio inferior a 5,00.'),
            new Paragraph({ text: '7. Control de calidad', heading: HeadingLevel.HEADING_1 }),
            new Paragraph(`Registros analizados: ${snapshot.quality.analyzed}. Fechas faltantes: ${snapshot.quality.missingDate}. Servicios faltantes: ${snapshot.quality.missingService}. Dependencias faltantes: ${snapshot.quality.missingDependency}. Registros con criterios incompletos: ${snapshot.quality.incompleteCriteria}.`),
            new Paragraph('Este informe es un apoyo para la revisión institucional y debe contrastarse con la base original antes de su aprobación.')
          ]
        }]
      });
      const blob = await Packer.toBlob(document);
      downloadBlob(blob, 'informe-encuestas.docx');
    });