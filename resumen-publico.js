(() => {
  const $ = id => document.getElementById(id);
  const criteriaLabels = {
    solution: 'Solución',
    timeliness: 'Oportunidad',
    attitude: 'Actitud',
    clarity: 'Claridad'
  };
  const charts = {};
  let summary = null;

  function setStatus(message, isError = false) {
    $('status').className = isError ? 'status error' : 'status';
    $('status').textContent = message;
  }

  function formatCount(value) {
    return new Intl.NumberFormat('es-CO').format(value);
  }

  function formatScore(value) {
    return value === null || value === undefined ? 'Protegido' : value.toFixed(2).replace('.', ',');
  }

  function renderChart(id, config) {
    if (charts[id]) charts[id].destroy();
    const canvas = $(id);
    if (canvas && config) charts[id] = new Chart(canvas, config);
  }

  function categoryChart(id, noteId, entries) {
    if (!entries) {
      if (charts[id]) charts[id].destroy();
      $(noteId).textContent = 'Distribución omitida para evitar mostrar grupos pequeños.';
      return;
    }
    $(noteId).textContent = '';
    renderChart(id, {
      type: 'bar',
      data: {
        labels: entries.map(entry => entry.name),
        datasets: [{ label: 'Encuestas', data: entries.map(entry => entry.count), backgroundColor: '#2563eb' }]
      },
      options: { indexAxis: 'y', maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
  }

  function renderPeriod() {
    const period = summary.periods[Number($('periodSelect').value)];
    if (!period) return;
    $('total').textContent = formatCount(period.count);
    $('average').textContent = formatScore(period.average);
    $('belowFive').textContent = period.below_five === null ? 'Protegido' : formatCount(period.below_five);

    const criteria = Object.entries(criteriaLabels);
    renderChart('criteriaChart', {
      type: 'bar',
      data: {
        labels: criteria.map(([, label]) => label),
        datasets: [{
          label: 'Promedio',
          data: criteria.map(([key]) => period.criteria[key]),
          backgroundColor: '#0f766e'
        }]
      },
      options: { maintainAspectRatio: false, scales: { y: { min: 0, max: 5 } }, plugins: { legend: { display: false } } }
    });
    categoryChart('serviceChart', 'serviceNote', period.services);
    categoryChart('dependencyChart', 'dependencyNote', period.dependencies);
  }

  function showSummary(data) {
    summary = data;
    $('threshold').textContent = formatCount(data.privacy_threshold);
    $('updatedAt').textContent = `Última actualización: ${new Intl.DateTimeFormat('es-CO', {
      dateStyle: 'medium',
      timeStyle: 'short'
    }).format(new Date(data.generated_at))}`;

    if (!data.available || !data.periods.length) {
      setStatus(data.message || 'No hay resultados agregados disponibles.');
      return;
    }

    const granularityLabels = {
      month: 'Mensual',
      quarter: 'Trimestral',
      year: 'Anual',
      general: 'Consolidado general'
    };
    $('granularity').textContent = `Vista ${granularityLabels[data.granularity] || 'agregada'} · grupos con al menos ${data.privacy_threshold} respuestas`;
    const select = $('periodSelect');
    select.innerHTML = data.periods.map((period, index) =>
      `<option value="${index}">${period.label}</option>`
    ).join('');
    select.value = String(data.periods.length - 1);
    select.disabled = data.periods.length < 2;
    setStatus(data.message);

    renderChart('trendChart', {
      type: 'bar',
      data: {
        labels: data.periods.map(period => period.label),
        datasets: [{ label: 'Encuestas', data: data.periods.map(period => period.count), backgroundColor: '#2563eb' }]
      },
      options: { maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
    renderPeriod();
  }

  async function loadSummary() {
    try {
      const url = new URL('data/resumen-publico.json', document.baseURI);
      url.searchParams.set('v', Date.now().toString());
      const response = await fetch(url, { cache: 'no-store' });
      if (!response.ok) throw new Error(`No se pudo cargar el resumen público (HTTP ${response.status}).`);
      showSummary(await response.json());
    } catch (error) {
      setStatus(`${error.message} Intente actualizar la página más tarde.`, true);
    }
  }

  $('periodSelect').addEventListener('change', renderPeriod);
  loadSummary();
})();
