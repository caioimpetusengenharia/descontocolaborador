(() => {
  const state = { employees: [], filename: '', targetRate: 0.1, workdayHours: 8, periodText: 'não identificado' };
  const offDatesStorageKey = 'impetus-banco-horas-adm-off-dates';
  let savedOffDates = {};
  let savedOffDatesLoadError = '';
  const byId = (id) => document.getElementById(id);
  const els = {
    fileInput: byId('file-input'),
    dropZone: byId('drop-zone'),
    notice: byId('notice'),
    periodBanner: byId('period-banner'),
    target: byId('target-input'),
    workday: byId('workday-input'),
    search: byId('search-input'),
    departmentFilter: byId('department-filter'),
    balanceFilter: byId('balance-filter'),
    employeeBody: byId('employee-body'),
    departmentBody: byId('department-body'),
    exportButton: byId('export-button')
  };

  const normalizeHeader = (value) => String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ');

  const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[char]);

  function roundHoursToMinute(hours) {
    const minutes = Math.sign(hours) * Math.round(Math.abs(hours) * 60);
    return minutes === 0 ? 0 : minutes / 60;
  }

  function parseDuration(value) {
    if (value === null || value === undefined || value === '') return { hours: null, status: 'missing' };
    if (typeof value === 'number' && Number.isFinite(value)) return { hours: roundHoursToMinute(value), status: 'valid' };
    const text = String(value).trim();
    if (!text || /^#(VALUE|N\/A|REF|DIV\/0|NAME|NUM|NULL)/i.test(text)) return { hours: null, status: 'invalid' };
    const negative = /^\s*-/.test(text);
    const clean = text.replace(/^\s*[+-]/, '');
    const match = clean.match(/^(\d+):([0-5]?\d)(?::([0-5]?\d))?$/);
    if (match) {
      const hours = Number(match[1]) + Number(match[2]) / 60 + Number(match[3] || 0) / 3600;
      return { hours: roundHoursToMinute(negative ? -hours : hours), status: 'valid' };
    }
    const number = Number(text.replace(',', '.'));
    if (Number.isFinite(number)) return { hours: roundHoursToMinute(number), status: 'valid' };
    return { hours: null, status: 'invalid' };
  }

  function parseDateValue(value) {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (Math.abs(value) < 1000) return null;
      const excelBase = new Date(Date.UTC(1899, 11, 30));
      return new Date(excelBase.getTime() + value * 86400000);
    }
    if (typeof value === 'string') {
      const text = value.trim();
      if (!text || /^#(VALUE|N\/A|REF|DIV\/0|NAME|NUM|NULL)/i.test(text)) return null;
      const simple = text.replace(/\s+/g, ' ');
      const iso = simple.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (iso) {
        const date = new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
        return date.getUTCFullYear() === Number(iso[1]) && date.getUTCMonth() === Number(iso[2]) - 1 && date.getUTCDate() === Number(iso[3]) ? date : null;
      }
      const slash = simple.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (slash) {
        const dd = Number(slash[1]);
        const mm = Number(slash[2]);
        const yy = Number(slash[3]);
        const year = yy < 100 ? 2000 + yy : yy;
        const date = new Date(Date.UTC(year, mm - 1, dd));
        if (date.getUTCFullYear() === year && date.getUTCMonth() === mm - 1 && date.getUTCDate() === dd) return date;
      }
    }
    return null;
  }

  function formatDate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return 'não identificado';
    return new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' }).format(date);
  }

  function detectPeriod(rows) {
    const headers = rows[0] || [];
    const dateColumns = headers.reduce((indexes, header, index) => {
      const normalized = normalizeHeader(header);
      const personalDate = /\b(admissao|nascimento|cadastro|desligamento)\b/.test(normalized);
      if (!personalDate && (/\bperiodo\b/.test(normalized) || /\b(inicio|inicial|fim|final)\b/.test(normalized) || /^(data|date)(\s|$)/.test(normalized))) indexes.push(index);
      return indexes;
    }, []);
    const candidates = [];
    for (const row of rows.slice(1)) {
      for (const index of dateColumns) {
        const cell = getCell(row, index);
        const parsed = parseDateValue(cell);
        if (parsed) candidates.push(parsed);
      }
    }

    if (candidates.length === 0) return 'não identificado';

    const sorted = [...new Set(candidates.map((date) => date.getTime()))]
      .map((time) => new Date(time))
      .sort((a, b) => a.getTime() - b.getTime());

    if (sorted.length === 1) return formatDate(sorted[0]);
    return `${formatDate(sorted[0])} a ${formatDate(sorted[sorted.length - 1])}`;
  }

  function isExcelDurationFormat(format) {
    return /\[(?:h+|m+|s+)\]/i.test(format) || /h{1,2}:m{1,2}/i.test(format) || /m{1,2}:s{1,2}/i.test(format);
  }

  function worksheetToRows(sheet) {
    if (!sheet['!ref']) return [];
    const range = window.XLSX.utils.decode_range(sheet['!ref']);
    const rows = [];
    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
      const row = [];
      for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex++) {
        const cell = sheet[window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
        let value = cell ? cell.v : null;
        if (cell && cell.t === 'e') value = cell.w || '#VALUE!';
        else if (cell && typeof value === 'number' && isExcelDurationFormat(cell.z || '')) value *= 24;
        row.push(value);
      }
      rows.push(row);
    }
    return rows;
  }

  function formatDuration(hours, signed = false) {
    if (!Number.isFinite(hours)) return '—';
    const totalMinutes = Math.round(Math.abs(hours) * 60);
    const prefix = signed && hours < 0 ? '−' : '';
    return `${prefix}${Math.floor(totalMinutes / 60)}:${String(totalMinutes % 60).padStart(2, '0')}`;
  }

  function formatHoursAndMinutes(hours) {
    const totalMinutes = Math.round(hours * 60);
    const wholeHours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${wholeHours}h${minutes ? ` ${minutes}min` : ''}`;
  }

  function formatDaysOff(hours) {
    if (!Number.isFinite(hours) || hours < 0 || !Number.isFinite(state.workdayHours) || state.workdayHours <= 0) return '—';
    const totalMinutes = Math.round(hours * 60);
    const workdayMinutes = Math.round(state.workdayHours * 60);
    if (totalMinutes === 0) return '0h';
    const wholeDays = Math.floor(totalMinutes / workdayMinutes);
    const remainingMinutes = totalMinutes % workdayMinutes;
    const dayLabel = `${wholeDays} ${wholeDays === 1 ? 'dia' : 'dias'}`;
    if (wholeDays === 0) return formatHoursAndMinutes(hours);
    return remainingMinutes ? `${dayLabel} + ${formatHoursAndMinutes(remainingMinutes / 60)}` : dayLabel;
  }

  function compensationHours(balance) {
    if (!Number.isFinite(balance) || balance <= 0) return 0;
    return roundHoursToMinute(balance * state.targetRate);
  }

  function showNotice(message, type) {
    els.notice.textContent = message;
    els.notice.className = `notice ${type}`;
    els.notice.hidden = false;
  }

  function loadSavedOffDates() {
    try {
      const saved = window.localStorage.getItem(offDatesStorageKey);
      if (!saved) return {};
      const parsed = JSON.parse(saved);
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw new Error('Formato de dados inválido.');
      for (const dates of Object.values(parsed)) {
        if (!Array.isArray(dates) || dates.some((date) => typeof date !== 'string' || !parseDateValue(date))) {
          throw new Error('A lista de datas salvas contém valores inválidos.');
        }
      }
      return parsed;
    } catch (error) {
      savedOffDatesLoadError = `Não foi possível carregar as datas salvas: ${error.message}`;
      showNotice(savedOffDatesLoadError, 'error');
      return {};
    }
  }

  function saveEmployeeOffDates(person, dates) {
    const updated = { ...savedOffDates };
    if (dates.length) updated[person.storageKey] = dates;
    else delete updated[person.storageKey];
    try {
      window.localStorage.setItem(offDatesStorageKey, JSON.stringify(updated));
    } catch (error) {
      showNotice(`Não foi possível salvar as datas no armazenamento local: ${error.message}`, 'error');
      return false;
    }
    savedOffDates = updated;
    person.offDates = dates;
    savedOffDatesLoadError = '';
    return true;
  }

  function getCell(row, index) {
    return index >= 0 && index < row.length ? row[index] : null;
  }

  function findColumn(headers, candidates, required = true) {
    const normalized = headers.map(normalizeHeader);
    for (const candidate of candidates) {
      const index = normalized.indexOf(normalizeHeader(candidate));
      if (index >= 0) return index;
    }
    if (required) throw new Error(`Não encontrei a coluna “${candidates[0]}”. Confira se a planilha mantém o formato esperado.`);
    return -1;
  }

  function parseRows(rows) {
    if (!rows.length) throw new Error('A planilha não contém linhas para importar.');
    const headers = rows[0] || [];
    const columns = {
      name: findColumn(headers, ['Nome do funcionário', 'Nome do funcionario']),
      registration: findColumn(headers, ['Número de matrícula', 'Numero de matricula', 'Matrícula', 'Matricula'], false),
      job: findColumn(headers, ['Nome do cargo', 'Cargo'], false),
      department: findColumn(headers, ['Nome do departamento', 'Departamento', 'Polo'], false),
      balance: findColumn(headers, ['BH Saldo', 'Saldo BH', 'Banco de horas saldo']),
      extra: findColumn(headers, ['Extra_50%', 'Extra 50%', 'Extra_50'], false)
    };
    const employees = [];
    for (const row of rows.slice(1)) {
      const name = String(getCell(row, columns.name) ?? '').trim();
      if (!name) continue;
      const balance = parseDuration(getCell(row, columns.balance));
      const extra = columns.extra < 0 ? { hours: null, status: 'missing' } : parseDuration(getCell(row, columns.extra));
      const registration = String(getCell(row, columns.registration) ?? '').trim();
      const storageKey = JSON.stringify([normalizeHeader(registration), normalizeHeader(name)]);
      employees.push({
        id: employees.length,
        name,
        registration,
        storageKey,
        job: String(getCell(row, columns.job) ?? '').trim(),
        department: String(getCell(row, columns.department) ?? '').trim() || 'Sem departamento',
        balance: balance.hours,
        balanceStatus: balance.status,
        extra: extra.hours,
        extraStatus: extra.status,
        offDates: [...(savedOffDates[storageKey] || [])]
      });
    }
    if (!employees.length) throw new Error('Não encontrei colaboradores na planilha.');
    return employees;
  }

  async function readFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    let rows;
    if (extension === 'csv') {
      rows = parseCsv(await file.text());
    } else {
      if (!window.XLSX) throw new Error('Não foi possível carregar o leitor de Excel. Verifique sua conexão e atualize a página.');
      const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false, cellNF: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error('Não encontrei uma aba com dados no arquivo.');
      rows = worksheetToRows(sheet);
    }
    state.periodText = detectPeriod(rows);
    state.employees = parseRows(rows);
    state.filename = file.name;
    els.periodBanner.textContent = `Período da planilha: ${state.periodText}`;
    render();
    byId('upload-title').textContent = file.name;
    byId('upload-description').textContent = `${state.employees.length} colaboradores carregados. Selecione outro arquivo para atualizar os dados.`;
    if (savedOffDatesLoadError) showNotice(savedOffDatesLoadError, 'error');
    else showNotice(`Planilha “${file.name}” carregada com sucesso. Os dados permanecem neste navegador.`, 'success');
  }

  function parseCsv(text) {
    const firstLine = text.split(/\r?\n/, 1)[0] || '';
    const delimiter = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [];
    let field = '';
    let quoted = false;
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (char === '"') {
        if (quoted && text[i + 1] === '"') { field += '"'; i++; }
        else quoted = !quoted;
      } else if (char === delimiter && !quoted) {
        row.push(field); field = '';
      } else if ((char === '\n' || char === '\r') && !quoted) {
        if (char === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        if (row.some((value) => value.trim())) rows.push(row);
        row = [];
      } else {
        field += char;
      }
    }
    row.push(field);
    if (row.some((value) => value.trim())) rows.push(row);
    return rows;
  }

  function getStats(employees = state.employees) {
    const positive = employees.filter((person) => person.balanceStatus === 'valid' && person.balance > 0);
    const negative = employees.filter((person) => person.balanceStatus === 'valid' && person.balance < 0);
    const invalid = employees.filter((person) => person.balanceStatus === 'invalid' || person.balanceStatus === 'missing');
    const positiveHours = roundHoursToMinute(positive.reduce((sum, person) => sum + person.balance, 0));
    const targetHours = positive.reduce((sum, person) => sum + compensationHours(person.balance), 0);
    const negativeHours = roundHoursToMinute(negative.reduce((sum, person) => sum + person.balance, 0));
    const extraPeople = employees.filter((person) => person.extraStatus === 'valid' && Number.isFinite(person.extra));
    const extraHours = roundHoursToMinute(extraPeople.reduce((sum, person) => sum + person.extra, 0));
    return { positive, negative, invalid, positiveHours, targetHours, negativeHours, extraPeople, extraHours };
  }

  function renderMetrics() {
    const stats = getStats();
    byId('metric-positive').textContent = formatDuration(stats.positiveHours);
    byId('metric-positive-count').textContent = `${stats.positive.length} colaboradores com saldo positivo`;
    byId('metric-target').textContent = formatDuration(stats.targetHours);
    byId('metric-target-days').textContent = `${formatDaysOff(stats.targetHours)} de ${state.workdayHours}h`;
    byId('metric-projected').textContent = formatDuration(roundHoursToMinute(stats.positiveHours - stats.targetHours));
    byId('metric-extra').textContent = formatDuration(stats.extraHours, true);
    byId('metric-extra-count').textContent = `${stats.extraPeople.length} registros válidos de hora extra`;
    byId('overview-total').textContent = state.employees.length;
    byId('overview-positive').textContent = `${stats.positive.length} colaboradores`;
    byId('overview-negative').textContent = `${stats.negative.length} colaboradores · ${formatDuration(stats.negativeHours, true)}`;
    byId('overview-invalid').textContent = `${stats.invalid.length} colaboradores`;
    byId('quality-note').textContent = stats.invalid.length
      ? `${stats.invalid.length} registro(s) sem saldo ou com erro no arquivo; não entram nos totais e devem ser conferidos.`
      : 'Os valores da planilha são processados localmente no navegador.';
  }

  function renderDepartments() {
    const groups = new Map();
    for (const person of state.employees) {
      if (!groups.has(person.department)) groups.set(person.department, { positivePeople: 0, positive: 0, target: 0 });
      const group = groups.get(person.department);
      if (person.balanceStatus === 'valid' && person.balance > 0) {
        group.positivePeople++;
        group.positive = roundHoursToMinute(group.positive + person.balance);
        group.target += compensationHours(person.balance);
      }
    }
    const sorted = [...groups.entries()].sort((a, b) => b[1].positive - a[1].positive || a[0].localeCompare(b[0], 'pt-BR'));
    byId('department-count').textContent = `${sorted.length} ${sorted.length === 1 ? 'departamento' : 'departamentos'}`;
    els.departmentBody.innerHTML = sorted.length ? sorted.map(([department, data]) => `
      <tr><td>${escapeHtml(department)}</td><td>${data.positivePeople}</td><td class="table-number">${formatDuration(data.positive)}</td><td class="table-number">${formatDuration(data.target)}</td></tr>
    `).join('') : '<tr><td class="empty-state" colspan="4">Importe uma planilha para visualizar o resumo.</td></tr>';
  }

  function filteredEmployees() {
    const query = els.search.value.trim().toLocaleLowerCase('pt-BR');
    const department = els.departmentFilter.value;
    const balanceType = els.balanceFilter.value;
    return state.employees.filter((person) => {
      if (department && person.department !== department) return false;
      if (query && !`${person.name} ${person.registration}`.toLocaleLowerCase('pt-BR').includes(query)) return false;
      if (!balanceType) return true;
      if (balanceType === 'invalid') return person.balanceStatus === 'invalid' || person.balanceStatus === 'missing';
      if (person.balanceStatus !== 'valid') return false;
      if (balanceType === 'positive') return person.balance > 0;
      if (balanceType === 'negative') return person.balance < 0;
      return person.balance === 0;
    });
  }

  function renderEmployees() {
    const filtered = filteredEmployees();
    const body = filtered.map((person) => {
      const balanceClass = person.balanceStatus !== 'valid' ? 'balance-invalid' : person.balance < 0 ? 'balance-negative' : person.balance > 0 ? 'balance-positive' : '';
      const balanceText = person.balanceStatus === 'invalid' ? 'Erro na planilha' : person.balanceStatus === 'missing' ? 'Sem informação' : formatDuration(person.balance, true);
      const plannedHours = person.balanceStatus === 'valid' && person.balance > 0 ? compensationHours(person.balance) : null;
      const planned = plannedHours === null ? '—' : formatDuration(plannedHours);
      const daysOff = plannedHours === null ? '—' : formatDaysOff(plannedHours);
      const offDates = person.offDates.map((date) => {
        const label = formatDate(parseDateValue(date));
        return `<span class="off-date-chip">${escapeHtml(label)}<button class="off-date-remove" type="button" data-action="remove-off-date" data-date="${escapeHtml(date)}" aria-label="Remover data ${escapeHtml(label)}">×</button></span>`;
      }).join('');
      const extra = person.extraStatus === 'valid' ? formatDuration(person.extra, true) : person.extraStatus === 'invalid' ? '<span class="balance-invalid">Conferir</span>' : '—';
      return `<tr>
        <td class="employee-name">${escapeHtml(person.name)}</td>
        <td>${escapeHtml(person.registration || '—')}</td>
        <td class="employee-job" title="${escapeHtml(person.job)}">${escapeHtml(person.job || '—')}</td>
        <td>${escapeHtml(person.department)}</td>
        <td class="${balanceClass} table-number">${balanceText}</td>
        <td class="table-number">${planned}</td>
        <td class="table-number">${daysOff}</td>
        <td class="off-date-cell" data-employee-index="${person.id}">
          <div class="off-date-entry"><input class="off-date-input" type="date" aria-label="Nova data de folga para ${escapeHtml(person.name)}"><button class="off-date-add" type="button" data-action="add-off-date">Adicionar</button></div>
          <div class="off-date-list">${offDates}</div>
        </td>
        <td class="table-number">${extra}</td>
      </tr>`;
    });
    els.employeeBody.innerHTML = body.length ? body.join('') : '<tr><td class="empty-state" colspan="9">Nenhum colaborador corresponde aos filtros.</td></tr>';
    byId('row-count').textContent = `${filtered.length} de ${state.employees.length} colaboradores`;
  }

  function updateDepartmentFilter() {
    const selected = els.departmentFilter.value;
    const departments = [...new Set(state.employees.map((person) => person.department))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    els.departmentFilter.innerHTML = '<option value="">Todos os departamentos</option>' + departments.map((department) => `<option value="${escapeHtml(department)}">${escapeHtml(department)}</option>`).join('');
    if (departments.includes(selected)) els.departmentFilter.value = selected;
  }

  function render() {
    renderMetrics();
    renderDepartments();
    updateDepartmentFilter();
    renderEmployees();
    els.exportButton.disabled = state.employees.length === 0;
  }

  function exportCsv() {
    const rows = [['Nome do funcionário', 'Matrícula', 'Cargo', 'Departamento', 'BH Saldo', 'Meta a compensar', 'Dias de folga', 'Datas possíveis para folga', 'Extra_50%']];
    for (const person of filteredEmployees()) {
      rows.push([
        person.name, person.registration, person.job, person.department,
        person.balanceStatus === 'valid' ? formatDuration(person.balance, true) : person.balanceStatus === 'invalid' ? 'Erro' : '',
        person.balanceStatus === 'valid' && person.balance > 0 ? formatDuration(compensationHours(person.balance)) : '',
        person.balanceStatus === 'valid' && person.balance > 0 ? formatDaysOff(compensationHours(person.balance)) : '',
        person.offDates.map((date) => formatDate(parseDateValue(date))).join(' | '),
        person.extraStatus === 'valid' ? formatDuration(person.extra, true) : person.extraStatus === 'invalid' ? 'Conferir' : ''
      ]);
    }
    const csv = '\uFEFF' + rows.map((row) => row.map((field) => `"${String(field).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'impetus-banco-horas-adm.csv';
    link.click();
    URL.revokeObjectURL(url);
  }

  savedOffDates = loadSavedOffDates();
  byId('today').textContent = new Intl.DateTimeFormat('pt-BR', { dateStyle: 'medium' }).format(new Date());
  byId('upload-trigger').addEventListener('click', () => els.fileInput.click());
  byId('browse-trigger').addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', async (event) => {
    const [file] = event.target.files;
    if (!file) return;
    try { await readFile(file); }
    catch (error) { showNotice(error.message || 'Não foi possível processar a planilha.', 'error'); }
    finally { event.target.value = ''; }
  });
  els.dropZone.addEventListener('dragover', (event) => { event.preventDefault(); els.dropZone.classList.add('drag-over'); });
  els.dropZone.addEventListener('dragleave', (event) => {
    if (!els.dropZone.contains(event.relatedTarget)) els.dropZone.classList.remove('drag-over');
  });
  els.dropZone.addEventListener('drop', async (event) => {
    event.preventDefault();
    els.dropZone.classList.remove('drag-over');
    const [file] = event.dataTransfer.files;
    if (!file) return;
    try { await readFile(file); }
    catch (error) { showNotice(error.message || 'Não foi possível processar a planilha.', 'error'); }
  });
  els.target.addEventListener('input', () => {
    const value = Number(els.target.value);
    if (!Number.isFinite(value) || value < 0 || value > 100) return;
    state.targetRate = value / 100;
    render();
  });
  els.workday.addEventListener('input', () => {
    if (els.workday.value === '') return;
    const value = Number(els.workday.value);
    if (!Number.isFinite(value) || value <= 0 || value > 24) return;
    state.workdayHours = value;
    render();
  });
  els.search.addEventListener('input', renderEmployees);
  els.departmentFilter.addEventListener('change', renderEmployees);
  els.balanceFilter.addEventListener('change', renderEmployees);
  els.employeeBody.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const cell = button.closest('.off-date-cell');
    const person = cell && state.employees[Number(cell.dataset.employeeIndex)];
    if (!person) return;

    if (button.dataset.action === 'add-off-date') {
      const input = cell.querySelector('.off-date-input');
      const date = input.value;
      if (!parseDateValue(date)) {
        showNotice('Selecione uma data válida para adicionar à lista de folgas.', 'error');
        input.focus();
        return;
      }
      if (person.offDates.includes(date)) {
        showNotice('Essa data já está na lista de folgas deste colaborador.', 'error');
        return;
      }
      const dates = [...person.offDates, date].sort();
      if (saveEmployeeOffDates(person, dates)) {
        renderEmployees();
        showNotice('Data de folga salva neste navegador.', 'success');
      }
      return;
    }

    if (button.dataset.action === 'remove-off-date') {
      const dates = person.offDates.filter((date) => date !== button.dataset.date);
      if (saveEmployeeOffDates(person, dates)) {
        renderEmployees();
        showNotice('Data de folga removida do armazenamento local.', 'success');
      }
    }
  });
  els.exportButton.addEventListener('click', exportCsv);
})();
