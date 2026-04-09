const API_URL = 'https://script.google.com/macros/s/AKfycbzk5EJXxI9_Qkg77ZU6EOYF2imQfsL7HtN6L9o40A_6Y8f0Rsxo3RG8_MKKKHccHJgg/exec';
const SETTINGS_KEY = 'akta-bookings-settings-web-v3';

const defaultSettings = {
  labName: 'AKTA 使用预约登记',
  defaultCleanupMinutes: 30,
  notice: '请把上样、清洗、系统平衡和结束整理时间一起算进预约时段。',
};

const state = {
  bookings: [],
  settings: { ...defaultSettings },
  editingRowNumber: null,
  isLoading: false,
};

const el = {
  tabs: [...document.querySelectorAll('.tab')],
  panels: {
    book: document.getElementById('panel-book'),
    list: document.getElementById('panel-list'),
    settings: document.getElementById('panel-settings'),
  },
  bookingForm: document.getElementById('bookingForm'),
  formTitle: document.getElementById('formTitle'),
  submitBtn: document.getElementById('submitBtn'),
  resetBtn: document.getElementById('resetBtn'),
  feedback: document.getElementById('feedback'),
  searchInput: document.getElementById('searchInput'),
  filterDate: document.getElementById('filterDate'),
  clearFiltersBtn: document.getElementById('clearFiltersBtn'),
  exportBtn: document.getElementById('exportBtn'),
  listContainer: document.getElementById('listContainer'),
  settingsForm: document.getElementById('settingsForm'),
  clearAllBtn: document.getElementById('clearAllBtn'),
  noticeText: document.getElementById('noticeText'),
  statToday: document.getElementById('statToday'),
  statUpcoming: document.getElementById('statUpcoming'),
};

const fields = {
  user: document.getElementById('user'),
  date: document.getElementById('date'),
  purpose: document.getElementById('purpose'),
  start: document.getElementById('start'),
  end: document.getElementById('end'),
  cleanupMinutes: document.getElementById('cleanupMinutes'),
  contact: document.getElementById('contact'),
  notes: document.getElementById('notes'),
};

const settingFields = {
  labName: document.getElementById('labName'),
  defaultCleanupMinutes: document.getElementById('defaultCleanupMinutes'),
  notice: document.getElementById('notice'),
};

function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function saveJSON(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function nowDate() {
  return new Date().toISOString().slice(0, 10);
}

function toMinutes(time) {
  const [h, m] = String(time || '').split(':').map(Number);
  return h * 60 + m;
}

function formatMinutes(total) {
  const safe = Math.max(0, Number(total || 0));
  const h = String(Math.floor(safe / 60)).padStart(2, '0');
  const m = String(safe % 60).padStart(2, '0');
  return `${h}:${m}`;
}

function effectiveEnd(end, cleanupMinutes) {
  return formatMinutes(toMinutes(end) + Number(cleanupMinutes || 0));
}

function overlaps(startA, endA, startB, endB) {
  return toMinutes(startA) < toMinutes(endB) && toMinutes(startB) < toMinutes(endA);
}

function statusOf(date) {
  const today = nowDate();
  if (date === today) return 'today';
  if (date > today) return 'upcoming';
  return 'past';
}

function statusLabel(status) {
  if (status === 'today') return 'today';
  if (status === 'upcoming') return 'upcoming';
  return 'past';
}

function showFeedback(message, type = 'success') {
  el.feedback.textContent = message;
  el.feedback.className = `feedback ${type}`;
}

function clearFeedback() {
  el.feedback.textContent = '';
  el.feedback.className = 'feedback hidden';
}

function sortBookings(a, b) {
  if (a.date !== b.date) return a.date.localeCompare(b.date);
  return a.start.localeCompare(b.start);
}

function getFormData() {
  return {
    user: fields.user.value.trim(),
    date: fields.date.value,
    purpose: fields.purpose.value.trim(),
    start: fields.start.value,
    end: fields.end.value,
    cleanupMinutes: Number(fields.cleanupMinutes.value || 0),
    contact: fields.contact.value.trim(),
    notes: fields.notes.value.trim(),
  };
}

function setFormData(data) {
  fields.user.value = data.user || '';
  fields.date.value = data.date || nowDate();
  fields.purpose.value = data.purpose || '';
  fields.start.value = data.start || '09:00';
  fields.end.value = data.end || '10:00';
  fields.cleanupMinutes.value = Number.isFinite(Number(data.cleanupMinutes)) ? data.cleanupMinutes : state.settings.defaultCleanupMinutes;
  fields.contact.value = data.contact || '';
  fields.notes.value = data.notes || '';
}

function resetForm() {
  state.editingRowNumber = null;
  el.formTitle.textContent = '登记新的预约';
  el.submitBtn.textContent = '提交预约';
  setFormData({
    date: nowDate(),
    purpose: '',
    start: '09:00',
    end: '10:00',
    cleanupMinutes: state.settings.defaultCleanupMinutes,
  });
  clearFeedback();
}

function findConflict(candidate, ignoreRowNumber = null) {
  const candidateEnd = effectiveEnd(candidate.end, candidate.cleanupMinutes);
  return state.bookings.find((item) => item._rowNumber !== ignoreRowNumber && item.date === candidate.date && overlaps(candidate.start, candidateEnd, item.start, effectiveEnd(item.end, item.cleanupMinutes)));
}

function renderStats() {
  const today = nowDate();
  el.statToday.textContent = String(state.bookings.filter((b) => b.date === today).length);
  el.statUpcoming.textContent = String(state.bookings.filter((b) => b.date >= today).length);
}

function renderHeader() {
  document.title = state.settings.labName;
  document.querySelector('h1').textContent = state.settings.labName;
  el.noticeText.textContent = state.settings.notice;
}

function renderSettingsForm() {
  settingFields.labName.value = state.settings.labName;
  settingFields.defaultCleanupMinutes.value = state.settings.defaultCleanupMinutes;
  settingFields.notice.value = state.settings.notice;
}

function bookingToCsvRow(item) {
  return [
    item.user,
    item.lab,
    item.project,
    'AKTA',
    item.date,
    item.start,
    item.end,
    item.cleanupMinutes,
    effectiveEnd(item.end, item.cleanupMinutes),
    item.purpose,
    item.contact,
    item.notes,
    item.createdAt || '',
  ];
}

function csvEscape(value) {
  const str = String(value ?? '');
  if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
  return str;
}

function exportCsv(rows) {
  const header = ['User', 'Lab', 'Project', 'Instrument', 'Date', 'Start', 'End', 'CleanupMinutes', 'EffectiveEnd', 'Purpose', 'Contact', 'Notes', 'CreatedAt'];
  const csv = [header, ...rows].map((row) => row.map(csvEscape).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `akta-bookings-${nowDate()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function filteredBookings() {
  const query = el.searchInput.value.trim().toLowerCase();
  const selectedDate = el.filterDate.value;
  return [...state.bookings]
    .filter((item) => (selectedDate ? item.date === selectedDate : true))
    .filter((item) => {
      if (!query) return true;
      const hay = [item.user, item.purpose, item.contact, item.notes].join(' ').toLowerCase();
      return hay.includes(query);
    })
    .sort(sortBookings);
}

function createEmpty(message) {
  const node = document.createElement('div');
  node.className = 'card empty';
  node.textContent = message;
  return node;
}

function createDayCard(date, items) {
  const wrapper = document.createElement('section');
  wrapper.className = 'card day-card';

  const header = document.createElement('div');
  header.className = 'day-header';
  const status = statusOf(date);
  header.innerHTML = `<div><h2>${date}</h2><p>${items.length} 条预约</p></div><span class="day-status ${status}">${statusLabel(status)}</span>`;
  wrapper.appendChild(header);

  const template = document.getElementById('bookingTemplate');
  items.forEach((item) => {
    const clone = template.content.firstElementChild.cloneNode(true);
    const badges = clone.querySelectorAll('.badge:not(.solid)');
    badges[0].textContent = item.user || '未填写姓名';
    const purposeBadge = clone.querySelector('.purpose-badge');
    if (item.purpose) {
      badges[1].textContent = '柱子类型';
      purposeBadge.textContent = item.purpose;
      purposeBadge.classList.remove('hidden');
    } else {
      badges[1].textContent = '预约';
    }
    clone.querySelector('.project-name').textContent = item.user || '预约信息';
    clone.querySelector('.meta').innerHTML = `时间：${item.start}–${item.end}<br>整理至：${effectiveEnd(item.end, item.cleanupMinutes)}`;

    const purpose = clone.querySelector('.purpose-row');
    const contact = clone.querySelector('.contact-row');
    const notes = clone.querySelector('.notes-row');

    if (item.purpose) {
      purpose.textContent = `用途（柱子类型）：${item.purpose}`;
      purpose.classList.remove('hidden');
    }
    if (item.contact) {
      contact.textContent = `联系方式：${item.contact}`;
      contact.classList.remove('hidden');
    }
    if (item.notes) {
      notes.textContent = `备注：${item.notes}`;
      notes.classList.remove('hidden');
    }

    clone.querySelector('.edit-btn').addEventListener('click', () => {
      state.editingRowNumber = item._rowNumber;
      el.formTitle.textContent = '编辑预约';
      el.submitBtn.textContent = '保存修改';
      setFormData(item);
      activateTab('book');
      showFeedback('正在编辑这条预约记录。', 'success');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    clone.querySelector('.delete-btn').addEventListener('click', async () => {
      if (!confirm('确定删除这条预约吗？删除后无法恢复。')) return;
      try {
        await apiRequest({ action: 'delete', rowNumber: item._rowNumber });
        await reloadBookings();
        showFeedback('预约已删除。', 'success');
      } catch (error) {
        showFeedback(`删除失败：${error.message}`, 'error');
      }
    });

    wrapper.appendChild(clone);
  });

  return wrapper;
}

function renderList() {
  const list = filteredBookings();
  el.listContainer.innerHTML = '';

  if (!list.length) {
    el.listContainer.appendChild(createEmpty(state.isLoading ? '正在加载预约记录…' : '还没有符合条件的预约记录。'));
    return;
  }

  const groups = list.reduce((acc, item) => {
    acc[item.date] ||= [];
    acc[item.date].push(item);
    return acc;
  }, {});

  Object.keys(groups).sort().forEach((date) => {
    el.listContainer.appendChild(createDayCard(date, groups[date]));
  });
}

function persistSettings() {
  saveJSON(SETTINGS_KEY, state.settings);
}

function renderAll() {
  renderHeader();
  renderStats();
  renderSettingsForm();
  renderList();
}

function activateTab(name) {
  el.tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.tab === name));
  Object.entries(el.panels).forEach(([key, panel]) => panel.classList.toggle('active', key === name));
}

function normalizeDateValue(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, '0');
    const day = String(parsed.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return String(value).slice(0, 10);
}

function normalizeTimeValue(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string' && /^\d{2}:\d{2}$/.test(value)) return value;
  const match = String(value).match(/(\d{2}):(\d{2})/);
  if (match) return `${match[1]}:${match[2]}`;
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    const h = String(parsed.getHours()).padStart(2, '0');
    const m = String(parsed.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  return String(value);
}

function normalizeBooking(item) {
  return {
    _rowNumber: Number(item._rowNumber),
    user: String(item.User || ''),
    lab: '',
    project: '',
    date: normalizeDateValue(item.Date),
    purpose: String(item.Purpose || ''),
    start: normalizeTimeValue(item.Start),
    end: normalizeTimeValue(item.End),
    cleanupMinutes: Number(item.CleanupMinutes || 0),
    contact: String(item.Contact || ''),
    notes: String(item.Notes || ''),
    createdAt: item.Timestamp ? String(item.Timestamp) : '',
  };
}

async function apiRequest(payload = null) {
  const options = payload
    ? {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        redirect: 'follow',
      }
    : { method: 'GET', redirect: 'follow' };

  const response = await fetch(API_URL, options);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  const json = await response.json();
  if (!json.success) {
    throw new Error(json.error || '服务器返回失败');
  }
  return json;
}

async function reloadBookings() {
  state.isLoading = true;
  renderList();
  const json = await apiRequest();
  state.bookings = (json.data || []).map(normalizeBooking).sort(sortBookings);
  state.isLoading = false;
  renderAll();
}

el.tabs.forEach((tab) => {
  tab.addEventListener('click', () => activateTab(tab.dataset.tab));
});

el.bookingForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  const data = getFormData();

  if (!data.user || !data.date || !data.purpose || !data.start || !data.end) {
    showFeedback('请填写姓名、日期、用途（柱子类型）和起止时间。', 'error');
    return;
  }

  if (toMinutes(data.end) <= toMinutes(data.start)) {
    showFeedback('结束时间必须晚于开始时间。', 'error');
    return;
  }

  try {
    el.submitBtn.disabled = true;
    el.submitBtn.textContent = state.editingRowNumber ? '保存中…' : '提交中…';

    await reloadBookings();
    const conflict = findConflict(data, state.editingRowNumber);
    if (conflict) {
      showFeedback(`时间冲突：${conflict.date} 已有 ${conflict.user} 的预约（${conflict.start}–${conflict.end}，另含 ${conflict.cleanupMinutes} 分钟整理时间）。`, 'error');
      return;
    }

    const payload = {
      action: state.editingRowNumber ? 'update' : 'create',
      rowNumber: state.editingRowNumber,
      User: data.user,
      Lab: '',
      Project: '',
      Date: data.date,
      Start: data.start,
      End: data.end,
      CleanupMinutes: data.cleanupMinutes,
      Purpose: data.purpose,
      Contact: data.contact,
      Notes: data.notes,
    };

    await apiRequest(payload);
    await reloadBookings();
    const message = state.editingRowNumber ? '预约已更新。' : '预约已登记。';
    resetForm();
    showFeedback(message, 'success');
    activateTab('list');
  } catch (error) {
    showFeedback(`提交失败：${error.message}`, 'error');
  } finally {
    el.submitBtn.disabled = false;
    el.submitBtn.textContent = state.editingRowNumber ? '保存修改' : '提交预约';
  }
});

el.resetBtn.addEventListener('click', resetForm);

el.searchInput.addEventListener('input', renderList);
el.filterDate.addEventListener('input', renderList);

el.clearFiltersBtn.addEventListener('click', () => {
  el.searchInput.value = '';
  el.filterDate.value = '';
  renderList();
});

el.exportBtn.addEventListener('click', () => {
  exportCsv(filteredBookings().map(bookingToCsvRow));
});

el.settingsForm.addEventListener('submit', (event) => {
  event.preventDefault();
  state.settings = {
    labName: settingFields.labName.value.trim() || defaultSettings.labName,
    defaultCleanupMinutes: Number(settingFields.defaultCleanupMinutes.value || 0),
    notice: settingFields.notice.value.trim() || defaultSettings.notice,
  };
  persistSettings();
  renderAll();
  if (!state.editingRowNumber) {
    fields.cleanupMinutes.value = state.settings.defaultCleanupMinutes;
  }
  activateTab('book');
  showFeedback('设置已保存（只保存在当前浏览器）。', 'success');
});

el.clearAllBtn.addEventListener('click', async () => {
  if (!confirm('确定清空所有预约记录吗？此操作不可恢复。')) return;
  try {
    const bookings = [...state.bookings];
    for (const item of bookings.sort((a, b) => b._rowNumber - a._rowNumber)) {
      await apiRequest({ action: 'delete', rowNumber: item._rowNumber });
    }
    await reloadBookings();
    activateTab('list');
    showFeedback('所有预约已清空。', 'success');
  } catch (error) {
    showFeedback(`清空失败：${error.message}`, 'error');
  }
});

async function init() {
  state.settings = { ...defaultSettings, ...loadJSON(SETTINGS_KEY, defaultSettings) };
  renderAll();
  resetForm();
  try {
    await reloadBookings();
  } catch (error) {
    state.isLoading = false;
    renderList();
    showFeedback(`无法连接 Google Sheets：${error.message}`, 'error');
  }
}

init();
