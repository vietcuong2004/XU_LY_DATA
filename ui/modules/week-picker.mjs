export function createWeekPicker({$, document, normalize}) {
let allWeekItems = [];
let currentDetectedWeek = 40;

function getSuggestedWeekNumbers(centerWeek = currentDetectedWeek) {
  const list = [];
  for (let offset = -2; offset <= 2; offset++) {
    const w = centerWeek + offset;
    if (w >= 1 && w <= 53) list.push(w);
  }
  return list;
}

function initWeekItems(year = 2026) {
  allWeekItems = [
    {
      label: 'Tất cả (Không lọc tuần)',
      week: 'all',
      isAll: true,
      year: year
    }
  ];
  for (let w = 1; w <= 53; w++) {
    allWeekItems.push({
      label: String(w),
      week: w,
      year: year,
      val: `${year}-W${String(w).padStart(2, '0')}`
    });
  }
}
initWeekItems(2026);

function getWeekDateRange(year, week) {
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const dayOfWeek = jan4.getUTCDay() || 7;
  const week1Monday = new Date(jan4.getTime() - (dayOfWeek - 1) * 86400000);
  const targetMonday = new Date(week1Monday.getTime() + (week - 1) * 7 * 86400000);
  const targetSunday = new Date(targetMonday.getTime() + 6 * 86400000);

  const pad = n => String(n).padStart(2, '0');
  const fmt = d => `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  return {
    start: fmt(targetMonday),
    end: fmt(targetSunday)
  };
}

function updateWeekDateRangeDisplay() {
  const weekInput = $('shipping-week');
  const yearInput = $('shipping-year');
  const rangeText = $('week-range-text');
  const rangeBox = $('week-date-range-box');
  if (!rangeText || !rangeBox) return;

  if (!weekInput || weekInput.disabled || !weekInput.value) {
    rangeText.textContent = 'Thời gian: (chưa chọn tuần)';
    rangeBox.classList.add('is-disabled');
    return;
  }

  const currentYear = yearInput ? (parseInt(yearInput.value, 10) || 2026) : 2026;
  const parsed = parseWeekInput(weekInput.value);

  if (!parsed || parsed.isAll) {
    rangeText.textContent = 'Chế độ: Xuất toàn bộ Family tất cả các tuần';
    rangeBox.classList.remove('is-disabled');
    return;
  }

  if (!parsed.week || parsed.week < 1 || parsed.week > 53) {
    rangeText.textContent = 'Thời gian: (chưa chọn tuần)';
    rangeBox.classList.add('is-disabled');
    return;
  }

  const range = getWeekDateRange(currentYear, parsed.week);
  rangeText.textContent = `Thời gian: (từ ${range.start} - ${range.end})`;
  rangeBox.classList.remove('is-disabled');
}

function parseWeekInput(text) {
  if (!text) return null;
  text = String(text).trim();
  const lower = normalize(text).toLowerCase();
  if (lower.includes('tat ca') || lower === 'all' || lower === 'tc' || lower === 'full' || lower === 'toan bo') {
    return { isAll: true, week: null, year: 2026 };
  }
  const viMatch = text.match(/Tuần\s*(\d+)(?:[,\s]+(\d{4}))?/i);
  if (viMatch) {
    const wk = parseInt(viMatch[1], 10);
    const yr = viMatch[2] ? parseInt(viMatch[2], 10) : 2026;
    if (wk >= 1 && wk <= 53) return { year: yr, week: wk };
  }
  const isoMatch = text.match(/(\d{4})[-/ ](?:W|w)?(\d{1,2})/);
  if (isoMatch) {
    const yr = parseInt(isoMatch[1], 10);
    const wk = parseInt(isoMatch[2], 10);
    if (wk >= 1 && wk <= 53) return { year: yr, week: wk };
  }
  const slashMatch = text.match(/^(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const wk = parseInt(slashMatch[1], 10);
    const yr = parseInt(slashMatch[2], 10);
    if (wk >= 1 && wk <= 53) return { year: yr, week: wk };
  }
  const wkOnlyMatch = text.match(/^(?:WK|W)\s*(\d{1,2})$/i);
  if (wkOnlyMatch) {
    const wk = parseInt(wkOnlyMatch[1], 10);
    if (wk >= 1 && wk <= 53) return { year: 2026, week: wk };
  }
  const numMatch = text.match(/^(\d{1,2})$/);
  if (numMatch) {
    const wk = parseInt(numMatch[1], 10);
    if (wk >= 1 && wk <= 53) return { year: 2026, week: wk };
  }
  return null;
}

function formatWeekDisplay(week) {
  return String(week);
}

function renderWeekDropdown(filterText = '', isTyping = false) {
  const menu = $('week-dropdown-list');
  if (!menu) return;
  menu.innerHTML = '';

  const q = normalize(filterText).trim().toLowerCase();
  const isAllText = q.includes('tat ca') || q === 'all' || q === 'toan bo';

  let items = [];

  // Mặc định hoặc khi đang là "Tất cả": hiển thị Tất cả + 5 tuần gần đây (2 trước, hiện tại, 2 tới)
  if (!isTyping || !q || isAllText) {
    items.push({
      label: 'Tất cả (Không lọc tuần)',
      value: 'Tất cả (Không lọc tuần)',
      isAll: true
    });
    const suggested = getSuggestedWeekNumbers(currentDetectedWeek);
    suggested.forEach(w => {
      items.push({
        label: String(w),
        value: String(w),
        week: w
      });
    });

    const curVal = $('shipping-week')?.value?.trim();
    const curParsed = parseWeekInput(curVal);
    if (curParsed && curParsed.week && !suggested.includes(curParsed.week)) {
      items.push({
        label: String(curParsed.week),
        value: String(curParsed.week),
        week: curParsed.week
      });
    }
  } else {
    // Khi đang chủ động gõ tìm kiếm:
    if ('tat ca (khong loc tuan)'.includes(q) || 'all'.includes(q) || 'toan bo'.includes(q)) {
      items.push({
        label: 'Tất cả (Không lọc tuần)',
        value: 'Tất cả (Không lọc tuần)',
        isAll: true
      });
    }

    for (let w = 1; w <= 53; w++) {
      const str = String(w);
      if (str === q || str.startsWith(q)) {
        items.push({
          label: str,
          value: str,
          week: w
        });
      }
    }
  }

  if (!items.length) {
    const empty = document.createElement('div');
    empty.className = 'week-dropdown-empty';
    empty.textContent = 'Không tìm thấy tuần phù hợp';
    menu.appendChild(empty);
    return;
  }

  const currentVal = $('shipping-week')?.value?.trim();
  const curParsed = parseWeekInput(currentVal);

  items.forEach(item => {
    let isSel = false;
    if (item.isAll) {
      isSel = !curParsed || curParsed.isAll;
    } else {
      isSel = Boolean(curParsed && curParsed.week === item.week);
    }

    const el = document.createElement('div');
    el.className = 'week-dropdown-item' + (isSel ? ' is-selected' : '') + (item.isAll ? ' is-all-option' : '');

    if (item.isAll) {
      el.innerHTML = `<span><strong>Tất cả</strong> (Không lọc tuần)</span><span class="week-badge all-badge">Toàn bộ</span>`;
    } else {
      el.innerHTML = `<span class="week-num-only">${item.week}</span>`;
    }

    el.addEventListener('mousedown', e => {
      e.preventDefault();
      selectWeek(item.value);
    });
    menu.appendChild(el);
  });
}

function selectWeek(value) {
  const inp = $('shipping-week');
  if (inp) {
    inp.value = value;
  }
  closeWeekDropdown();
  updateWeekDateRangeDisplay();
}

function openWeekDropdown() {
  const inp = $('shipping-week');
  const menu = $('week-dropdown-list');
  if (!inp || inp.disabled || !menu) return;
  renderWeekDropdown('', false);
  menu.hidden = false;
  inp.classList.add('is-active');

  const selEl = menu.querySelector('.is-selected');
  if (selEl) selEl.scrollIntoView({ block: 'nearest' });
}

function closeWeekDropdown() {
  const menu = $('week-dropdown-list');
  const inp = $('shipping-week');
  if (menu) menu.hidden = true;
  if (inp) inp.classList.remove('is-active');
}

function toggleWeekDropdown() {
  const menu = $('week-dropdown-list');
  if (!menu) return;
  if (menu.hidden) {
    openWeekDropdown();
  } else {
    closeWeekDropdown();
  }
}

function populateWeekOptions(weekNum = 40, yr = 2026) {
  const inp = $('shipping-week');
  const yearSel = $('shipping-year');
  const btnToggle = $('btn-toggle-week');
  if (!inp) return;

  currentDetectedWeek = weekNum;
  initWeekItems(yr);

  inp.value = 'Tất cả (Không lọc tuần)';
  inp.disabled = false;
  if (yearSel) {
    yearSel.value = String(yr);
    yearSel.disabled = false;
  }
  if (btnToggle) btnToggle.disabled = false;
  updateWeekDateRangeDisplay();
}

function resetWeekOptions() {
  const inp = $('shipping-week');
  const yearSel = $('shipping-year');
  const btnToggle = $('btn-toggle-week');
  if (!inp) return;
  inp.value = '';
  inp.placeholder = 'Tất cả (Không lọc tuần)';
  inp.disabled = true;
  if (yearSel) yearSel.disabled = true;
  if (btnToggle) btnToggle.disabled = true;
  closeWeekDropdown();
  updateWeekDateRangeDisplay();
}

function bind() {
const shippingWeekEl = $('shipping-week');
const btnToggleWeekEl = $('btn-toggle-week');

if (shippingWeekEl) {
  shippingWeekEl.addEventListener('input', () => {
    if ($('week-dropdown-list')?.hidden) {
      $('week-dropdown-list').hidden = false;
      shippingWeekEl.classList.add('is-active');
    }
    renderWeekDropdown(shippingWeekEl.value, true);
    updateWeekDateRangeDisplay();
  });

  shippingWeekEl.addEventListener('focus', () => {
    if (!shippingWeekEl.disabled) openWeekDropdown();
  });

  shippingWeekEl.addEventListener('click', () => {
    if (!shippingWeekEl.disabled && $('week-dropdown-list')?.hidden) {
      openWeekDropdown();
    }
  });

  shippingWeekEl.addEventListener('blur', () => {
    setTimeout(closeWeekDropdown, 180);
    const text = shippingWeekEl.value.trim();
    if (!text) return;
    const parsed = parseWeekInput(text);
    if (parsed) {
      if (parsed.isAll) {
        shippingWeekEl.value = 'Tất cả (Không lọc tuần)';
      } else {
        shippingWeekEl.value = formatWeekDisplay(parsed.week);
      }
    }
    updateWeekDateRangeDisplay();
  });

  shippingWeekEl.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      closeWeekDropdown();
    } else if (e.key === 'Enter') {
      closeWeekDropdown();
      const text = shippingWeekEl.value.trim();
      const parsed = parseWeekInput(text);
      if (parsed) {
        if (parsed.isAll) {
          shippingWeekEl.value = 'Tất cả (Không lọc tuần)';
        } else {
          shippingWeekEl.value = formatWeekDisplay(parsed.week);
        }
      }
      updateWeekDateRangeDisplay();
    }
  });
}

const shippingYearEl = $('shipping-year');
if (shippingYearEl) {
  shippingYearEl.addEventListener('change', () => {
    const yr = parseInt(shippingYearEl.value, 10) || 2026;
    initWeekItems(yr);
    updateWeekDateRangeDisplay();
  });
}

if (btnToggleWeekEl) {
  btnToggleWeekEl.addEventListener('click', e => {
    e.stopPropagation();
    if (shippingWeekEl?.disabled) return;
    toggleWeekDropdown();
    if (!$('week-dropdown-list')?.hidden) shippingWeekEl?.focus();
  });
}

document.addEventListener('click', e => {
  if (!e.target.closest('#week-combobox-wrap')) {
    closeWeekDropdown();
  }
});

}
return {parseWeekInput, populateWeekOptions, resetWeekOptions, bind};
}
