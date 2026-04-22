document.addEventListener('DOMContentLoaded', () => {
    // State
    let holidays = [];
    let deletedDefaults = [];
    let customHolidays = [];
    
    let state = {
        year: new Date().getFullYear(),
        monthStart: 1, 
        libraryClosed: 'odd', // 1, 3주차 기본값
        personalSettings: {}, 
        mode: 'regular' 
    };

    // DOM Elements
    const yearSelect = document.getElementById('year-select');
    const monthPairSelect = document.getElementById('month-pair-select');
    const libraryClosedSelect = document.getElementById('library-closed-select');
    const personalContainer = document.getElementById('personal-holidays-container');
    const runBtn = document.getElementById('run-btn');
    const resetBtn = document.getElementById('reset-btn');
    
    // Modal Elements
    const holidayModal = document.getElementById('holiday-modal');
    const openHolidayModalBtn = document.getElementById('open-holiday-modal-btn');
    const closeHolidayModalBtn = document.getElementById('close-holiday-modal-btn');
    const holidayDateInput = document.getElementById('new-holiday-date');
    const holidayNameInput = document.getElementById('new-holiday-name');
    const addHolidayBtn = document.getElementById('add-holiday-btn');
    const holidayListUl = document.getElementById('holiday-list');

    // Alert Modal Elements
    const alertModal = document.getElementById('alert-modal');
    const closeAlertBtn = document.getElementById('close-alert-btn');

    // Init
    loadHolidaysAndSettings();
    initDefaults();
    setupEventListeners();
    renderAll();

    function loadHolidaysAndSettings() {
        // 휴일 관련 로드
        const savedDeleted = localStorage.getItem('countWorkDay_deleted_defaults');
        if (savedDeleted) deletedDefaults = JSON.parse(savedDeleted);

        const savedCustom = localStorage.getItem('countWorkDay_custom_holidays');
        if (savedCustom) customHolidays = JSON.parse(savedCustom);

        buildActiveHolidays();

        // 설정 로드 (새로운 키 사용으로 기존 설정 초기화)
        const savedLib = localStorage.getItem('countWorkDay_libClosed_v2');
        if (savedLib) {
            state.libraryClosed = savedLib;
        } else {
            state.libraryClosed = ''; // 강제 선택 유도를 위해 비워둠
        }

        // 개인 휴무일은 여러 번 테스트할 수 있도록 저장하지 않고 초기화
        state.personalSettings = {};
    }

    function buildActiveHolidays() {
        holidays = [];
        if (typeof DEFAULT_HOLIDAYS !== 'undefined') {
            DEFAULT_HOLIDAYS.forEach(dh => {
                if (!deletedDefaults.includes(dh.date)) {
                    holidays.push(dh);
                }
            });
        }
        customHolidays.forEach(ch => {
            if (!holidays.find(h => h.date === ch.date)) {
                holidays.push(ch);
            }
        });
    }

    function saveHolidaysToLocalStorage() {
        localStorage.setItem('countWorkDay_custom_holidays', JSON.stringify(customHolidays));
        localStorage.setItem('countWorkDay_deleted_defaults', JSON.stringify(deletedDefaults));
        buildActiveHolidays();
        renderHolidayList();
        renderAll();
    }

    function saveSettingsToLocalStorage() {
        // 도서관 휴관일만 저장 (개인 휴무일은 매번 테스트해야 하므로 저장 제외)
        localStorage.setItem('countWorkDay_libClosed_v2', state.libraryClosed);
    }

    function initDefaults() {
        const now = new Date();
        state.year = now.getFullYear();
        
        let m = now.getMonth() + 1; 
        if (m % 2 === 0) m++; 
        if (m > 11) {
            m = 1;
            state.year++;
        }
        
        yearSelect.value = state.year;
        state.monthStart = m;
        monthPairSelect.value = m;
        libraryClosedSelect.value = state.libraryClosed;

        updatePersonalSettingsUI();
        setHolidayInputDefaultDate();
    }

    function setHolidayInputDefaultDate() {
        const y = state.year;
        const m = String(state.monthStart).padStart(2, '0');
        holidayDateInput.value = `${y}-${m}-01`;
    }

    function setupEventListeners() {
        yearSelect.addEventListener('change', (e) => {
            state.year = parseInt(e.target.value);
            state.mode = 'regular';
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            renderAll();
        });
        
        monthPairSelect.addEventListener('change', (e) => {
            state.monthStart = parseInt(e.target.value);
            state.mode = 'regular';
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            renderAll();
        });

        libraryClosedSelect.addEventListener('change', (e) => {
            state.libraryClosed = e.target.value;
            saveSettingsToLocalStorage();
            // 화면이 즉시 바뀌거나 초기화되지 않도록 renderAll() 제거
            // 오직 [공무직 달력 확인하기] 버튼을 눌렀을 때만 결과가 반영되도록 통일
        });

        resetBtn.addEventListener('click', () => {
            // 개인 휴무 설정 초기화 (로컬에는 원래 저장하지 않음)
            state.personalSettings = {};
            updatePersonalSettingsUI();
            
            // 모드를 정직원 기준으로 변경
            state.mode = 'regular';
            renderAll();
        });

        runBtn.addEventListener('click', () => {
            if (!state.libraryClosed || state.libraryClosed === '') {
                alertModal.classList.remove('hidden');
                alertModal.classList.add('flex');
                libraryClosedSelect.focus();
                libraryClosedSelect.classList.add('ring-4', 'ring-blue-200', 'border-blue-500');
                setTimeout(() => libraryClosedSelect.classList.remove('ring-4', 'ring-blue-200', 'border-blue-500'), 1500);
                return;
            }

            const rows = document.querySelectorAll('.personal-month-row');
            rows.forEach(row => {
                const key = row.dataset.key;
                const weekday = parseInt(row.querySelector('.weekday-sel').value);
                const weekend = parseInt(row.querySelector('.weekend-sel').value);
                state.personalSettings[key] = { weekday, weekend };
            });
            saveSettingsToLocalStorage();
            state.mode = 'public';
            renderAll();
        });

        // Modal Events
        openHolidayModalBtn.addEventListener('click', () => {
            holidayModal.classList.remove('hidden');
            holidayModal.classList.add('flex');
            setHolidayInputDefaultDate();
            renderHolidayList();
        });

        closeHolidayModalBtn.addEventListener('click', () => {
            holidayModal.classList.add('hidden');
            holidayModal.classList.remove('flex');
        });

        holidayModal.addEventListener('click', (e) => {
            if (e.target === holidayModal) {
                holidayModal.classList.add('hidden');
                holidayModal.classList.remove('flex');
            }
        });

        closeAlertBtn.addEventListener('click', () => {
            alertModal.classList.add('hidden');
            alertModal.classList.remove('flex');
        });

        alertModal.addEventListener('click', (e) => {
            if (e.target === alertModal) {
                alertModal.classList.add('hidden');
                alertModal.classList.remove('flex');
            }
        });

        addHolidayBtn.addEventListener('click', () => {
            if(holidayDateInput.value && holidayNameInput.value) {
                const dateStr = holidayDateInput.value;
                const nameStr = holidayNameInput.value;

                if (holidays.find(h => h.date === dateStr)) {
                    alert('이미 등록된 휴일 날짜입니다.');
                    return;
                }

                if (deletedDefaults.includes(dateStr)) {
                    deletedDefaults = deletedDefaults.filter(d => d !== dateStr);
                }

                if (!customHolidays.find(ch => ch.date === dateStr)) {
                    customHolidays.push({ date: dateStr, name: nameStr });
                }

                saveHolidaysToLocalStorage();
                holidayNameInput.value = '';
            }
        });
    }

    function renderHolidayList() {
        holidayListUl.innerHTML = '';
        holidays.sort((a, b) => a.date.localeCompare(b.date)).forEach((h) => {
            const isDefault = (typeof DEFAULT_HOLIDAYS !== 'undefined') && 
                              DEFAULT_HOLIDAYS.find(dh => dh.date === h.date && dh.name === h.name);
            
            const li = document.createElement('li');
            li.className = 'flex justify-between items-center bg-white border border-stone-100 px-3 py-2 rounded-xl shadow-sm';
            
            let tagHtml = isDefault 
                ? `<span class="text-[10px] text-stone-500 font-bold px-1.5 py-0.5 rounded bg-stone-100 border border-stone-200">기본</span>`
                : `<span class="text-[10px] text-amber-600 font-bold px-1.5 py-0.5 rounded bg-amber-50 border border-amber-200">추가됨</span>`;

            li.innerHTML = `
                <div class="flex items-center gap-2">
                    <span class="font-bold text-stone-700 w-[85px]">${h.date}</span>
                    <span class="text-stone-500 text-xs truncate max-w-[100px]">${h.name}</span>
                    ${tagHtml}
                </div>
                <button onclick="window.removeHoliday('${h.date}')" class="text-stone-400 hover:text-rose-500 hover:bg-rose-50 w-7 h-7 flex items-center justify-center rounded-lg transition-colors font-bold text-lg" title="삭제">&times;</button>
            `;
            holidayListUl.appendChild(li);
        });
    }
    
    window.removeHoliday = function(dateStr) {
        const isDefault = (typeof DEFAULT_HOLIDAYS !== 'undefined') && DEFAULT_HOLIDAYS.find(dh => dh.date === dateStr);
        
        if (isDefault) {
            if (!deletedDefaults.includes(dateStr)) {
                deletedDefaults.push(dateStr);
            }
        } else {
            customHolidays = customHolidays.filter(ch => ch.date !== dateStr);
        }
        
        saveHolidaysToLocalStorage();
    };

    function updatePersonalSettingsUI() {
        personalContainer.innerHTML = '';
        const targetMonths = [];
        
        let pM = state.monthStart - 1;
        let pY = state.year;
        if(pM === 0) { pM = 12; pY--; }
        targetMonths.push({ y: pY, m: pM, label: `${pM}월(이전)` });
        
        targetMonths.push({ y: state.year, m: state.monthStart, label: `${state.monthStart}월` });
        targetMonths.push({ y: state.year, m: state.monthStart + 1, label: `${state.monthStart + 1}월` });

        targetMonths.forEach(tm => {
            const key = `${tm.y}-${tm.m}`;
            if(!state.personalSettings[key]) {
                state.personalSettings[key] = { weekday: 2, weekend: 6 }; // Default: Tue, Sat
            }
            
            const s = state.personalSettings[key];
            const div = document.createElement('div');
            div.className = 'personal-month-row flex items-center gap-2 mb-1';
            div.dataset.key = key;
            
            div.innerHTML = `
                <span class="w-16 font-bold text-stone-700 text-right mr-1">${tm.label}</span>
                <select class="weekday-sel flex-1 p-2 rounded-xl border-2 border-stone-200 bg-white focus:ring-0 focus:border-amber-500 outline-none text-sm shadow-sm cursor-pointer">
                    <option value="1" ${s.weekday===1?'selected':''}>월요일 휴무</option>
                    <option value="2" ${s.weekday===2?'selected':''}>화요일 휴무</option>
                    <option value="3" ${s.weekday===3?'selected':''}>수요일 휴무</option>
                    <option value="4" ${s.weekday===4?'selected':''}>목요일 휴무</option>
                    <option value="5" ${s.weekday===5?'selected':''}>금요일 휴무</option>
                </select>
                <select class="weekend-sel flex-1 p-2 rounded-xl border-2 border-stone-200 bg-white focus:ring-0 focus:border-amber-500 outline-none text-sm shadow-sm cursor-pointer">
                    <option value="6" ${s.weekend===6?'selected':''}>토요일 휴무</option>
                    <option value="0" ${s.weekend===0?'selected':''}>일요일 휴무</option>
                </select>
            `;
            personalContainer.appendChild(div);
        });
    }

    function getWeekMonday(date) {
        const d = new Date(date);
        const day = d.getDay() === 0 ? 7 : d.getDay();
        d.setDate(d.getDate() - day + 1);
        return d;
    }

    function getHolidayName(dateStr) {
        const h = holidays.find(h => h.date === dateStr);
        return h ? h.name : null;
    }

    function formatDate(d) {
        const y = d.getFullYear();
        const m = String(d.getMonth()+1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    }

    function getPublicOfficialDayInfo(date) {
        const weekMonday = getWeekMonday(date);
        const schedY = weekMonday.getFullYear();
        const schedM = weekMonday.getMonth() + 1;
        const key = `${schedY}-${schedM}`;
        
        const ps = state.personalSettings[key] || { weekday: 2, weekend: 6 };
        const mondayNum = Math.floor((weekMonday.getDate() - 1) / 7) + 1;
        
        let isLibraryClosedWeek = false;
        if(state.libraryClosed === 'odd' && (mondayNum === 1 || mondayNum === 3)) {
            isLibraryClosedWeek = true;
        } else if (state.libraryClosed === 'even' && (mondayNum === 2 || mondayNum === 4)) {
            isLibraryClosedWeek = true;
        }

        const dateDay = date.getDay();
        let isOff = false;
        let reason = '';

        if (isLibraryClosedWeek && dateDay === 1) {
            isOff = true;
            reason = '도서관휴무';
        } else if (isLibraryClosedWeek && dateDay >= 1 && dateDay <= 5) {
            isOff = false;
        } else if (dateDay === ps.weekday) {
            isOff = true;
            reason = '평일휴무';
        }

        if (dateDay === ps.weekend) {
            isOff = true;
            reason = '주말휴무';
        }

        const hName = getHolidayName(formatDate(date));
        if (hName) {
            isOff = true;
            reason = hName;
        }

        return { isOff, reason, isHoliday: !!hName };
    }

    function getRegularDayInfo(date) {
        const day = date.getDay();
        const hName = getHolidayName(formatDate(date));
        
        let isOff = false;
        let reason = '';
        let isHoliday = false;

        if (day === 0 || day === 6) {
            isOff = true;
            reason = '주말';
        }
        
        if (hName) {
            isOff = true;
            reason = hName;
            isHoliday = true;
        }

        return { isOff, reason, isHoliday };
    }

    function renderAll() {
        const titleEl = document.getElementById('target-employee-type');
        if (state.mode === 'regular') {
            titleEl.innerText = '정직원 기준 (설정 전)';
            titleEl.className = 'text-stone-500 font-medium mt-2 flex items-center gap-2';
            titleEl.previousElementSibling.className = 'w-2 h-2 rounded-full bg-stone-400';
        } else {
            titleEl.innerText = '공무직 기준 (설정 적용됨)';
            titleEl.className = 'text-amber-600 font-bold mt-2 flex items-center gap-2';
            titleEl.previousElementSibling.className = 'w-2 h-2 rounded-full bg-amber-500 animate-pulse shadow-lg shadow-amber-500';
        }
        
        const cal1 = renderCalendar(1, state.year, state.monthStart);
        const cal2 = renderCalendar(2, state.year, state.monthStart + 1);

        const totalPublicWorkDays = cal1.publicDays + cal2.publicDays;
        const totalRegularWorkDays = cal1.regularDays + cal2.regularDays;

        const totalEl = document.getElementById('total-work-days');
        const regTotalEl = document.getElementById('regular-work-days');

        if(state.mode === 'regular') {
            totalEl.innerText = totalRegularWorkDays;
        } else {
            totalEl.innerText = totalPublicWorkDays;
        }
        regTotalEl.innerText = totalRegularWorkDays;

        totalEl.style.transform = 'scale(1.15)';
        setTimeout(() => totalEl.style.transform = 'scale(1)', 200);
    }

    function renderCalendar(calIndex, year, month) {
        document.getElementById(`cal-title-${calIndex}`).innerText = `${year}년 ${month}월`;
        const container = document.getElementById(`cal-days-${calIndex}`);
        container.innerHTML = '';

        const firstDay = new Date(year, month - 1, 1);
        const lastDay = new Date(year, month, 0);
        
        let emptyDays = firstDay.getDay(); 
        
        for (let i = 0; i < emptyDays; i++) {
            const div = document.createElement('div');
            div.className = 'day-cell empty min-h-[4.5rem]';
            container.appendChild(div);
        }

        let publicDays = 0;
        let regularDays = 0;

        for (let d = 1; d <= lastDay.getDate(); d++) {
            const date = new Date(year, month - 1, d);
            const div = document.createElement('div');
            // aspect-square 제거 및 min-h 지정으로 이름 짤림 방지, break-words 사용
            div.className = 'day-cell min-h-[4.5rem] flex flex-col items-center justify-start rounded-2xl font-medium text-sm relative p-1.5 transition-all hover:-translate-y-1 hover:shadow-lg hover:z-10 cursor-default shadow-sm border-2';
            
            const regInfo = getRegularDayInfo(date);
            const pubInfo = getPublicOfficialDayInfo(date);

            if (!regInfo.isOff) regularDays++;
            if (!pubInfo.isOff) publicDays++;

            const info = state.mode === 'regular' ? regInfo : pubInfo;

            if (info.isOff) {
                if (info.isHoliday || info.reason.includes('휴관')) {
                    div.classList.add('bg-hol', 'text-holText', 'border-holBorder');
                } else {
                    div.classList.add('bg-off', 'text-offText', 'border-offBorder');
                }
            } else {
                div.classList.add('bg-work', 'text-workText', 'border-workBorder');
            }

            let html = `<div class="font-bold text-base mt-0.5">${d}</div>`;
            if (info.reason) {
                // break-keep을 사용해 줄바꿈이 자연스럽게 일어나도록 유도
                html += `<div class="text-[0.65rem] leading-tight mt-1 opacity-85 text-center break-keep w-full flex-1 px-0.5">${info.reason}</div>`;
            }
            div.innerHTML = html;
            container.appendChild(div);
        }
        
        return { publicDays, regularDays };
    }
});
