document.addEventListener('DOMContentLoaded', () => {
    const DEFAULT_LIBRARY_NAME = '노을빛도서관';
    const SCHEDULE_START_COLUMN = 2;
    const SCHEDULE_END_COLUMN = 8;
    const ROOM_SETTING_OPTIONS = ['종합자료실', '어린이실', '유아실', '장난감실'];
    const SHIFT_SETTING_OPTIONS = ['주간', '야간'];
    const CUSTOM_ROOM_SETTING_VALUE = '__custom__';
    const NIGHT_WORK_TIME_TEXT = '18:00~22:00';
    const DAY_WORK_TIME_TEXT = '09:00~18:00';
    const NIGHT_WEEKEND_WORK_TIME_TEXT = ' (주말: 09:00~18:00)';
    const MIXED_WORK_TIME_TEXT = '(주간) 09:00~18:00, (야간)18:00~22:00';
    const MIXED_WEEKEND_WORK_TIME_TEXT = ' (주말)09:00~18:00';

    let holidays = [];
    let deletedDefaults = [];
    let customHolidays = [];

    let excelWorkers = [];
    let excelDataByDate = {};
    let excelWorkerDetails = {};
    let isExcelExporting = false;
    let excelExportOptions = {
        libraryName: '',
        showOffWorkers: false,
        showDetailedSettings: false,
        useRoomSettings: true,
        useShiftSettings: true
    };

    const state = {
        year: new Date().getFullYear(),
        monthStart: 1,
        libraryClosed: 'odd',
        personalSettings: {},
        mode: 'regular'
    };

    const yearSelect = document.getElementById('year-select');
    const monthPairSelect = document.getElementById('month-pair-select');
    const libraryClosedSelect = document.getElementById('library-closed-select');
    const personalContainer = document.getElementById('personal-holidays-container');
    const runBtn = document.getElementById('run-btn');
    const runBtnLabel = document.getElementById('run-btn-label');
    const runBtnHint = document.getElementById('run-btn-hint');
    const resetBtn = document.getElementById('reset-btn');

    const workerNameInput = document.getElementById('worker-name-input');
    const addExcelBtn = document.getElementById('add-excel-btn');
    const downloadExcelBtn = document.getElementById('download-excel-btn');
    const excelCountEl = document.getElementById('excel-count');
    const excelWorkersList = document.getElementById('excel-workers-list');
    const excelToggleBtn = document.getElementById('excel-toggle-btn');
    const excelToggleArrow = document.getElementById('excel-toggle-arrow');
    const excelPanel = document.getElementById('excel-panel');

    const excelExportModal = document.getElementById('excel-export-modal');
    const closeExcelExportModalBtn = document.getElementById('close-excel-export-modal-btn');
    const cancelExcelExportBtn = document.getElementById('cancel-excel-export-btn');
    const confirmExcelExportBtn = document.getElementById('confirm-excel-export-btn');
    const excelLibraryNameInput = document.getElementById('excel-library-name-input');
    const excelShowOffWorkersInput = document.getElementById('excel-show-off-workers');
    const excelShowDetailSettingsInput = document.getElementById('excel-show-detail-settings');
    const excelDetailSettingsSection = document.getElementById('excel-detail-settings-section');
    const excelEnableRoomSettingsInput = document.getElementById('excel-enable-room-settings');
    const excelEnableShiftSettingsInput = document.getElementById('excel-enable-shift-settings');
    const excelDetailSettingsNote = document.getElementById('excel-detail-settings-note');
    const excelDetailSettingsEmpty = document.getElementById('excel-detail-settings-empty');
    const excelDetailSettingsTable = document.getElementById('excel-detail-settings-table');
    const excelDetailSettingsHead = document.getElementById('excel-detail-settings-head');
    const excelDetailSettingsBody = document.getElementById('excel-detail-settings-body');

    const holidayModal = document.getElementById('holiday-modal');
    const openHolidayModalBtn = document.getElementById('open-holiday-modal-btn');
    const closeHolidayModalBtn = document.getElementById('close-holiday-modal-btn');
    const holidayDateInput = document.getElementById('new-holiday-date');
    const holidayNameInput = document.getElementById('new-holiday-name');
    const addHolidayBtn = document.getElementById('add-holiday-btn');
    const holidayListUl = document.getElementById('holiday-list');

    const alertModal = document.getElementById('alert-modal');
    const alertModalIcon = document.getElementById('alert-modal-icon');
    const alertModalTitle = document.getElementById('alert-modal-title');
    const alertModalMessage = document.getElementById('alert-modal-message');
    const closeAlertBtn = document.getElementById('close-alert-btn');

    const confirmActionModal = document.getElementById('confirm-action-modal');
    const confirmActionModalIcon = document.getElementById('confirm-action-modal-icon');
    const confirmActionModalLabel = document.getElementById('confirm-action-modal-label');
    const confirmActionModalTitle = document.getElementById('confirm-action-modal-title');
    const confirmActionModalMessage = document.getElementById('confirm-action-modal-message');
    const closeConfirmActionModalBtn = document.getElementById('close-confirm-action-modal-btn');
    const cancelConfirmActionModalBtn = document.getElementById('cancel-confirm-action-modal-btn');
    const confirmActionModalBtn = document.getElementById('confirm-action-modal-btn');

    const roomInputModal = document.getElementById('room-input-modal');
    const closeRoomInputModalBtn = document.getElementById('close-room-input-modal-btn');
    const cancelRoomInputModalBtn = document.getElementById('cancel-room-input-modal-btn');
    const confirmRoomInputModalBtn = document.getElementById('confirm-room-input-modal-btn');
    const roomInputField = document.getElementById('room-input-field');

    let alertModalOnClose = null;
    let confirmActionModalResolver = null;
    let roomInputModalResolver = null;

    loadHolidaysAndSettings();
    initDefaults();
    setupEventListeners();
    renderAll();

    function loadHolidaysAndSettings() {
        const savedDeleted = localStorage.getItem('countWorkDay_deleted_defaults');
        if (savedDeleted) deletedDefaults = JSON.parse(savedDeleted);

        const savedCustom = localStorage.getItem('countWorkDay_custom_holidays');
        if (savedCustom) customHolidays = JSON.parse(savedCustom);

        buildActiveHolidays();

        const savedLib = localStorage.getItem('countWorkDay_libClosed_v2');
        state.libraryClosed = savedLib || '';
        state.personalSettings = {};
    }

    function buildActiveHolidays() {
        holidays = [];

        if (typeof DEFAULT_HOLIDAYS !== 'undefined') {
            DEFAULT_HOLIDAYS.forEach((holiday) => {
                if (!deletedDefaults.includes(holiday.date)) {
                    holidays.push(holiday);
                }
            });
        }

        customHolidays.forEach((holiday) => {
            if (!holidays.find((item) => item.date === holiday.date)) {
                holidays.push(holiday);
            }
        });
    }

    function saveHolidaysToLocalStorage() {
        localStorage.setItem('countWorkDay_custom_holidays', JSON.stringify(customHolidays));
        localStorage.setItem('countWorkDay_deleted_defaults', JSON.stringify(deletedDefaults));
        buildActiveHolidays();
        renderHolidayList();
        resetPublicCalendarPreview();
    }

    function saveSettingsToLocalStorage() {
        localStorage.setItem('countWorkDay_libClosed_v2', state.libraryClosed);
    }

    function initDefaults() {
        const now = new Date();
        state.year = now.getFullYear();

        let month = now.getMonth() + 1;
        if (month % 2 === 0) month++;
        if (month > 11) {
            month = 1;
            state.year++;
        }

        yearSelect.value = state.year;
        state.monthStart = month;
        monthPairSelect.value = month;
        libraryClosedSelect.value = state.libraryClosed;

        updatePersonalSettingsUI();
        setHolidayInputDefaultDate();
    }

    function setHolidayInputDefaultDate() {
        holidayDateInput.value = `${state.year}-${String(state.monthStart).padStart(2, '0')}-01`;
    }

    function shiftYearMonth(year, month, delta) {
        let nextYear = year;
        let nextMonth = month + delta;

        while (nextMonth < 1) {
            nextMonth += 12;
            nextYear -= 1;
        }

        while (nextMonth > 12) {
            nextMonth -= 12;
            nextYear += 1;
        }

        return { year: nextYear, month: nextMonth };
    }

    function getExcelDetailSettingKey(year, month) {
        return `${year}-${month}`;
    }

    function getExcelDetailMonthConfigs() {
        const previous = shiftYearMonth(state.year, state.monthStart, -1);
        const current = { year: state.year, month: state.monthStart };
        const next = shiftYearMonth(state.year, state.monthStart, 1);

        return [
            {
                key: getExcelDetailSettingKey(previous.year, previous.month),
                label: `${previous.month}월(이전)`,
                year: previous.year,
                month: previous.month,
                isReference: true
            },
            {
                key: getExcelDetailSettingKey(current.year, current.month),
                label: `${current.month}월`,
                year: current.year,
                month: current.month,
                isReference: false
            },
            {
                key: getExcelDetailSettingKey(next.year, next.month),
                label: `${next.month}월`,
                year: next.year,
                month: next.month,
                isReference: false
            }
        ];
    }

    function createDefaultExcelWorkerDetail() {
        return {
            weekdayRoom: '',
            weekendRoom: '',
            shift: ''
        };
    }

    function isRoomDetailType(detailType) {
        return detailType === 'room' || detailType === 'weekdayRoom' || detailType === 'weekendRoom';
    }

    function getExcelWorkerDetailFieldValue(detail, fieldKey) {
        if (!detail) {
            return '';
        }

        if (fieldKey === 'weekdayRoom') {
            return detail.weekdayRoom || detail.room || '';
        }

        if (fieldKey === 'weekendRoom') {
            return detail.weekendRoom || detail.weekdayRoom || detail.room || '';
        }

        return detail[fieldKey] || '';
    }

    function ensureExcelWorkerDetailSettings(name) {
        if (!excelWorkerDetails[name]) {
            excelWorkerDetails[name] = {};
        }

        getExcelDetailMonthConfigs().forEach(({ key }) => {
            if (!excelWorkerDetails[name][key]) {
                excelWorkerDetails[name][key] = createDefaultExcelWorkerDetail();
            }
        });
    }

    function ensureAllExcelWorkerDetailSettings() {
        excelWorkers.forEach((name) => {
            ensureExcelWorkerDetailSettings(name);
        });
    }

    function getExcelWorkerDetail(name, year, month) {
        const detailKey = getExcelDetailSettingKey(year, month);
        return excelWorkerDetails[name]?.[detailKey] || null;
    }

    function resetExcelWorkerData() {
        excelWorkers = [];
        excelDataByDate = {};
        excelWorkerDetails = {};
        updateExcelWorkersUI();
    }

    function confirmExcelScheduleReset(changeLabel) {
        if (excelWorkers.length === 0) {
            return Promise.resolve(true);
        }

        return openConfirmActionModal({
            label: 'Schedule Reset',
            title: `${changeLabel}을 변경할까요?`,
            message: `${changeLabel} 변경은 엑셀 근무편성표 계산 기준에 영향을 줍니다.\n이미 추가한 근무자 명단과 상세설정이 모두 삭제됩니다.`,
            icon: '⚠️',
            confirmText: '삭제 후 변경',
            cancelText: '취소'
        });
    }

    function getActiveExcelDetailSettingColumns() {
        const columns = [];

        if (excelEnableRoomSettingsInput.checked) {
            columns.push({ key: 'weekdayRoom', label: '평일자료실' });
        }

        if (excelEnableShiftSettingsInput.checked) {
            columns.push({ key: 'shift', label: '주야간' });
        }

        if (excelEnableRoomSettingsInput.checked) {
            columns.push({ key: 'weekendRoom', label: '주말자료실' });
        }

        return columns;
    }

    function closeAlertModal() {
        alertModal.classList.add('hidden');
        alertModal.classList.remove('flex');

        const onClose = alertModalOnClose;
        alertModalOnClose = null;
        if (typeof onClose === 'function') {
            onClose();
        }
    }

    function openAlertModal(options = {}) {
        const {
            title = '안내',
            message = '',
            icon = '🔔',
            tone = 'info',
            onClose = null
        } = options;

        const toneClassMap = {
            info: {
                iconWrap: ['bg-blue-50', 'text-blue-500', 'border-blue-100'],
                button: ['from-blue-500', 'to-indigo-500', 'hover:from-blue-600', 'hover:to-indigo-600', 'shadow-blue-500/30']
            },
            warning: {
                iconWrap: ['bg-amber-50', 'text-amber-500', 'border-amber-100'],
                button: ['from-amber-500', 'to-orange-500', 'hover:from-amber-600', 'hover:to-orange-600', 'shadow-orange-500/30']
            },
            danger: {
                iconWrap: ['bg-rose-50', 'text-rose-500', 'border-rose-100'],
                button: ['from-rose-500', 'to-red-500', 'hover:from-rose-600', 'hover:to-red-600', 'shadow-rose-500/30']
            }
        };

        const currentTone = toneClassMap[tone] || toneClassMap.info;
        const allIconClasses = ['bg-blue-50', 'text-blue-500', 'border-blue-100', 'bg-amber-50', 'text-amber-500', 'border-amber-100', 'bg-rose-50', 'text-rose-500', 'border-rose-100'];
        const allButtonClasses = ['from-blue-500', 'to-indigo-500', 'hover:from-blue-600', 'hover:to-indigo-600', 'shadow-blue-500/30', 'from-amber-500', 'to-orange-500', 'hover:from-amber-600', 'hover:to-orange-600', 'shadow-orange-500/30', 'from-rose-500', 'to-red-500', 'hover:from-rose-600', 'hover:to-red-600', 'shadow-rose-500/30'];

        alertModalIcon.classList.remove(...allIconClasses);
        alertModalIcon.classList.add(...currentTone.iconWrap);
        closeAlertBtn.classList.remove(...allButtonClasses);
        closeAlertBtn.classList.add(...currentTone.button);

        alertModalIcon.textContent = icon;
        alertModalTitle.textContent = title;
        alertModalMessage.textContent = message;
        alertModalOnClose = onClose;

        alertModal.classList.remove('hidden');
        alertModal.classList.add('flex');
        requestAnimationFrame(() => {
            closeAlertBtn.focus();
        });
    }

    function closeConfirmActionModal(result = false) {
        confirmActionModal.classList.add('hidden');
        confirmActionModal.classList.remove('flex');

        const resolver = confirmActionModalResolver;
        confirmActionModalResolver = null;
        if (typeof resolver === 'function') {
            resolver(result);
        }
    }

    function openConfirmActionModal(options = {}) {
        const {
            label = 'Schedule Reset',
            title = '변경 내용을 적용할까요?',
            message = '',
            icon = '⚠️',
            confirmText = '확인',
            cancelText = '취소'
        } = options;

        if (confirmActionModalResolver) {
            closeConfirmActionModal(false);
        }

        confirmActionModalLabel.textContent = label;
        confirmActionModalTitle.textContent = title;
        confirmActionModalMessage.textContent = message;
        confirmActionModalIcon.textContent = icon;
        confirmActionModalBtn.textContent = confirmText;
        cancelConfirmActionModalBtn.textContent = cancelText;

        confirmActionModal.classList.remove('hidden');
        confirmActionModal.classList.add('flex');

        requestAnimationFrame(() => {
            confirmActionModalBtn.focus();
        });

        return new Promise((resolve) => {
            confirmActionModalResolver = resolve;
        });
    }

    function closeRoomInputModal(result = null) {
        roomInputModal.classList.add('hidden');
        roomInputModal.classList.remove('flex');

        const resolver = roomInputModalResolver;
        roomInputModalResolver = null;
        if (typeof resolver === 'function') {
            resolver(result);
        }
    }

    function openRoomInputModal(initialValue = '') {
        if (roomInputModalResolver) {
            closeRoomInputModal(null);
        }

        roomInputField.value = initialValue;
        roomInputModal.classList.remove('hidden');
        roomInputModal.classList.add('flex');

        requestAnimationFrame(() => {
            roomInputField.focus();
            roomInputField.select();
        });

        return new Promise((resolve) => {
            roomInputModalResolver = resolve;
        });
    }

    function syncPersonalSettingsFromUI() {
        const rows = document.querySelectorAll('.personal-month-row');
        rows.forEach((row) => {
            const key = row.dataset.key;
            const weekday = parseInt(row.querySelector('.weekday-sel').value, 10);
            const weekend = parseInt(row.querySelector('.weekend-sel').value, 10);
            state.personalSettings[key] = { weekday, weekend };
        });
    }

    function resetPublicCalendarPreview() {
        state.mode = 'regular';
        renderAll();
    }

    function updateRunButtonState() {
        const needsApply = state.mode === 'regular';
        const hasLibraryClosed = !!state.libraryClosed;
        const shouldEmphasize = needsApply && hasLibraryClosed;

        runBtn.style.animation = shouldEmphasize ? 'ctaFloat 2.2s ease-in-out infinite' : 'none';
        runBtn.style.boxShadow = shouldEmphasize
            ? '0 18px 40px rgba(249, 115, 22, 0.38)'
            : '0 10px 24px rgba(249, 115, 22, 0.18)';
        runBtn.style.outline = shouldEmphasize ? '3px solid rgba(253, 230, 138, 0.7)' : 'none';
        runBtn.style.outlineOffset = shouldEmphasize ? '2px' : '0';

        if (!hasLibraryClosed) {
            runBtnLabel.textContent = '공무직 휴관일 달력 확인하기';
            runBtnHint.textContent = '먼저 휴관일을 선택하세요';
            return;
        }

        if (needsApply) {
            runBtnLabel.textContent = '공무직 휴관일 달력 확인하기';
            runBtnHint.textContent = '설정을 바꿨다면 눌러서 반영하세요';
            return;
        }

        runBtnLabel.textContent = '공무직 휴관일 달력 확인하기';
        runBtnHint.textContent = '현재 설정이 달력에 반영되어 있습니다';
    }

    function setupEventListeners() {
        yearSelect.addEventListener('change', async (e) => {
            const newYear = parseInt(e.target.value, 10);
            if (newYear === state.year) {
                return;
            }

            if (!(await confirmExcelScheduleReset('대상 연도'))) {
                yearSelect.value = state.year;
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            state.year = newYear;
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            resetPublicCalendarPreview();
        });

        monthPairSelect.addEventListener('change', async (e) => {
            const newMonth = parseInt(e.target.value, 10);
            if (newMonth === state.monthStart) {
                return;
            }

            if (!(await confirmExcelScheduleReset('대상 월'))) {
                monthPairSelect.value = state.monthStart;
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            state.monthStart = newMonth;
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            resetPublicCalendarPreview();
        });

        libraryClosedSelect.addEventListener('change', async (e) => {
            const nextLibraryClosed = e.target.value;
            if (nextLibraryClosed === state.libraryClosed) {
                return;
            }

            if (!(await confirmExcelScheduleReset('도서관 휴관일'))) {
                libraryClosedSelect.value = state.libraryClosed;
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            state.libraryClosed = nextLibraryClosed;
            saveSettingsToLocalStorage();
            resetPublicCalendarPreview();
        });

        personalContainer.addEventListener('change', async (e) => {
            if (!e.target.matches('.weekday-sel, .weekend-sel')) {
                return;
            }

            if (!(await confirmExcelScheduleReset('공무직 개인 휴일'))) {
                updatePersonalSettingsUI();
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            syncPersonalSettingsFromUI();
            resetPublicCalendarPreview();
        });

        resetBtn.addEventListener('click', async () => {
            if (!(await confirmExcelScheduleReset('공무직 개인 휴일 초기화'))) {
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            state.personalSettings = {};
            updatePersonalSettingsUI();
            resetPublicCalendarPreview();
        });

        excelToggleBtn.addEventListener('click', () => {
            const isHidden = excelPanel.classList.contains('hidden');
            if (isHidden) {
                excelPanel.classList.remove('hidden');
                excelPanel.classList.add('flex');
                excelToggleArrow.textContent = '▲';
            } else {
                excelPanel.classList.add('hidden');
                excelPanel.classList.remove('flex');
                excelToggleArrow.textContent = '▼';
            }
        });

        workerNameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                addExcelBtn.click();
            }
        });

        addExcelBtn.addEventListener('click', () => {
            if (state.mode === 'regular') {
                openAlertModal({
                    title: '근무자 추가 안내',
                    message: '먼저 공무직 휴관일 달력을 설정한 뒤 확인하기 버튼을 눌러주세요.\n기준 근무일 상태에서는 근무자를 추가할 수 없습니다.',
                    icon: '🔔',
                    tone: 'info'
                });
                return;
            }

            const name = workerNameInput.value.trim();
            if (!name) {
                openAlertModal({
                    title: '입력 확인',
                    message: '근무자 이름을 입력해주세요.',
                    icon: '✏️',
                    tone: 'warning',
                    onClose: () => {
                        workerNameInput.focus();
                    }
                });
                return;
            }
            if (excelWorkers.includes(name)) {
                openAlertModal({
                    title: '중복된 근무자',
                    message: '이미 추가된 근무자입니다.',
                    icon: '⚠️',
                    tone: 'warning'
                });
                return;
            }
            if (!state.libraryClosed) {
                openAlertModal({
                    title: '휴관일 선택 안내',
                    message: '공무직 도서관 휴관일을 먼저 선택해 주세요.\n해당 사항이 없다면 "없음"을 선택하세요.',
                    icon: '🔔',
                    tone: 'info',
                    onClose: () => {
                        libraryClosedSelect.focus();
                    }
                });
                return;
            }

            syncPersonalSettingsFromUI();

            excelWorkers.push(name);
            ensureExcelWorkerDetailSettings(name);
            updateExcelWorkersUI();
            workerNameInput.value = '';

            for (let calIndex = 1; calIndex <= 2; calIndex++) {
                const month = calIndex === 1 ? state.monthStart : state.monthStart + 1;
                const year = state.year;
                const lastDay = new Date(year, month, 0).getDate();

                for (let day = 1; day <= lastDay; day++) {
                    const date = new Date(year, month - 1, day);
                    const dateStr = formatDate(date);
                    const info = getPublicOfficialDayInfo(date);

                    if (!excelDataByDate[dateStr]) {
                        excelDataByDate[dateStr] = { workers: [], isClosed: false, holidayName: null };
                    }

                    if (info.isHoliday) {
                        excelDataByDate[dateStr].holidayName = info.reason;
                    } else if (info.reason === '도서관휴무') {
                        excelDataByDate[dateStr].isClosed = true;
                    } else if (!info.isOff) {
                        excelDataByDate[dateStr].workers.push(name);
                    }
                }
            }
        });

        downloadExcelBtn.addEventListener('click', () => {
            if (excelWorkers.length === 0) {
                openAlertModal({
                    title: '엑셀 명단 확인',
                    message: '엑셀에 추가된 인원이 없습니다. 먼저 명단에 추가해주세요.',
                    icon: '📋',
                    tone: 'warning'
                });
                return;
            }

            openExcelExportModal();
        });

        confirmExcelExportBtn.addEventListener('click', async () => {
            if (isExcelExporting) return;

            excelExportOptions = getExcelExportOptionsFromForm();
            setExcelExportLoadingState(true);

            try {
                await generateExcelWorkbook(excelExportOptions);
                setExcelExportLoadingState(false);
                closeExcelExportModal(true);
            } catch (error) {
                console.error(error);
                setExcelExportLoadingState(false);
                openAlertModal({
                    title: '엑셀 생성 오류',
                    message: '엑셀 파일을 생성하는 중 오류가 발생했습니다.',
                    icon: '⚠️',
                    tone: 'danger'
                });
            }
        });

        closeExcelExportModalBtn.addEventListener('click', () => {
            closeExcelExportModal();
        });

        cancelExcelExportBtn.addEventListener('click', () => {
            closeExcelExportModal();
        });

        excelExportModal.addEventListener('click', (e) => {
            if (e.target === excelExportModal) {
                closeExcelExportModal();
            }
        });

        excelLibraryNameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                confirmExcelExportBtn.click();
            }
        });

        excelShowDetailSettingsInput.addEventListener('change', () => {
            if (excelShowDetailSettingsInput.checked) {
                ensureAllExcelWorkerDetailSettings();
            }
            renderExcelDetailSettingsPanel();
        });

        excelEnableRoomSettingsInput.addEventListener('change', () => {
            renderExcelDetailSettingsPanel();
        });

        excelEnableShiftSettingsInput.addEventListener('change', () => {
            renderExcelDetailSettingsPanel();
        });

        excelDetailSettingsBody.addEventListener('change', async (e) => {
            const target = e.target;
            if (!(target instanceof HTMLSelectElement)) {
                return;
            }

            const workerIndex = parseInt(target.dataset.workerIndex, 10);
            const monthKey = target.dataset.monthKey;
            const detailType = target.dataset.detailType;
            const normalizedDetailType = detailType === 'room' ? 'weekdayRoom' : detailType;
            const workerName = excelWorkers[workerIndex];

            if (!workerName || !monthKey || !normalizedDetailType) {
                return;
            }

            ensureExcelWorkerDetailSettings(workerName);
            const detail = excelWorkerDetails[workerName][monthKey];

            if (isRoomDetailType(normalizedDetailType)) {
                if (target.value === CUSTOM_ROOM_SETTING_VALUE) {
                    const currentValue = getExcelWorkerDetailFieldValue(detail, normalizedDetailType);
                    const initialValue = currentValue && !ROOM_SETTING_OPTIONS.includes(currentValue) ? currentValue : currentValue;
                    const customValue = await openRoomInputModal(initialValue);
                    if (customValue === null) {
                        renderExcelDetailSettingsPanel();
                        return;
                    }

                    detail[normalizedDetailType] = customValue.trim();
                } else {
                    detail[normalizedDetailType] = target.value;
                }

                renderExcelDetailSettingsPanel();
                return;
            }

            if (normalizedDetailType === 'shift') {
                detail.shift = target.value;
            }
        });

        runBtn.addEventListener('click', () => {
            if (!state.libraryClosed) {
                openAlertModal({
                    title: '휴관일 선택 안내',
                    message: '공무직 도서관 휴관일을 먼저 선택해 주세요.\n해당 사항이 없다면 "없음"을 선택하세요.',
                    icon: '🔔',
                    tone: 'info',
                    onClose: () => {
                        libraryClosedSelect.focus();
                        libraryClosedSelect.classList.add('ring-4', 'ring-blue-200', 'border-blue-500');
                        setTimeout(() => {
                            libraryClosedSelect.classList.remove('ring-4', 'ring-blue-200', 'border-blue-500');
                        }, 1500);
                    }
                });
                return;
            }

            syncPersonalSettingsFromUI();

            saveSettingsToLocalStorage();
            state.mode = 'public';
            renderAll();
        });

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
            closeAlertModal();
        });

        closeConfirmActionModalBtn.addEventListener('click', () => {
            closeConfirmActionModal(false);
        });

        cancelConfirmActionModalBtn.addEventListener('click', () => {
            closeConfirmActionModal(false);
        });

        confirmActionModalBtn.addEventListener('click', () => {
            closeConfirmActionModal(true);
        });

        closeRoomInputModalBtn.addEventListener('click', () => {
            closeRoomInputModal(null);
        });

        cancelRoomInputModalBtn.addEventListener('click', () => {
            closeRoomInputModal(null);
        });

        confirmRoomInputModalBtn.addEventListener('click', () => {
            closeRoomInputModal(roomInputField.value);
        });

        alertModal.addEventListener('click', (e) => {
            if (e.target === alertModal) {
                closeAlertModal();
            }
        });

        confirmActionModal.addEventListener('click', (e) => {
            if (e.target === confirmActionModal) {
                closeConfirmActionModal(false);
            }
        });

        roomInputModal.addEventListener('click', (e) => {
            if (e.target === roomInputModal) {
                closeRoomInputModal(null);
            }
        });

        document.addEventListener('keydown', (e) => {
            if (!confirmActionModalResolver) {
                return;
            }

            if (e.key === 'Escape') {
                e.preventDefault();
                closeConfirmActionModal(false);
            }

            if (e.key === 'Enter') {
                e.preventDefault();
                closeConfirmActionModal(true);
            }
        });

        roomInputField.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                closeRoomInputModal(roomInputField.value);
            }

            if (e.key === 'Escape') {
                e.preventDefault();
                closeRoomInputModal(null);
            }
        });

        addHolidayBtn.addEventListener('click', async () => {
            if (!holidayDateInput.value || !holidayNameInput.value) return;

            const dateStr = holidayDateInput.value;
            const nameStr = holidayNameInput.value;

            if (holidays.find((holiday) => holiday.date === dateStr)) {
                openAlertModal({
                    title: '휴일 중복 안내',
                    message: '이미 등록된 휴일 날짜입니다.',
                    icon: '📅',
                    tone: 'warning'
                });
                return;
            }

            if (!(await confirmExcelScheduleReset('휴일 목록'))) {
                return;
            }

            if (excelWorkers.length > 0) {
                resetExcelWorkerData();
            }

            if (deletedDefaults.includes(dateStr)) {
                deletedDefaults = deletedDefaults.filter((date) => date !== dateStr);
            }

            if (!customHolidays.find((holiday) => holiday.date === dateStr)) {
                customHolidays.push({ date: dateStr, name: nameStr });
            }

            saveHolidaysToLocalStorage();
            holidayNameInput.value = '';
        });
    }

    async function generateExcelWorkbook(exportOptions) {
        const libraryName = getResolvedLibraryName(exportOptions.libraryName);

        function base64ToArrayBuffer(base64) {
            const binaryString = window.atob(base64);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        }

        if (typeof templateBase64 === 'undefined') {
            throw new Error('템플릿 데이터를 찾을 수 없습니다. (template.js 파일 로드 실패)');
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(base64ToArrayBuffer(templateBase64));

        const regularWorkDays = calculateRegularWorkDaysForRange();
        const totalWorkerDays = createWorkerCountMap();

        for (let index = 0; index < 2; index++) {
            const sheet = workbook.worksheets[index];
            if (!sheet) continue;

            const year = state.year;
            const month = index === 0 ? state.monthStart : state.monthStart + 1;
            const showOffWorkers = exportOptions.showOffWorkers;

            sheet.name = `${year.toString().slice(-2)}년 ${month}월`;
            sheet.getCell('A1').value = getExcelSheetTitle(year, month, libraryName, exportOptions);
            replaceLibraryNameInSheet(sheet, libraryName);
            replaceWorkTimeInSheet(sheet, year, month, exportOptions);

            const monthlyWorkerDays = createWorkerCountMap();
            clearScheduleArea(sheet, 17);

            const { lastUsedRow, scheduleEndRow } = writeMonthlyScheduleArea(
                sheet,
                year,
                month,
                exportOptions,
                showOffWorkers,
                monthlyWorkerDays,
                totalWorkerDays
            );

            if (lastUsedRow < scheduleEndRow) {
                for (let row = lastUsedRow + 1; row <= scheduleEndRow; row++) {
                    for (let col = 1; col <= 9; col++) {
                        sheet.getCell(row, col).value = '';
                    }
                }

                const rowsToDelete = scheduleEndRow - lastUsedRow;
                for (let count = 0; count < rowsToDelete; count++) {
                    sheet.spliceRows(lastUsedRow + 1, 1);
                }
            }

            const monthRegularDays = calculateRegularWorkDaysForMonth(year, month);
            const summaryStartRow = sheet.rowCount + 4;
            let summaryRow = summaryStartRow;

            sheet.getCell(summaryRow, 1).value = `${month}월 근무일수`;
            sheet.getCell(summaryRow, 1).font = { bold: true, size: 11 };
            summaryRow++;

            excelWorkers.forEach((name) => {
                sheet.getCell(summaryRow, 1).value = name;
                sheet.getCell(summaryRow, 1).font = { bold: true, size: 10 };
                sheet.getCell(summaryRow, 2).value = `${monthlyWorkerDays[name]}일`;
                sheet.getCell(summaryRow, 2).font = { size: 10 };
                summaryRow++;
            });

            summaryRow++;
            sheet.getCell(summaryRow, 1).value = '기준 근무일';
            sheet.getCell(summaryRow, 1).font = { bold: true, size: 10 };
            sheet.getCell(summaryRow, 2).value = `${monthRegularDays}일`;
            sheet.getCell(summaryRow, 2).font = { bold: true, size: 10 };

            if (index === 1) {
                let totalRow = summaryStartRow;
                sheet.getCell(totalRow, 8).value = '총 근무일수 합계';
                sheet.getCell(totalRow, 8).font = { bold: true, size: 11 };
                totalRow++;

                excelWorkers.forEach((name) => {
                    sheet.getCell(totalRow, 8).value = name;
                    sheet.getCell(totalRow, 8).font = { bold: true, size: 10 };
                    sheet.getCell(totalRow, 9).value = `${totalWorkerDays[name]}일`;
                    sheet.getCell(totalRow, 9).font = { size: 10 };
                    totalRow++;
                });

                totalRow++;
                sheet.getCell(totalRow, 8).value = '기준 근무일(2개월)';
                sheet.getCell(totalRow, 8).font = { bold: true, size: 10 };
                sheet.getCell(totalRow, 9).value = `${regularWorkDays}일`;
                sheet.getCell(totalRow, 9).font = { bold: true, size: 10 };
            }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${state.year.toString().slice(-2)}년 ${state.monthStart}월,${state.monthStart + 1}월 근무 편성표.xlsx`;
        link.click();
        window.URL.revokeObjectURL(url);
    }

    function openExcelExportModal() {
        syncExcelExportModal();
        excelExportModal.classList.remove('hidden');
        excelExportModal.classList.add('flex');
        requestAnimationFrame(() => {
            excelLibraryNameInput.focus();
        });
    }

    function closeExcelExportModal(force = false) {
        if (isExcelExporting && !force) return;
        excelExportModal.classList.add('hidden');
        excelExportModal.classList.remove('flex');
    }

    function syncExcelExportModal() {
        excelLibraryNameInput.value = excelExportOptions.libraryName;
        excelShowOffWorkersInput.checked = excelExportOptions.showOffWorkers;
        excelShowDetailSettingsInput.checked = excelExportOptions.showDetailedSettings;
    }

    function getExcelExportOptionsFromForm() {
        return {
            libraryName: excelLibraryNameInput.value.trim(),
            showOffWorkers: excelShowOffWorkersInput.checked,
            showDetailedSettings: excelShowDetailSettingsInput.checked
        };
    }

    function setExcelExportLoadingState(isLoading) {
        isExcelExporting = isLoading;
        downloadExcelBtn.innerHTML = isLoading
            ? '<span>⏳ 파일 생성 중...</span>'
            : '<span>최종 엑셀 파일 다운로드</span>';
        downloadExcelBtn.disabled = isLoading;

        confirmExcelExportBtn.textContent = isLoading ? '파일 생성 중...' : '다운로드';
        confirmExcelExportBtn.disabled = isLoading;
        cancelExcelExportBtn.disabled = isLoading;
        closeExcelExportModalBtn.disabled = isLoading;
        excelLibraryNameInput.disabled = isLoading;
        excelShowOffWorkersInput.disabled = isLoading;
        excelShowDetailSettingsInput.disabled = isLoading;
    }

    function getResolvedLibraryName(value) {
        return value ? value.trim() || DEFAULT_LIBRARY_NAME : DEFAULT_LIBRARY_NAME;
    }

    function createWorkerCountMap() {
        return excelWorkers.reduce((acc, name) => {
            acc[name] = 0;
            return acc;
        }, {});
    }

    function getRoomSettingSortOrder(roomValue) {
        if (roomValue === '종합자료실') return 0;
        if (roomValue === '어린이실') return 1;
        if (roomValue === '유아실') return 2;
        return 3;
    }

    function getShiftSettingSortOrder(shiftValue) {
        if (shiftValue === '주간') return 0;
        if (shiftValue === '야간') return 1;
        return 2;
    }

    function isWeekendDate(date) {
        const day = date.getDay();
        return day === 0 || day === 6;
    }

    function shouldUseDetailedScheduleRows(exportOptions) {
        return !!(exportOptions.showDetailedSettings && (exportOptions.useRoomSettings || exportOptions.useShiftSettings));
    }

    function createDefaultScheduleGroup(options = {}) {
        const {
            includeWorkers = true,
            workers = [],
            useDetailedLabel = false
        } = options;

        return {
            key: '__default__',
            label: useDetailedLabel ? '' : '근무자',
            ...(includeWorkers ? { workers: [...workers] } : {}),
            roomSort: 99,
            shiftSort: 99,
            originalIndex: 0
        };
    }

    function getExcelWorkerDetailValues(name, year, month) {
        const detail = getExcelWorkerDetail(name, year, month);
        const weekdayRoom = getExcelWorkerDetailFieldValue(detail, 'weekdayRoom').trim();
        const weekendRoom = getExcelWorkerDetailFieldValue(detail, 'weekendRoom').trim() || weekdayRoom;
        const shiftValue = getExcelWorkerDetailFieldValue(detail, 'shift').trim();

        return {
            detail,
            weekdayRoom,
            weekendRoom,
            shiftValue
        };
    }

    function getExcelWorkerShiftMeta(name, year, month, exportOptions) {
        const useRoomSettings = exportOptions.showDetailedSettings && exportOptions.useRoomSettings;
        const useShiftSettings = exportOptions.showDetailedSettings && exportOptions.useShiftSettings;
        const { weekdayRoom, weekendRoom, shiftValue } = getExcelWorkerDetailValues(name, year, month);
        const hasRoomSetting = useRoomSettings && !!(weekdayRoom || weekendRoom);

        if (useShiftSettings && shiftValue) {
            return {
                value: shiftValue,
                isExplicit: true,
                shouldShowLabel: true
            };
        }

        if (hasRoomSetting) {
            return {
                value: '주간',
                isExplicit: false,
                shouldShowLabel: false
            };
        }

        return {
            value: '야간',
            isExplicit: false,
            shouldShowLabel: useShiftSettings
        };
    }

    function buildWeekendRoomLabelMap(weekDays, year, month, exportOptions) {
        const weekendRoomLabelMap = new Map();

        if (!(exportOptions.showDetailedSettings && exportOptions.useRoomSettings && exportOptions.useShiftSettings)) {
            return weekendRoomLabelMap;
        }

        weekDays.forEach(({ dateStr, date }) => {
            if (isWeekendDate(date)) {
                return;
            }

            const dayData = getExcelDateData(dateStr, false);
            dayData.workers.forEach((name) => {
                const { weekdayRoom } = getExcelWorkerDetailValues(name, year, month);
                const shiftMeta = getExcelWorkerShiftMeta(name, year, month, exportOptions);
                if (weekdayRoom && shiftMeta.isExplicit && shiftMeta.value === '주간' && !weekendRoomLabelMap.has(weekdayRoom)) {
                    weekendRoomLabelMap.set(weekdayRoom, `${weekdayRoom} 주간`);
                }
            });
        });

        return weekendRoomLabelMap;
    }

    function buildExcelWorkerDisplayInfo(name, year, month, exportOptions, originalIndex, date, weekendRoomLabelMap = new Map()) {
        const useRoomSettings = exportOptions.showDetailedSettings && exportOptions.useRoomSettings;
        const useShiftSettings = exportOptions.showDetailedSettings && exportOptions.useShiftSettings;
        const { weekdayRoom, weekendRoom } = getExcelWorkerDetailValues(name, year, month);
        const shiftMeta = getExcelWorkerShiftMeta(name, year, month, exportOptions);
        const weekend = isWeekendDate(date);

        let roomValue = '';
        let label = '';

        if (weekend) {
            roomValue = useRoomSettings ? (weekendRoom || weekdayRoom) : '';

            if (useRoomSettings && roomValue) {
                label = weekendRoomLabelMap.get(roomValue) || roomValue;
            } else if (useShiftSettings && shiftMeta.shouldShowLabel) {
                label = shiftMeta.value;
            }
        } else {
            const labelParts = [];
            roomValue = useRoomSettings ? weekdayRoom : '';

            if (useRoomSettings && roomValue) {
                labelParts.push(roomValue);
            }
            if (useShiftSettings && shiftMeta.shouldShowLabel) {
                labelParts.push(shiftMeta.value);
            }

            label = labelParts.join(' ');
        }

        const hasConfiguredLabel = !!label;
        const effectiveShiftSort = label.endsWith('주간') ? 0 : label.endsWith('야간') ? 1 : 99;

        return {
            name,
            label: hasConfiguredLabel ? label : '',
            hasConfiguredLabel,
            roomSort: hasConfiguredLabel && useRoomSettings ? getRoomSettingSortOrder(roomValue) : 99,
            shiftSort: hasConfiguredLabel ? effectiveShiftSort : 99,
            originalIndex
        };
    }

    function compareExcelWorkerGroups(left, right) {
        if (left.roomSort !== right.roomSort) {
            return left.roomSort - right.roomSort;
        }

        if (left.shiftSort !== right.shiftSort) {
            return left.shiftSort - right.shiftSort;
        }

        return left.originalIndex - right.originalIndex;
    }

    function getExcelWorkerGroupsForDay(workers, year, month, exportOptions, date, weekendRoomLabelMap = new Map()) {
        if (!shouldUseDetailedScheduleRows(exportOptions)) {
            return [createDefaultScheduleGroup({ workers })];
        }

        if (workers.length === 0) {
            return [];
        }

        const groups = new Map();

        workers.forEach((name, index) => {
            const displayInfo = buildExcelWorkerDisplayInfo(name, year, month, exportOptions, index, date, weekendRoomLabelMap);
            const groupKey = displayInfo.hasConfiguredLabel ? displayInfo.label : '__default__';

            if (!groups.has(groupKey)) {
                groups.set(groupKey, {
                    key: groupKey,
                    label: displayInfo.label,
                    workers: [],
                    roomSort: displayInfo.roomSort,
                    shiftSort: displayInfo.shiftSort,
                    originalIndex: displayInfo.originalIndex
                });
            }

            const group = groups.get(groupKey);
            group.workers.push(name);
            group.originalIndex = Math.min(group.originalIndex, displayInfo.originalIndex);
        });

        const groupedWorkers = Array.from(groups.values()).sort(compareExcelWorkerGroups);
        if (groupedWorkers.length > 0) {
            return groupedWorkers;
        }

        return [createDefaultScheduleGroup({ workers, useDetailedLabel: true })];
    }

    function buildWeeklyScheduleRows(weekDays, year, month, exportOptions, weekendRoomLabelMap = new Map()) {
        const groups = new Map();

        weekDays.forEach(({ dateStr, date }) => {
            const dayData = getExcelDateData(dateStr, false);
            getExcelWorkerGroupsForDay(dayData.workers, year, month, exportOptions, date, weekendRoomLabelMap).forEach((group) => {
                if (!groups.has(group.key)) {
                    groups.set(group.key, {
                        key: group.key,
                        label: group.label,
                        roomSort: group.roomSort,
                        shiftSort: group.shiftSort,
                        originalIndex: group.originalIndex
                    });
                }
            });
        });

        const weekGroups = Array.from(groups.values()).sort(compareExcelWorkerGroups);
        if (weekGroups.length > 0) {
            return weekGroups;
        }

        return [createDefaultScheduleGroup({ includeWorkers: false, useDetailedLabel: shouldUseDetailedScheduleRows(exportOptions) })];
    }

    function buildMonthWeekCells(year, month) {
        const weekBlocks = [];
        let currentWeek = [];
        const lastDay = new Date(year, month, 0).getDate();

        for (let day = 1; day <= lastDay; day++) {
            const date = new Date(year, month - 1, day);
            const dayOfWeek = date.getDay();
            const colIdx = dayOfWeek === 0 ? SCHEDULE_END_COLUMN : dayOfWeek + 1;

            currentWeek.push({
                date,
                dateStr: formatDate(date),
                colIdx
            });

            if (colIdx === SCHEDULE_END_COLUMN) {
                weekBlocks.push(currentWeek);
                currentWeek = [];
            }
        }

        if (currentWeek.length > 0) {
            weekBlocks.push(currentWeek);
        }

        return weekBlocks;
    }

    function clearScheduleRow(sheet, rowNumber) {
        for (let col = SCHEDULE_START_COLUMN; col <= SCHEDULE_END_COLUMN; col++) {
            sheet.getCell(rowNumber, col).value = null;
        }
    }

    function writeMergedScheduleTextCell(sheet, startRow, endRow, colIdx, value, options = {}) {
        if (endRow <= startRow) {
            writeScheduleTextCell(sheet.getCell(startRow, colIdx), value, options);
            return;
        }

        for (let rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
            sheet.getCell(rowNumber, colIdx).value = null;
        }

        sheet.mergeCells(startRow, colIdx, endRow, colIdx);
        writeScheduleTextCell(sheet.getCell(startRow, colIdx), value, options);
    }

    function writeMonthlyScheduleArea(sheet, year, month, exportOptions, showOffWorkers, monthlyWorkerDays, totalWorkerDays) {
        const weekBlocks = buildMonthWeekCells(year, month);
        const useDetailedRows = shouldUseDetailedScheduleRows(exportOptions);
        const templateWorkerRowHeight = sheet.getRow(7).height;
        let currentRow = 6;
        let insertedRowCount = 0;
        let lastUsedRow = 7;

        weekBlocks.forEach((weekDays) => {
            const dateRowNumber = currentRow;
            const weekendRoomLabelMap = buildWeekendRoomLabelMap(weekDays, year, month, exportOptions);
            const weekGroups = useDetailedRows
                ? buildWeeklyScheduleRows(weekDays, year, month, exportOptions, weekendRoomLabelMap)
                : [{ key: '__default__', label: '근무자' }];

            setScheduleLabelRow(sheet, dateRowNumber, '일자');
            clearScheduleRow(sheet, dateRowNumber);

            const scheduleRows = [];
            let nextRowNumber = dateRowNumber + 1;

            weekGroups.forEach((group, index) => {
                if (index > 0) {
                    sheet.insertRow(nextRowNumber, [], 'i+');
                    insertedRowCount++;
                    sheet.getRow(nextRowNumber).height = templateWorkerRowHeight;
                }

                clearScheduleRow(sheet, nextRowNumber);
                setScheduleLabelRow(sheet, nextRowNumber, group.label);
                scheduleRows.push({ key: group.key, rowNumber: nextRowNumber });
                nextRowNumber++;
            });

            let offRowNumber = null;
            if (showOffWorkers) {
                sheet.insertRow(nextRowNumber, [], 'i+');
                insertedRowCount++;
                sheet.getRow(nextRowNumber).height = templateWorkerRowHeight;
                clearScheduleRow(sheet, nextRowNumber);
                setScheduleLabelRow(sheet, nextRowNumber, '휴무자', { color: 'FFFF0000' });
                offRowNumber = nextRowNumber;
                nextRowNumber++;
            }

            weekDays.forEach(({ date, dateStr, colIdx }) => {
                const dayData = getExcelDateData(dateStr, showOffWorkers);
                const groupedWorkers = getExcelWorkerGroupsForDay(
                    dayData.workers,
                    year,
                    month,
                    exportOptions,
                    date,
                    weekendRoomLabelMap
                );
                const groupedMap = new Map(groupedWorkers.map((group) => [group.key, group.workers]));
                const mergedEndRow = offRowNumber ?? scheduleRows[scheduleRows.length - 1].rowNumber;
                const specialScheduleLabel = dayData.holidayName || (dayData.isClosed ? '정기휴관일' : '');

                sheet.getCell(dateRowNumber, colIdx).value = new Date(Date.UTC(year, month - 1, date.getDate()));

                if (specialScheduleLabel) {
                    writeMergedScheduleTextCell(
                        sheet,
                        scheduleRows[0].rowNumber,
                        mergedEndRow,
                        colIdx,
                        specialScheduleLabel,
                        { color: 'FFFF0000', bold: true }
                    );
                } else {
                    scheduleRows.forEach(({ key, rowNumber }) => {
                        writeScheduleTextCell(
                            sheet.getCell(rowNumber, colIdx),
                            formatWorkersForExcel(groupedMap.get(key) || [])
                        );
                    });
                }

                dayData.workers.forEach((name) => {
                    monthlyWorkerDays[name]++;
                    totalWorkerDays[name]++;
                });

                if (offRowNumber !== null && !specialScheduleLabel) {
                    writeScheduleTextCell(
                        sheet.getCell(offRowNumber, colIdx),
                        formatWorkersForExcel(dayData.offWorkers),
                        { color: 'FF000000' }
                    );
                }
            });

            scheduleRows.forEach(({ rowNumber }) => {
                setScheduleRowHeight(sheet, rowNumber, 18, 15);
            });

            if (offRowNumber !== null) {
                setScheduleRowHeight(sheet, offRowNumber, 18, 15);
            }

            lastUsedRow = nextRowNumber - 1;
            currentRow = nextRowNumber;
        });

        return {
            lastUsedRow,
            scheduleEndRow: 17 + insertedRowCount
        };
    }

    function calculateRegularWorkDaysForMonth(year, month) {
        let workDays = 0;
        const lastDay = new Date(year, month, 0).getDate();

        for (let day = 1; day <= lastDay; day++) {
            const date = new Date(year, month - 1, day);
            const dayOfWeek = date.getDay();
            if (dayOfWeek === 0 || dayOfWeek === 6) continue;
            if (!getHolidayName(formatDate(date))) {
                workDays++;
            }
        }

        return workDays;
    }

    function calculateRegularWorkDaysForRange() {
        return calculateRegularWorkDaysForMonth(state.year, state.monthStart)
            + calculateRegularWorkDaysForMonth(state.year, state.monthStart + 1);
    }

    function getExcelDateData(dateStr, includeOffWorkers) {
        const source = excelDataByDate[dateStr] || { workers: [], isClosed: false, holidayName: null };
        const workers = [...source.workers];
        const isClosed = !!source.isClosed;
        const holidayName = source.holidayName || null;
        const dayOfWeek = new Date(`${dateStr}T00:00:00`).getDay();
        const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
        const shouldShowOffWorkers = includeOffWorkers && isWeekday && !isClosed && !holidayName;

        return {
            workers,
            isClosed,
            holidayName,
            offWorkers: shouldShowOffWorkers ? excelWorkers.filter((name) => !workers.includes(name)) : []
        };
    }

    function replaceLibraryNameInSheet(sheet, libraryName) {
        sheet.eachRow((row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                const replacedValue = replaceLibraryNameInCellValue(cell.value, libraryName);
                if (replacedValue !== cell.value) {
                    cell.value = replacedValue;
                }
            });
        });
    }

    function replaceLibraryNameInCellValue(cellValue, libraryName) {
        if (typeof cellValue === 'string') {
            return cellValue.includes(DEFAULT_LIBRARY_NAME)
                ? cellValue.split(DEFAULT_LIBRARY_NAME).join(libraryName)
                : cellValue;
        }

        if (cellValue && Array.isArray(cellValue.richText)) {
            let hasReplacement = false;
            const richText = cellValue.richText.map((part) => {
                if (typeof part.text !== 'string' || !part.text.includes(DEFAULT_LIBRARY_NAME)) {
                    return part;
                }

                hasReplacement = true;
                return {
                    ...part,
                    text: part.text.split(DEFAULT_LIBRARY_NAME).join(libraryName)
                };
            });

            return hasReplacement ? { ...cellValue, richText } : cellValue;
        }

        return cellValue;
    }

    function getMonthlyWorkTimeConfig(year, month, exportOptions) {
        let hasDayShift = false;
        let hasNightShift = false;

        excelWorkers.forEach((name) => {
            const shiftMeta = getExcelWorkerShiftMeta(name, year, month, exportOptions);

            if (shiftMeta.value === '야간') {
                hasNightShift = true;
            } else {
                hasDayShift = true;
            }
        });

        if (hasDayShift && hasNightShift) {
            return {
                mainText: MIXED_WORK_TIME_TEXT,
                weekendText: MIXED_WEEKEND_WORK_TIME_TEXT,
                fullText: `${MIXED_WORK_TIME_TEXT}${MIXED_WEEKEND_WORK_TIME_TEXT}`
            };
        }

        if (hasNightShift) {
            return {
                mainText: NIGHT_WORK_TIME_TEXT,
                weekendText: NIGHT_WEEKEND_WORK_TIME_TEXT,
                fullText: `${NIGHT_WORK_TIME_TEXT}${NIGHT_WEEKEND_WORK_TIME_TEXT}`
            };
        }

        return {
            mainText: DAY_WORK_TIME_TEXT,
            weekendText: '',
            fullText: DAY_WORK_TIME_TEXT
        };
    }

    function getExcelSheetTitle(year, month, libraryName, exportOptions) {
        const workTimeConfig = getMonthlyWorkTimeConfig(year, month, exportOptions);
        const scheduleLabel = workTimeConfig.mainText === NIGHT_WORK_TIME_TEXT
            ? '야간근무 편성표'
            : '근무 편성표';

        return `${year}년 ${String(month).padStart(2, '0')}월 ${libraryName} ${scheduleLabel}`;
    }

    function replaceWorkTimeInSheet(sheet, year, month, exportOptions) {
        const workTimeConfig = getMonthlyWorkTimeConfig(year, month, exportOptions);

        sheet.eachRow((row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                const replacedValue = replaceWorkTimeInCellValue(cell.value, workTimeConfig);
                if (replacedValue !== cell.value) {
                    cell.value = replacedValue;
                }
            });
        });
    }

    function replaceWorkTimeInCellValue(cellValue, workTimeConfig) {
        if (typeof cellValue === 'string') {
            return cellValue.includes('근무시간:')
                ? cellValue.replace(/(근무시간:\s*).*/, `$1${workTimeConfig.fullText}`)
                : cellValue;
        }

        if (cellValue && Array.isArray(cellValue.richText)) {
            let hasReplacement = false;

            const richText = cellValue.richText
                .map((part) => {
                    if (typeof part.text !== 'string') {
                        return part;
                    }

                    if (part.text.includes('근무시간:')) {
                        hasReplacement = true;
                        return {
                            ...part,
                            text: part.text.replace(/(근무시간:\s*).*/, `$1${workTimeConfig.mainText}`)
                        };
                    }

                    if (part.text.includes('주말')) {
                        hasReplacement = true;
                        return {
                            ...part,
                            text: workTimeConfig.weekendText
                        };
                    }

                    return part;
                })
                .filter((part) => typeof part.text !== 'string' || part.text !== '');

            return hasReplacement ? { ...cellValue, richText } : cellValue;
        }

        return cellValue;
    }

    function ensureOffWorkerRows(sheet) {
        const originalWorkerRows = [7, 9, 11, 13, 15, 17];

        for (let idx = originalWorkerRows.length - 1; idx >= 0; idx--) {
            const workerRow = originalWorkerRows[idx];
            const offRow = workerRow + 1;

            sheet.insertRow(offRow, [], 'i+');
            sheet.getRow(offRow).height = sheet.getRow(workerRow).height;

            const labelCell = sheet.getCell(offRow, 1);

            for (let col = SCHEDULE_START_COLUMN; col <= SCHEDULE_END_COLUMN; col++) {
                sheet.getCell(offRow, col).value = null;
            }

            labelCell.value = '휴무자';
            labelCell.style = {
                ...labelCell.style,
                font: {
                    ...(labelCell.style?.font || {}),
                    bold: true,
                    color: { argb: 'FFFF0000' }
                },
                alignment: {
                    ...(labelCell.style?.alignment || {}),
                    vertical: 'middle',
                    horizontal: 'center'
                }
            };
        }
    }

    function normalizeScheduleLabelRows(sheet, showOffWorkers) {
        const rowStride = showOffWorkers ? 3 : 2;
        const templateLastRow = 17 + (showOffWorkers ? 6 : 0);

        for (let row = 6; row <= templateLastRow; row += rowStride) {
            setScheduleLabelRow(sheet, row, '일자');
            setScheduleLabelRow(sheet, row + 1, '근무자');
            if (showOffWorkers) {
                setScheduleLabelRow(sheet, row + 2, '휴무자', { color: 'FFFF0000' });
            }
        }
    }

    function setScheduleLabelRow(sheet, rowNumber, text, options = {}) {
        const labelCell = sheet.getCell(rowNumber, 1);
        const baseStyle = labelCell.style || {};
        const baseFont = baseStyle.font || {};

        labelCell.value = text;
        labelCell.style = {
            ...baseStyle,
            font: {
                ...baseFont,
                bold: true,
                color: { argb: options.color || baseFont.color?.argb || 'FF000000' }
            },
            alignment: {
                ...(baseStyle.alignment || {}),
                horizontal: 'center',
                vertical: 'middle'
            }
        };
    }

    function clearScheduleArea(sheet, templateLastRow) {
        for (let row = 6; row <= templateLastRow; row++) {
            for (let col = SCHEDULE_START_COLUMN; col <= SCHEDULE_END_COLUMN; col++) {
                sheet.getCell(row, col).value = null;
            }
        }
    }

    function writeScheduleTextCell(cell, value, options = {}) {
        const existingBorder = cell.border;
        cell.value = value;
        cell.style = {
            font: {
                color: { argb: options.color || 'FF000000' },
                bold: !!options.bold,
                size: 10
            },
            alignment: {
                wrapText: true,
                vertical: 'middle',
                horizontal: 'center'
            },
            border: existingBorder
        };
    }

    function adjustScheduleRowHeights(sheet, lastUsedRow, showOffWorkers, rowStride) {
        const baseRowHeight = 18;
        const lineHeight = 15;

        for (let dateRow = 6; dateRow <= lastUsedRow; dateRow += rowStride) {
            setScheduleRowHeight(sheet, dateRow + 1, baseRowHeight, lineHeight);
            if (showOffWorkers) {
                setScheduleRowHeight(sheet, dateRow + 2, baseRowHeight, lineHeight);
            }
        }
    }

    function setScheduleRowHeight(sheet, rowNumber, baseRowHeight, lineHeight) {
        let maxLines = 1;
        for (let col = SCHEDULE_START_COLUMN; col <= SCHEDULE_END_COLUMN; col++) {
            const value = sheet.getCell(rowNumber, col).value;
            if (typeof value === 'string' && value) {
                maxLines = Math.max(maxLines, value.split('\n').length);
            }
        }

        sheet.getRow(rowNumber).height = maxLines > 1
            ? baseRowHeight + (maxLines - 1) * lineHeight
            : 24.95;
    }

    function createExcelDetailSettingSelect(detailType, workerIndex, monthKey, value) {
        const select = document.createElement('select');
        select.dataset.workerIndex = String(workerIndex);
        select.dataset.monthKey = monthKey;
        select.dataset.detailType = detailType;
        select.className = 'w-full rounded-xl border border-teal-100 bg-white px-3 py-2 text-xs text-stone-700 focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-100';

        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.textContent = '미설정';
        select.appendChild(emptyOption);

        if (isRoomDetailType(detailType)) {
            ROOM_SETTING_OPTIONS.forEach((roomName) => {
                const option = document.createElement('option');
                option.value = roomName;
                option.textContent = roomName;
                select.appendChild(option);
            });

            if (value && !ROOM_SETTING_OPTIONS.includes(value)) {
                const customOption = document.createElement('option');
                customOption.value = value;
                customOption.textContent = `직접입력: ${value}`;
                select.appendChild(customOption);
            }

            const promptOption = document.createElement('option');
            promptOption.value = CUSTOM_ROOM_SETTING_VALUE;
            promptOption.textContent = '직접입력...';
            select.appendChild(promptOption);
        } else {
            SHIFT_SETTING_OPTIONS.forEach((shiftName) => {
                const option = document.createElement('option');
                option.value = shiftName;
                option.textContent = shiftName;
                select.appendChild(option);
            });
        }

        select.value = value || '';
        return select;
    }

    function renderExcelDetailSettingsPanel() {
        const showDetailedSettings = excelShowDetailSettingsInput.checked;
        excelDetailSettingsSection.classList.toggle('hidden', !showDetailedSettings);

        if (!showDetailedSettings) {
            excelDetailSettingsHead.innerHTML = '';
            excelDetailSettingsBody.innerHTML = '';
            return;
        }

        ensureAllExcelWorkerDetailSettings();

        const monthConfigs = getExcelDetailMonthConfigs();
        const detailColumns = getActiveExcelDetailSettingColumns();
        const shouldShowTable = excelWorkers.length > 0 && detailColumns.length > 0;
        const monthThemeClasses = [
            {
                header: 'bg-amber-100',
                subHeader: 'bg-amber-50',
                body: 'bg-amber-50',
                divider: 'border-l-[3px] border-l-amber-300'
            },
            {
                header: 'bg-emerald-100',
                subHeader: 'bg-emerald-50',
                body: 'bg-emerald-50',
                divider: 'border-l-[3px] border-l-emerald-300'
            },
            {
                header: 'bg-sky-100',
                subHeader: 'bg-sky-50',
                body: 'bg-sky-50',
                divider: 'border-l-[3px] border-l-sky-300'
            }
        ];

        excelDetailSettingsNote.textContent = '';

        excelDetailSettingsTable.classList.toggle('hidden', !shouldShowTable);
        excelDetailSettingsEmpty.classList.toggle('hidden', shouldShowTable);

        if (!shouldShowTable) {
            excelDetailSettingsHead.innerHTML = '';
            excelDetailSettingsBody.innerHTML = '';
            excelDetailSettingsEmpty.textContent = excelWorkers.length === 0
                ? '먼저 엑셀 근무자 명단에 사람을 추가하세요.'
                : '자료실 설정 또는 주야간 설정을 하나 이상 켜세요.';
            return;
        }

        excelDetailSettingsHead.innerHTML = '';
        excelDetailSettingsBody.innerHTML = '';

        const headerRow = document.createElement('tr');
        const workerHeader = document.createElement('th');
        workerHeader.rowSpan = 2;
        workerHeader.className = 'sticky top-0 z-30 px-3 py-3 text-left text-xs font-bold text-stone-700 border-b border-r border-teal-300 bg-white shadow-[0_1px_0_0_rgba(20,83,45,0.08)] w-[4.4rem] min-w-[4.4rem] whitespace-nowrap align-middle';
        workerHeader.textContent = '근무자';
        headerRow.appendChild(workerHeader);

        monthConfigs.forEach(({ label }, monthIndex) => {
            const theme = monthThemeClasses[monthIndex % monthThemeClasses.length];
            const monthHeader = document.createElement('th');
            monthHeader.colSpan = detailColumns.length;
            monthHeader.className = `sticky top-0 z-20 px-3 py-3 text-center text-xs font-bold text-stone-700 border-b border-r border-teal-300 shadow-[0_1px_0_0_rgba(20,83,45,0.08)] ${theme.header} ${theme.divider}`;
            monthHeader.textContent = label;
            headerRow.appendChild(monthHeader);
        });

        const subHeaderRow = document.createElement('tr');
        monthConfigs.forEach((_, monthIndex) => {
            const theme = monthThemeClasses[monthIndex % monthThemeClasses.length];
            detailColumns.forEach(({ label }, columnIndex) => {
                const subHeader = document.createElement('th');
                subHeader.className = `sticky top-[43px] z-20 px-3 py-2 text-center text-[11px] font-semibold text-stone-600 border-b border-r border-teal-300 shadow-[0_1px_0_0_rgba(20,83,45,0.08)] ${theme.subHeader} ${columnIndex === 0 ? theme.divider : ''}`;
                subHeader.textContent = label;
                subHeaderRow.appendChild(subHeader);
            });
        });

        excelDetailSettingsHead.appendChild(headerRow);
        excelDetailSettingsHead.appendChild(subHeaderRow);

        excelWorkers.forEach((name, workerIndex) => {
            ensureExcelWorkerDetailSettings(name);

            const row = document.createElement('tr');
            if (workerIndex % 2 === 1) {
                row.className = 'bg-stone-50/40';
            }

            const nameCell = document.createElement('th');
            nameCell.className = `px-3 py-3 text-left text-sm font-semibold text-stone-800 border-r border-teal-300 whitespace-nowrap w-[4.4rem] min-w-[4.4rem] ${workerIndex % 2 === 1 ? 'bg-stone-50' : 'bg-white'}`;
            nameCell.textContent = name;
            row.appendChild(nameCell);

            monthConfigs.forEach(({ key }, monthIndex) => {
                const theme = monthThemeClasses[monthIndex % monthThemeClasses.length];
                const detail = excelWorkerDetails[name][key];

                detailColumns.forEach(({ key: columnKey }, columnIndex) => {
                    const cell = document.createElement('td');
                    cell.className = `px-3 py-2 border-r border-b border-teal-300 ${theme.body} ${columnIndex === 0 ? theme.divider : ''}`;
                    cell.appendChild(
                        createExcelDetailSettingSelect(
                            columnKey,
                            workerIndex,
                            key,
                            getExcelWorkerDetailFieldValue(detail, columnKey)
                        )
                    );
                    row.appendChild(cell);
                });
            });

            excelDetailSettingsBody.appendChild(row);
        });
    }

    function openExcelExportModal() {
        syncExcelExportModal();
        excelExportModal.classList.remove('hidden');
        excelExportModal.classList.add('flex');
        requestAnimationFrame(() => {
            excelLibraryNameInput.focus();
        });
    }

    function closeExcelExportModal(force = false) {
        if (isExcelExporting && !force) return;
        excelExportModal.classList.add('hidden');
        excelExportModal.classList.remove('flex');
    }

    function syncExcelExportModal() {
        excelLibraryNameInput.value = excelExportOptions.libraryName;
        excelShowOffWorkersInput.checked = excelExportOptions.showOffWorkers;
        excelShowDetailSettingsInput.checked = excelExportOptions.showDetailedSettings;
        excelEnableRoomSettingsInput.checked = excelExportOptions.useRoomSettings;
        excelEnableShiftSettingsInput.checked = excelExportOptions.useShiftSettings;
        renderExcelDetailSettingsPanel();
    }

    function getExcelExportOptionsFromForm() {
        return {
            libraryName: excelLibraryNameInput.value.trim(),
            showOffWorkers: excelShowOffWorkersInput.checked,
            showDetailedSettings: excelShowDetailSettingsInput.checked,
            useRoomSettings: excelEnableRoomSettingsInput.checked,
            useShiftSettings: excelEnableShiftSettingsInput.checked
        };
    }

    function setExcelExportLoadingState(isLoading) {
        isExcelExporting = isLoading;
        downloadExcelBtn.innerHTML = isLoading
            ? '<span>⏳ 파일 생성 중...</span>'
            : '<span>최종 엑셀 파일 다운로드</span>';
        downloadExcelBtn.disabled = isLoading;

        confirmExcelExportBtn.textContent = isLoading ? '파일 생성 중...' : '다운로드';
        confirmExcelExportBtn.disabled = isLoading;
        cancelExcelExportBtn.disabled = isLoading;
        closeExcelExportModalBtn.disabled = isLoading;
        excelLibraryNameInput.disabled = isLoading;
        excelShowOffWorkersInput.disabled = isLoading;
        excelShowDetailSettingsInput.disabled = isLoading;
        excelEnableRoomSettingsInput.disabled = isLoading;
        excelEnableShiftSettingsInput.disabled = isLoading;

        excelDetailSettingsBody.querySelectorAll('select').forEach((select) => {
            select.disabled = isLoading;
        });
    }

    function updateExcelWorkersUI() {
        excelCountEl.innerText = excelWorkers.length;
        excelWorkersList.innerHTML = '';

        excelWorkers.forEach((name) => {
            const span = document.createElement('span');
            span.className = 'bg-teal-700 text-teal-50 px-2 py-0.5 rounded border border-teal-500 text-[11px] font-bold flex items-center gap-1 shadow-sm';

            const text = document.createElement('span');
            text.textContent = name;

            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'text-teal-300 hover:text-white transition-colors';
            button.textContent = '×';
            button.addEventListener('click', () => {
                window.removeExcelWorker(name);
            });

            span.appendChild(text);
            span.appendChild(button);
            excelWorkersList.appendChild(span);
        });

        if (excelExportModal.classList.contains('flex')) {
            renderExcelDetailSettingsPanel();
        }
    }

    window.removeExcelWorker = function(nameToRemove) {
        excelWorkers = excelWorkers.filter((name) => name !== nameToRemove);
        delete excelWorkerDetails[nameToRemove];

        for (const dateStr in excelDataByDate) {
            excelDataByDate[dateStr].workers = excelDataByDate[dateStr].workers.filter((name) => name !== nameToRemove);
        }

        updateExcelWorkersUI();
    };

    function formatWorkersForExcel(workers) {
        if (!workers || workers.length === 0) return '';

        const lines = [];
        for (let index = 0; index < workers.length; index += 2) {
            const lineWorkers = workers.slice(index, index + 2);
            lines.push(lineWorkers.join(','));
        }

        return lines.join('\n');
    }

    function renderHolidayList() {
        holidayListUl.innerHTML = '';
        holidays
            .sort((a, b) => a.date.localeCompare(b.date))
            .forEach((holiday) => {
                const isDefault = (typeof DEFAULT_HOLIDAYS !== 'undefined')
                    && DEFAULT_HOLIDAYS.find((item) => item.date === holiday.date && item.name === holiday.name);

                const li = document.createElement('li');
                li.className = 'flex justify-between items-center bg-white border border-stone-100 px-3 py-2 rounded-xl shadow-sm';

                const tagHtml = isDefault
                    ? '<span class="text-[10px] text-stone-500 font-bold px-1.5 py-0.5 rounded bg-stone-100 border border-stone-200">기본</span>'
                    : '<span class="text-[10px] text-amber-600 font-bold px-1.5 py-0.5 rounded bg-amber-50 border border-amber-200">추가됨</span>';

                li.innerHTML = `
                    <div class="flex items-center gap-2">
                        <span class="font-bold text-stone-700 w-[85px]">${holiday.date}</span>
                        <span class="text-stone-500 text-xs truncate max-w-[100px]">${holiday.name}</span>
                        ${tagHtml}
                    </div>
                    <button onclick="window.removeHoliday('${holiday.date}')" class="text-stone-400 hover:text-rose-500 hover:bg-rose-50 w-7 h-7 flex items-center justify-center rounded-lg transition-colors font-bold text-lg" title="삭제">&times;</button>
                `;
                holidayListUl.appendChild(li);
            });
    }

    window.removeHoliday = async function(dateStr) {
        if (!(await confirmExcelScheduleReset('휴일 목록'))) {
            return;
        }

        if (excelWorkers.length > 0) {
            resetExcelWorkerData();
        }

        const isDefault = (typeof DEFAULT_HOLIDAYS !== 'undefined')
            && DEFAULT_HOLIDAYS.find((item) => item.date === dateStr);

        if (isDefault) {
            if (!deletedDefaults.includes(dateStr)) {
                deletedDefaults.push(dateStr);
            }
        } else {
            customHolidays = customHolidays.filter((holiday) => holiday.date !== dateStr);
        }

        saveHolidaysToLocalStorage();
    };

    function updatePersonalSettingsUI() {
        personalContainer.innerHTML = '';

        const targetMonths = [];
        let previousMonth = state.monthStart - 1;
        let previousYear = state.year;
        if (previousMonth === 0) {
            previousMonth = 12;
            previousYear--;
        }

        targetMonths.push({ y: previousYear, m: previousMonth, label: `${previousMonth}월(이전)` });
        targetMonths.push({ y: state.year, m: state.monthStart, label: `${state.monthStart}월` });
        targetMonths.push({ y: state.year, m: state.monthStart + 1, label: `${state.monthStart + 1}월` });

        targetMonths.forEach((targetMonth) => {
            const key = `${targetMonth.y}-${targetMonth.m}`;
            if (!state.personalSettings[key]) {
                state.personalSettings[key] = { weekday: 2, weekend: 6 };
            }

            const setting = state.personalSettings[key];
            const div = document.createElement('div');
            div.className = 'personal-month-row flex items-center gap-2 mb-1';
            div.dataset.key = key;
            div.innerHTML = `
                <span class="w-16 font-bold text-stone-700 text-right mr-1">${targetMonth.label}</span>
                <select class="weekday-sel flex-1 p-2 rounded-xl border-2 border-stone-200 bg-white focus:ring-0 focus:border-amber-500 outline-none text-sm shadow-sm cursor-pointer">
                    <option value="1" ${setting.weekday === 1 ? 'selected' : ''}>월요일 휴무</option>
                    <option value="2" ${setting.weekday === 2 ? 'selected' : ''}>화요일 휴무</option>
                    <option value="3" ${setting.weekday === 3 ? 'selected' : ''}>수요일 휴무</option>
                    <option value="4" ${setting.weekday === 4 ? 'selected' : ''}>목요일 휴무</option>
                    <option value="5" ${setting.weekday === 5 ? 'selected' : ''}>금요일 휴무</option>
                </select>
                <select class="weekend-sel flex-1 p-2 rounded-xl border-2 border-stone-200 bg-white focus:ring-0 focus:border-amber-500 outline-none text-sm shadow-sm cursor-pointer">
                    <option value="6" ${setting.weekend === 6 ? 'selected' : ''}>토요일 휴무</option>
                    <option value="0" ${setting.weekend === 0 ? 'selected' : ''}>일요일 휴무</option>
                </select>
            `;
            personalContainer.appendChild(div);
        });
    }

    function getWeekMonday(date) {
        const monday = new Date(date);
        const day = monday.getDay() === 0 ? 7 : monday.getDay();
        monday.setDate(monday.getDate() - day + 1);
        return monday;
    }

    function getHolidayName(dateStr) {
        const holiday = holidays.find((item) => item.date === dateStr);
        return holiday ? holiday.name : null;
    }

    function formatDate(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function getPublicOfficialDayInfo(date) {
        const weekMonday = getWeekMonday(date);
        const schedY = weekMonday.getFullYear();
        const schedM = weekMonday.getMonth() + 1;
        const key = `${schedY}-${schedM}`;

        const personalSetting = state.personalSettings[key] || { weekday: 2, weekend: 6 };
        const mondayNum = Math.floor((weekMonday.getDate() - 1) / 7) + 1;

        let isLibraryClosedWeek = false;
        if (state.libraryClosed === 'odd' && (mondayNum === 1 || mondayNum === 3)) {
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
        } else if (dateDay === personalSetting.weekday) {
            isOff = true;
            reason = '평일휴무';
        }

        if (dateDay === personalSetting.weekend) {
            isOff = true;
            reason = '주말휴무';
        }

        const holidayName = getHolidayName(formatDate(date));
        if (holidayName) {
            isOff = true;
            reason = holidayName;
        }

        return { isOff, reason, isHoliday: !!holidayName };
    }

    function getRegularDayInfo(date) {
        const day = date.getDay();
        const holidayName = getHolidayName(formatDate(date));

        let isOff = false;
        let reason = '';
        let isHoliday = false;

        if (day === 0 || day === 6) {
            isOff = true;
            reason = '주말';
        }

        if (holidayName) {
            isOff = true;
            reason = holidayName;
            isHoliday = true;
        }

        return { isOff, reason, isHoliday };
    }

    function renderAll() {
        updateRunButtonState();

        const titleEl = document.getElementById('target-employee-type');
        if (state.mode === 'regular') {
            titleEl.innerText = '기준 근무일 (미적용 상태)';
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

        totalEl.innerText = state.mode === 'regular' ? totalRegularWorkDays : totalPublicWorkDays;
        regTotalEl.innerText = totalRegularWorkDays;

        totalEl.style.transform = 'scale(1.15)';
        setTimeout(() => {
            totalEl.style.transform = 'scale(1)';
        }, 200);
    }

    function renderCalendar(calIndex, year, month) {
        document.getElementById(`cal-title-${calIndex}`).innerText = `${year}년 ${month}월`;
        const container = document.getElementById(`cal-days-${calIndex}`);
        container.innerHTML = '';

        const firstDay = new Date(year, month - 1, 1);
        const lastDay = new Date(year, month, 0);

        for (let i = 0; i < firstDay.getDay(); i++) {
            const empty = document.createElement('div');
            empty.className = 'day-cell empty min-h-[4.5rem]';
            container.appendChild(empty);
        }

        let publicDays = 0;
        let regularDays = 0;

        for (let day = 1; day <= lastDay.getDate(); day++) {
            const date = new Date(year, month - 1, day);
            const cell = document.createElement('div');
            cell.className = 'day-cell min-h-[4.5rem] flex flex-col items-center justify-start rounded-2xl font-medium text-sm relative p-1.5 transition-all hover:-translate-y-1 hover:shadow-lg hover:z-10 cursor-default shadow-sm border-2';

            const regularInfo = getRegularDayInfo(date);
            const publicInfo = getPublicOfficialDayInfo(date);

            if (!regularInfo.isOff) regularDays++;
            if (!publicInfo.isOff) publicDays++;

            const info = state.mode === 'regular' ? regularInfo : publicInfo;

            if (info.isOff) {
                if (info.isHoliday || info.reason.includes('휴관')) {
                    cell.classList.add('bg-hol', 'text-holText', 'border-holBorder');
                } else {
                    cell.classList.add('bg-off', 'text-offText', 'border-offBorder');
                }
            } else {
                cell.classList.add('bg-work', 'text-workText', 'border-workBorder');
            }

            let html = `<div class="font-bold text-base mt-0.5">${day}</div>`;
            if (info.reason) {
                html += `<div class="text-[0.6rem] sm:text-[0.65rem] leading-tight mt-1 opacity-85 text-center break-words w-full flex-1 px-0.5">${info.reason}</div>`;
            }

            cell.innerHTML = html;
            container.appendChild(cell);
        }

        return { publicDays, regularDays };
    }
});
