document.addEventListener('DOMContentLoaded', () => {
    const DEFAULT_LIBRARY_NAME = '노을빛도서관';
    const SCHEDULE_START_COLUMN = 2;
    const SCHEDULE_END_COLUMN = 8;

    let holidays = [];
    let deletedDefaults = [];
    let customHolidays = [];

    let excelWorkers = [];
    let excelDataByDate = {};
    let isExcelExporting = false;
    let excelExportOptions = {
        libraryName: '',
        showOffWorkers: false,
        showDetailedSettings: false
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

    const holidayModal = document.getElementById('holiday-modal');
    const openHolidayModalBtn = document.getElementById('open-holiday-modal-btn');
    const closeHolidayModalBtn = document.getElementById('close-holiday-modal-btn');
    const holidayDateInput = document.getElementById('new-holiday-date');
    const holidayNameInput = document.getElementById('new-holiday-name');
    const addHolidayBtn = document.getElementById('add-holiday-btn');
    const holidayListUl = document.getElementById('holiday-list');

    const alertModal = document.getElementById('alert-modal');
    const closeAlertBtn = document.getElementById('close-alert-btn');

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
        renderAll();
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

    function setupEventListeners() {
        yearSelect.addEventListener('change', (e) => {
            const newYear = parseInt(e.target.value, 10);
            if (excelWorkers.length > 0) {
                const confirmed = confirm('대상 연도가 변경됩니다.\n저장된 엑셀 근무자 명단이 모두 삭제됩니다.\n\n계속하시겠습니까?');
                if (!confirmed) {
                    yearSelect.value = state.year;
                    return;
                }
                excelWorkers = [];
                excelDataByDate = {};
                updateExcelWorkersUI();
            }

            state.year = newYear;
            state.mode = 'regular';
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            renderAll();
        });

        monthPairSelect.addEventListener('change', (e) => {
            const newMonth = parseInt(e.target.value, 10);
            if (excelWorkers.length > 0) {
                const confirmed = confirm('대상 월이 변경됩니다.\n저장된 엑셀 근무자 명단이 모두 삭제됩니다.\n\n계속하시겠습니까?');
                if (!confirmed) {
                    monthPairSelect.value = state.monthStart;
                    return;
                }
                excelWorkers = [];
                excelDataByDate = {};
                updateExcelWorkersUI();
            }

            state.monthStart = newMonth;
            state.mode = 'regular';
            updatePersonalSettingsUI();
            setHolidayInputDefaultDate();
            renderAll();
        });

        libraryClosedSelect.addEventListener('change', (e) => {
            state.libraryClosed = e.target.value;
            saveSettingsToLocalStorage();
        });

        resetBtn.addEventListener('click', () => {
            state.personalSettings = {};
            updatePersonalSettingsUI();
            state.mode = 'regular';
            renderAll();
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
                alert('먼저 공무직 휴관일 달력을 설정한 뒤 확인하기 버튼을 눌러주세요.\n기준 근무일 상태에서는 근무자를 추가할 수 없습니다.');
                return;
            }

            const name = workerNameInput.value.trim();
            if (!name) {
                alert('근무자 이름을 입력해주세요.');
                workerNameInput.focus();
                return;
            }
            if (excelWorkers.includes(name)) {
                alert('이미 추가된 근무자입니다.');
                return;
            }
            if (!state.libraryClosed) {
                alert('공무직 도서관 휴관일을 먼저 선택해주세요.');
                libraryClosedSelect.focus();
                return;
            }

            const rows = document.querySelectorAll('.personal-month-row');
            rows.forEach((row) => {
                const key = row.dataset.key;
                const weekday = parseInt(row.querySelector('.weekday-sel').value, 10);
                const weekend = parseInt(row.querySelector('.weekend-sel').value, 10);
                state.personalSettings[key] = { weekday, weekend };
            });

            excelWorkers.push(name);
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
                alert('엑셀에 추가된 인원이 없습니다. 먼저 명단에 추가해주세요.');
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
                alert('엑셀 파일을 생성하는 중 오류가 발생했습니다.');
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

        runBtn.addEventListener('click', () => {
            if (!state.libraryClosed) {
                alertModal.classList.remove('hidden');
                alertModal.classList.add('flex');
                libraryClosedSelect.focus();
                libraryClosedSelect.classList.add('ring-4', 'ring-blue-200', 'border-blue-500');
                setTimeout(() => {
                    libraryClosedSelect.classList.remove('ring-4', 'ring-blue-200', 'border-blue-500');
                }, 1500);
                return;
            }

            const rows = document.querySelectorAll('.personal-month-row');
            rows.forEach((row) => {
                const key = row.dataset.key;
                const weekday = parseInt(row.querySelector('.weekday-sel').value, 10);
                const weekend = parseInt(row.querySelector('.weekend-sel').value, 10);
                state.personalSettings[key] = { weekday, weekend };
            });

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
            if (!holidayDateInput.value || !holidayNameInput.value) return;

            const dateStr = holidayDateInput.value;
            const nameStr = holidayNameInput.value;

            if (holidays.find((holiday) => holiday.date === dateStr)) {
                alert('이미 등록된 휴일 날짜입니다.');
                return;
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
            const lastDay = new Date(year, month, 0).getDate();
            const showOffWorkers = exportOptions.showOffWorkers;
            const rowStride = showOffWorkers ? 3 : 2;
            const templateLastRow = 17 + (showOffWorkers ? 6 : 0);

            if (showOffWorkers) {
                ensureOffWorkerRows(sheet);
            }

            normalizeScheduleLabelRows(sheet, showOffWorkers);

            clearScheduleArea(sheet, templateLastRow);

            sheet.name = `${year.toString().slice(-2)}년 ${month}월`;
            sheet.getCell('A1').value = `${year}년 ${String(month).padStart(2, '0')}월 ${libraryName} 야간근무 편성표`;
            replaceLibraryNameInSheet(sheet, libraryName);

            const monthlyWorkerDays = createWorkerCountMap();
            let currentRow = 6;
            let lastUsedRow = showOffWorkers ? 8 : 7;

            for (let day = 1; day <= lastDay; day++) {
                const date = new Date(year, month - 1, day);
                const dateStr = formatDate(date);
                const dayOfWeek = date.getDay();
                const colIdx = dayOfWeek === 0 ? SCHEDULE_END_COLUMN : dayOfWeek + 1;
                const dayData = getExcelDateData(dateStr, showOffWorkers);

                sheet.getCell(currentRow, colIdx).value = new Date(Date.UTC(year, month - 1, day));

                const workerCell = sheet.getCell(currentRow + 1, colIdx);
                if (dayData.holidayName) {
                    writeScheduleTextCell(workerCell, dayData.holidayName, { color: 'FFFF0000', bold: true });
                } else if (dayData.isClosed) {
                    writeScheduleTextCell(workerCell, '휴관일', { color: 'FFFF0000', bold: true });
                } else {
                    writeScheduleTextCell(workerCell, formatWorkersForExcel(dayData.workers));
                }

                dayData.workers.forEach((name) => {
                    monthlyWorkerDays[name]++;
                    totalWorkerDays[name]++;
                });

                if (showOffWorkers) {
                    writeScheduleTextCell(
                        sheet.getCell(currentRow + 2, colIdx),
                        formatWorkersForExcel(dayData.offWorkers),
                        { color: 'FF000000' }
                    );
                }

                lastUsedRow = currentRow + (showOffWorkers ? 2 : 1);
                if (colIdx === SCHEDULE_END_COLUMN) {
                    currentRow += rowStride;
                }
            }

            adjustScheduleRowHeights(sheet, lastUsedRow, showOffWorkers, rowStride);

            if (lastUsedRow < templateLastRow) {
                for (let row = lastUsedRow + 1; row <= templateLastRow; row++) {
                    for (let col = 1; col <= 9; col++) {
                        sheet.getCell(row, col).value = '';
                    }
                }

                const rowsToDelete = templateLastRow - lastUsedRow;
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
                sheet.getCell(totalRow, 8).value = '기준 근무일 (2개월)';
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
                if (typeof cell.value === 'string' && cell.value.includes(DEFAULT_LIBRARY_NAME)) {
                    cell.value = cell.value.split(DEFAULT_LIBRARY_NAME).join(libraryName);
                }
            });
        });
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

    function updateExcelWorkersUI() {
        excelCountEl.innerText = excelWorkers.length;
        excelWorkersList.innerHTML = '';

        excelWorkers.forEach((name) => {
            const span = document.createElement('span');
            span.className = 'bg-teal-700 text-teal-50 px-2 py-0.5 rounded border border-teal-500 text-[11px] font-bold flex items-center gap-1 shadow-sm';
            span.innerHTML = `${name} <button onclick="window.removeExcelWorker('${name}')" class="text-teal-300 hover:text-white transition-colors">&times;</button>`;
            excelWorkersList.appendChild(span);
        });
    }

    window.removeExcelWorker = function(nameToRemove) {
        excelWorkers = excelWorkers.filter((name) => name !== nameToRemove);
        for (const dateStr in excelDataByDate) {
            excelDataByDate[dateStr].workers = excelDataByDate[dateStr].workers.filter((name) => name !== nameToRemove);
        }
        updateExcelWorkersUI();
    };

    function formatWorkersForExcel(workers) {
        if (!workers || workers.length === 0) return '';

        const lines = [];
        for (let i = 0; i < workers.length; i += 2) {
            if (i + 1 < workers.length) {
                lines.push(`${workers[i]},${workers[i + 1]}`);
            } else {
                lines.push(workers[i]);
            }
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

    window.removeHoliday = function(dateStr) {
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
        const titleEl = document.getElementById('target-employee-type');
        if (state.mode === 'regular') {
            titleEl.innerText = '기준 근무일 (설정 전)';
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
