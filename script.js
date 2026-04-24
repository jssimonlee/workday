document.addEventListener('DOMContentLoaded', () => {
    // State
    let holidays = [];
    let deletedDefaults = [];
    let customHolidays = [];
    
    // Excel State
    let excelWorkers = [];
    let excelDataByDate = {}; 
    
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
    
    // Excel Elements
    const workerNameInput = document.getElementById('worker-name-input');
    const addExcelBtn = document.getElementById('add-excel-btn');
    const downloadExcelBtn = document.getElementById('download-excel-btn');
    const excelCountEl = document.getElementById('excel-count');
    const excelWorkersList = document.getElementById('excel-workers-list');
    const excelToggleBtn = document.getElementById('excel-toggle-btn');
    const excelToggleArrow = document.getElementById('excel-toggle-arrow');
    const excelPanel = document.getElementById('excel-panel');
    
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
            const newYear = parseInt(e.target.value);
            // 엑셀에 추가된 인원이 있으면 경고
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
            const newMonth = parseInt(e.target.value);
            // 엑셀에 추가된 인원이 있으면 경고
            if (excelWorkers.length > 0) {
                const confirmed = confirm('대상 월이 변경됩니다.\n저장된 엑셀 근무자 명단이 모두 삭제됩니다.\n\n계속하시겠습니까?');
                if (!confirmed) {
                    // 되돌리기
                    monthPairSelect.value = state.monthStart;
                    return;
                }
                // 엑셀 데이터 초기화
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
            // 화면이 즉시 바뀌거나 초기화되지 않도록 renderAll() 제거
            // 오직 [공무직 달력 확인하기] 버튼을 눌렀을 때만 결과가 반영되도록 통일
        });

        resetBtn.addEventListener('click', () => {
            // 개인 휴무 설정 초기화 (로컬에는 원래 저장하지 않음)
            state.personalSettings = {};
            updatePersonalSettingsUI();
            
            // 모드를 기준 근무일로 변경
            state.mode = 'regular';
            renderAll();
        });

        // 엑셀 패널 토글
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

        // 이름 입력 후 엔터키로 명단에 추가
        workerNameInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                addExcelBtn.click();
            }
        });

        // 엑셀 명단에 추가
        addExcelBtn.addEventListener('click', () => {
            // 기준 근무일 모드에서는 입력 차단
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

            // 현재 DOM에 있는 설정값을 state에 반영 (Run을 누르지 않아도 반영되게)
            if (!state.libraryClosed || state.libraryClosed === '') {
                alert('공무직 도서관 휴관일을 먼저 선택해주세요.');
                libraryClosedSelect.focus();
                return;
            }
            
            const rows = document.querySelectorAll('.personal-month-row');
            rows.forEach(row => {
                const key = row.dataset.key;
                const weekday = parseInt(row.querySelector('.weekday-sel').value);
                const weekend = parseInt(row.querySelector('.weekend-sel').value);
                state.personalSettings[key] = { weekday, weekend };
            });

            // 인원 추가
            excelWorkers.push(name);
            updateExcelWorkersUI();
            workerNameInput.value = '';

            // 1, 2월 달력의 모든 날짜를 순회하며 이 근무자의 근무일을 기록
            for (let calIndex = 1; calIndex <= 2; calIndex++) {
                const month = calIndex === 1 ? state.monthStart : state.monthStart + 1;
                const year = state.year;
                const lastDay = new Date(year, month, 0).getDate();

                for (let d = 1; d <= lastDay; d++) {
                    const date = new Date(year, month - 1, d);
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
                        // 근무일인 경우에만 이름 추가
                        excelDataByDate[dateStr].workers.push(name);
                    }
                }
            }
        });

        // 엑셀 다운로드
        downloadExcelBtn.addEventListener('click', async () => {
            if (excelWorkers.length === 0) {
                alert('엑셀에 추가된 인원이 없습니다. 먼저 명단에 추가해주세요.');
                return;
            }

            downloadExcelBtn.innerHTML = '<span>⏳ 파일 생성 중...</span>';
            downloadExcelBtn.disabled = true;

            try {
                // Base64로 인코딩된 템플릿 파일 데이터를 ArrayBuffer로 변환 (CORS 에러 원천 차단)
                function base64ToArrayBuffer(base64) {
                    const binary_string = window.atob(base64);
                    const len = binary_string.length;
                    const bytes = new Uint8Array(len);
                    for (let i = 0; i < len; i++) {
                        bytes[i] = binary_string.charCodeAt(i);
                    }
                    return bytes.buffer;
                }

                if (typeof templateBase64 === 'undefined') {
                    throw new Error('템플릿 데이터를 찾을 수 없습니다. (template.js 파일 로드 실패)');
                }
                const arrayBuffer = base64ToArrayBuffer(templateBase64);

                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(arrayBuffer);

                // 각 근무자별 총 근무일수 계산용
                const workerTotalDays = {};
                excelWorkers.forEach(name => { workerTotalDays[name] = 0; });

                // 시트별 lastUsedRow 저장 (크로스시트 수식 참조용)
                const lastUsedRowPerSheet = [];
                const sheetNames = [];

                // 기준 근무일 계산 (정직원 평일 기준)
                let regularWorkDays = 0;
                for (let ci = 0; ci < 2; ci++) {
                    const m = ci === 0 ? state.monthStart : state.monthStart + 1;
                    const ld = new Date(state.year, m, 0).getDate();
                    for (let d = 1; d <= ld; d++) {
                        const dt = new Date(state.year, m - 1, d);
                        const dow = dt.getDay();
                        if (dow === 0 || dow === 6) continue;
                        const ds = formatDate(dt);
                        const isHol = holidays.some(h => h.date === ds);
                        if (!isHol) regularWorkDays++;
                    }
                }

                // 시트 2개 처리 (1번 시트, 2번 시트)
                for (let i = 0; i < 2; i++) {
                    const sheet = workbook.worksheets[i];
                    if (!sheet) continue;

                    const year = state.year;
                    const month = i === 0 ? state.monthStart : state.monthStart + 1;
                    const lastDay = new Date(year, month, 0).getDate();

                    // 시트 이름 변경
                    sheet.name = `${year.toString().slice(-2)}년 ${month}월`;
                    sheetNames.push(sheet.name);

                    // A1 셀 제목 변경
                    sheet.getCell('A1').value = `${year}년 ${String(month).padStart(2, '0')}월 노을빛도서관 야간근무 편성표`;

                    // 기존 데이터 모두 지우기 (6행~17행, C~I열)
                    for (let r = 6; r <= 17; r++) {
                        for (let c = 3; c <= 9; c++) {
                            const cell = sheet.getCell(r, c);
                            cell.value = null;
                        }
                    }

                    // 해당 월의 날짜를 동적으로 계산하여 셀에 배치
                    let currentRow = 6;
                    let maxRowUsed = 6; // 실제 사용된 마지막 일자 행
                    for (let d = 1; d <= lastDay; d++) {
                        const date = new Date(year, month - 1, d);
                        const dayOfWeek = date.getDay(); // 0:일, 1:월, ... 6:토
                        const colIdx = (dayOfWeek === 0) ? 9 : dayOfWeek + 2; // 월:3 ~ 일:9

                        // 날짜 입력
                        const dateCell = sheet.getCell(currentRow, colIdx);
                        dateCell.value = new Date(Date.UTC(year, month - 1, d));

                        // 근무자 입력 (바로 아래 셀)
                        const workerCell = sheet.getCell(currentRow + 1, colIdx);
                        const dateStr = formatDate(date);
                        const data = excelDataByDate[dateStr];

                        if (data) {
                            const existingBorder = workerCell.border;
                            if (data.holidayName) {
                                workerCell.value = data.holidayName;
                                workerCell.style = {
                                    font: { color: { argb: 'FFFF0000' }, bold: true, size: 10 },
                                    alignment: { vertical: 'middle', horizontal: 'center' },
                                    border: existingBorder
                                };
                            } else if (data.isClosed) {
                                workerCell.value = '휴관일';
                                workerCell.style = {
                                    font: { color: { argb: 'FFFF0000' }, bold: true, size: 10 },
                                    alignment: { vertical: 'middle', horizontal: 'center' },
                                    border: existingBorder
                                };
                            } else {
                                const workersText = formatWorkersForExcel(data.workers);
                                workerCell.value = workersText;
                                workerCell.style = {
                                    font: { color: { argb: 'FF000000' }, size: 10 },
                                    alignment: { wrapText: true, vertical: 'middle', horizontal: 'center' },
                                    border: existingBorder
                                };
                            }
                        }

                        maxRowUsed = currentRow;

                        // 일요일(9열)을 채웠으면 다음 주로 이동 (2줄 아래로)
                        if (colIdx === 9) {
                            currentRow += 2;
                        }
                    }

                    // 근무자 행 높이 자동 조절 (줄바꿈 수에 따라)
                    const baseRowHeight = 18; // 1줄일 때 기본 높이
                    const lineHeight = 15;    // 추가 줄당 높이
                    for (let wr = 7; wr <= maxRowUsed + 1; wr += 2) {
                        let maxLines = 1;
                        for (let c = 3; c <= 9; c++) {
                            const val = sheet.getCell(wr, c).value;
                            if (val && typeof val === 'string') {
                                const lines = val.split('\n').length;
                                if (lines > maxLines) maxLines = lines;
                            }
                        }
                        if (maxLines > 1) {
                            sheet.getRow(wr).height = baseRowHeight + (maxLines - 1) * lineHeight;
                        }
                    }

                    // 미사용 주차 행 삭제 (템플릿은 6주 = 6행~17행, 사용된 마지막 일자행 이후의 행 삭제)
                    // maxRowUsed: 마지막으로 사용된 일자행, maxRowUsed+1: 근무자행
                    const lastUsedRow = maxRowUsed + 1; // 근무자 행까지 포함
                    const templateLastRow = 17; // 템플릿 6주차 근무자 행 (16:일자, 17:근무자)
                    lastUsedRowPerSheet.push(lastUsedRow); // 크로스시트 참조용 저장

                    if (lastUsedRow < templateLastRow) {
                        // 미사용 행의 내용과 병합 영역 정리
                        for (let r = lastUsedRow + 1; r <= templateLastRow; r++) {
                            for (let c = 1; c <= 9; c++) {
                                sheet.getCell(r, c).value = '';
                            }
                        }
                        // 행 자체를 삭제 (아래에서 위로 삭제해야 인덱스가 안 꼬임)
                        const rowsToDelete = templateLastRow - lastUsedRow;
                        for (let r = 0; r < rowsToDelete; r++) {
                            sheet.spliceRows(lastUsedRow + 1, 1);
                        }
                    }
                    // ===== 해당 월의 기준 근무일 수 계산 (하드코딩 - 셀 수 없으므로) =====
                    let monthRegularDays = 0;
                    for (let d = 1; d <= lastDay; d++) {
                        const dt = new Date(year, month - 1, d);
                        const dow = dt.getDay();
                        const ds = formatDate(dt);
                        const isHol = holidays.some(h => h.date === ds);
                        if (dow !== 0 && dow !== 6 && !isHol) monthRegularDays++;
                    }

                    // ===== 해당 월 근무일 합계를 달력 아래 A열에 작성 (SUMPRODUCT 수식 사용) =====
                    // 근무자 셀 범위: C7:I{lastUsedRow} (worker rows)
                    const workerRange = `C7:I${lastUsedRow}`;

                    // 요약 시작 행: 현재 시트의 마지막 행 + 4 (3칸 띄우기)
                    const monthlySummaryStart = sheet.rowCount + 4;
                    let msr = monthlySummaryStart;

                    // 헤더
                    sheet.getCell(msr, 1).value = `${month}월 근무일수`;
                    sheet.getCell(msr, 1).font = { bold: true, size: 11 };
                    msr++;

                    // 각 근무자: SUMPRODUCT로 이름 등장 횟수 계산 (셀 내 여러 명 포함 대응)
                    excelWorkers.forEach(n => {
                        sheet.getCell(msr, 1).value = n;
                        sheet.getCell(msr, 1).font = { bold: true, size: 10 };
                        // 이름이 셀에 포함된 횟수를 SUMPRODUCT로 계산
                        const formula = `=SUMPRODUCT((LEN(${workerRange})-LEN(SUBSTITUTE(${workerRange},"${n}","")))/LEN("${n}"))&"일"`;
                        sheet.getCell(msr, 2).value = { formula };
                        sheet.getCell(msr, 2).font = { size: 10 };
                        msr++;
                    });

                    // 기준 근무일 (하드코딩 - 계산 불가이므로)
                    msr++;
                    sheet.getCell(msr, 1).value = '기준 근무일';
                    sheet.getCell(msr, 1).font = { bold: true, size: 10 };
                    sheet.getCell(msr, 2).value = `${monthRegularDays}일`;
                    sheet.getCell(msr, 2).font = { bold: true, size: 10 };

                    // ===== 마지막 시트(i===1)에 H열로 2개월 합산 총계 추가 (크로스시트 SUMPRODUCT) =====
                    if (i === 1) {
                        const sheet0Name = sheetNames[0];
                        const lastRow0 = lastUsedRowPerSheet[0];
                        const workerRange0 = `'${sheet0Name}'!C7:I${lastRow0}`;
                        const workerRange1 = `C7:I${lastUsedRow}`;

                        const totalStartRow = monthlySummaryStart;
                        let tsr = totalStartRow;

                        // 헤더
                        sheet.getCell(tsr, 8).value = '총 근무일수 합계';
                        sheet.getCell(tsr, 8).font = { bold: true, size: 11 };
                        tsr++;

                        // 각 근무자: 두 시트 합산 SUMPRODUCT
                        excelWorkers.forEach(n => {
                            sheet.getCell(tsr, 8).value = n;
                            sheet.getCell(tsr, 8).font = { bold: true, size: 10 };
                            const totalFormula =
                                `=SUMPRODUCT((LEN(${workerRange0})-LEN(SUBSTITUTE(${workerRange0},"${n}","")))/LEN("${n}"))` +
                                `+SUMPRODUCT((LEN(${workerRange1})-LEN(SUBSTITUTE(${workerRange1},"${n}","")))/LEN("${n}"))&"일"`;
                            sheet.getCell(tsr, 9).value = { formula: totalFormula };
                            sheet.getCell(tsr, 9).font = { size: 10 };
                            tsr++;
                        });

                        // 기준 근무일 합산 (하드코딩)
                        tsr++;
                        sheet.getCell(tsr, 8).value = '기준 근무일 (2개월)';
                        sheet.getCell(tsr, 8).font = { bold: true, size: 10 };
                        sheet.getCell(tsr, 9).value = `${regularWorkDays}일`;
                        sheet.getCell(tsr, 9).font = { bold: true, size: 10 };
                    }
                }

                // 파일 다운로드
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${state.year.toString().slice(-2)}년 ${state.monthStart}월,${state.monthStart+1}월 근무 편성표.xlsx`;
                a.click();
                window.URL.revokeObjectURL(url);

            } catch (error) {
                console.error(error);
                alert('엑셀 파일을 생성하는 중 오류가 발생했습니다.');
            } finally {
                downloadExcelBtn.innerHTML = '<span>📥 최종 엑셀 파일 다운로드</span>';
                downloadExcelBtn.disabled = false;
            }
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

    function updateExcelWorkersUI() {
        excelCountEl.innerText = excelWorkers.length;
        excelWorkersList.innerHTML = '';
        excelWorkers.forEach((name, idx) => {
            const span = document.createElement('span');
            span.className = 'bg-teal-700 text-teal-50 px-2 py-0.5 rounded border border-teal-500 text-[11px] font-bold flex items-center gap-1 shadow-sm';
            span.innerHTML = `${name} <button onclick="window.removeExcelWorker('${name}')" class="text-teal-300 hover:text-white transition-colors">&times;</button>`;
            excelWorkersList.appendChild(span);
        });
    }

    window.removeExcelWorker = function(nameToRemove) {
        excelWorkers = excelWorkers.filter(n => n !== nameToRemove);
        
        // 엑셀 데이터에서도 해당 근무자 이름 삭제
        for (const dateStr in excelDataByDate) {
            excelDataByDate[dateStr].workers = excelDataByDate[dateStr].workers.filter(w => w !== nameToRemove);
        }
        
        updateExcelWorkersUI();
    };

    function formatWorkersForExcel(workers) {
        if (!workers || workers.length === 0) return '';
        let result = [];
        for (let i = 0; i < workers.length; i += 2) {
            if (i + 1 < workers.length) {
                result.push(`${workers[i]},${workers[i+1]}`);
            } else {
                result.push(workers[i]);
            }
        }
        return result.join('\n');
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
                // 모바일 환경에서 긴 단어가 셀을 넓히는 것을 방지하기 위해 break-words 사용
                html += `<div class="text-[0.6rem] sm:text-[0.65rem] leading-tight mt-1 opacity-85 text-center break-words w-full flex-1 px-0.5">${info.reason}</div>`;
            }
            div.innerHTML = html;
            container.appendChild(div);
        }
        
        return { publicDays, regularDays };
    }
});
