/* ============================================================
 * Code.gs — скрипт Google Таблицы для сайта «Voda» (версия 1.2)
 *
 * КАК УСТАНОВИТЬ (один раз):
 *  1. Создайте пустую Google Таблицу.
 *  2. Расширения → Apps Script. Удалите всё в файле Code.gs,
 *     вставьте этот код целиком, нажмите «Сохранить».
 *  3. Вверху выберите функцию setup и нажмите «Выполнить».
 *     Google попросит разрешение — разрешите (это ваш же скрипт).
 *     Появятся листы «Таблица», «Настройки», «Журнал».
 *     Внизу всплывёт сообщение с двумя паролями (их же видно и
 *     можно поменять на листе «Настройки», клетки B5 и B6):
 *       пользователь — создаёт и сохраняет расчёты;
 *       администратор — ещё и журнал: смотрит, открывает, удаляет.
 *  4. Начать развёртывание → Новое развёртывание → тип «Веб-приложение».
 *     Выполнять как: «Я». Доступ: «Все».
 *     Скопируйте адрес вида https://script.google.com/macros/s/…/exec
 *     и вставьте его в index.html в строку API_URL.
 *  5. При изменении кода: Начать развёртывание → Управление
 *     развёртываниями → ✎ → Версия: «Новая версия» → Развернуть
 *     (адрес тогда не меняется).
 *
 * В самой таблице тоже всё считается: на листе «Таблица» меняйте
 * жёлтые клетки (общее, итоги колонок, галочки «не распределять») —
 * таблица пересчитается сама. Меню «💧 Вода» — пересчёт, сохранение
 * в журнал, перестройка листа после правки «Настроек».
 * ============================================================ */

var VERSION = '1.2';
var SH_TAB = 'Таблица', SH_SET = 'Настройки', SH_LOG = 'Журнал';
var FIRST_ROW = 9;          // первая строка данных на листе «Таблица»
var C_NO = 1, C_DESC = 2, C_TOT = 3, C_COL1 = 4;   // столбцы A, B, C, D…

var COLOR = {
    head: '#1f3b63', headText: '#ffffff', input: '#fff6cc', auto: '#7a7a7a',
    sum: '#e8eef7', total: '#d6e2f3', err: '#fde2e1', ok: '#e3f4e6', border: '#000000'
};


// ============================================================
// Меню и установка
// ============================================================

function onOpen() {
    SpreadsheetApp.getUi().createMenu('💧 Вода')
        .addItem('Пересчитать таблицу', 'menuRecalc')
        .addItem('Перестроить лист «Таблица» (после правки Настроек)', 'menuRebuild')
        .addSeparator()
        .addItem('Сохранить текущий расчёт в Журнал', 'menuSaveToLog')
        .addItem('Очистить ввод колонок', 'menuClearInputs')
        .addSeparator()
        .addItem('Первичная установка (setup)', 'setup')
        .addToUi();
}

/** Запустите один раз: создаёт листы «Таблица», «Настройки», «Журнал». */
function setup() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var set = ss.getSheetByName(SH_SET);
    var mig = {};
    if (!set) { set = ss.insertSheet(SH_SET); buildSettingsSheet(set); } else mig = migrateSettings(set);
    var log = ss.getSheetByName(SH_LOG);
    if (!log) { log = ss.insertSheet(SH_LOG); buildLogSheet(log); }
    var tab = ss.getSheetByName(SH_TAB);
    var state = null;
    if (tab) { var L = getLayout(tab); if (L) state = readState(tab, L); }
    else tab = ss.insertSheet(SH_TAB, 0);
    if (state && mig.nd5 && state.cols['№5']) state.cols['№5'].nd = false;
    if (!state) state = {
        lines: ['', '', ''], notes: ['', '', ''],
        date: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
        month: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM'),
        total: 30000, cols: {}
    };
    writeTable(tab, readSettings(), state);
    ss.setActiveSheet(tab);
    // лишний пустой «Лист1», если он остался от новой таблицы
    ['Лист1', 'Sheet1'].forEach(function (n) {
        var s = ss.getSheetByName(n);
        if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
    try { onOpen(); } catch (e) { }
    var pw = set.getRange('B5:B6').getValues();
    try {
        ss.toast('Пароль пользователя: ' + pw[0][0] + ', администратора: ' + pw[1][0] +
            '. Поменять — на листе «Настройки».', '💧 Вода: листы созданы', 15);
    } catch (e) { }
}

function buildSettingsSheet(sh) {
    var D = VODA_DEFAULTS;
    sh.clear();
    sh.getRange('A1').setValue('НАСТРОЙКИ').setFontWeight('bold').setFontSize(14);
    sh.getRange('A3:A6').setValues([['Название таблицы'], ['Знаков после запятой (0–3)'], [LBL_USER], [LBL_ADMIN]]);
    sh.getRange('B5:B6').setNumberFormat('@');
    sh.getRange('B3:B6').setValues([[D.title], [D.decimals], [randomPwd()], [randomPwd()]]).setBackground(COLOR.input);
    sh.getRange('C5').setValue('создаёт и сохраняет расчёты').setFontColor(COLOR.auto);
    sh.getRange('C6').setValue('+ журнал: смотрит, открывает, удаляет').setFontColor(COLOR.auto);
    sh.getRange('E3:E4').setValues([['Сумма весов строк:'], ['Сумма долей колонок:']]).setHorizontalAlignment('right');
    sh.getRange('F3').setFormula('=SUM(C9:C300)'); sh.getRange('F4').setFormula('=SUM(F9:F60)');

    sh.getRange('A7').setValue('СТРОКИ (потребители). Вес — доля строки; сумма не обязана быть 100, веса нормируются.').setFontWeight('bold');
    sh.getRange('A8:C8').setValues([['№', 'Описание', 'Вес, %']]);
    var rows = D.rows.map(function (r) { return [r.no, r.desc, r.w]; });
    sh.getRange(9, 1, rows.length, 3).setValues(rows);
    sh.getRange(9, 1, rows.length, 1).setNumberFormat('@').setValues(D.rows.map(function (r) { return [r.no]; }));
    sh.getRange(9, 1, 292, 3).setBackground(COLOR.input);

    sh.getRange('E7').setValue('КОЛОНКИ (в этом порядке на сайте). Доля — как делится остаток общего.').setFontWeight('bold');
    sh.getRange('E8:G8').setValues([['Колонка', 'Доля, %', 'Не распределять по умолчанию']]);
    var cols = D.cols.map(function (c) { return [c.name, c.share, c.nd]; });
    sh.getRange(9, 5, cols.length, 3).setValues(cols);
    sh.getRange(9, 7, 12, 1).insertCheckboxes();
    sh.getRange(9, 7, cols.length, 1).setValues(D.cols.map(function (c) { return [c.nd]; }));
    sh.getRange(9, 5, 12, 3).setBackground(COLOR.input);

    [sh.getRange('A8:C8'), sh.getRange('E8:G8')].forEach(function (r) {
        r.setFontWeight('bold').setBackground(COLOR.head).setFontColor(COLOR.headText).setWrap(true);
    });
    sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 80);
    sh.setColumnWidth(4, 30); sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 80); sh.setColumnWidth(7, 130);
    buildFooterBlock(sh);
    sh.getRange('A11').setNote('Чтобы добавить потребителя — допишите строку в конец списка. ' +
        'Строки читаются до первой полностью пустой. После правки выберите меню «💧 Вода → Перестроить лист».');
    sh.setFrozenRows(8);
}

var LBL_USER = 'Пароль пользователя', LBL_ADMIN = 'Пароль администратора';

function randomPwd() { return String(100000 + Math.floor(Math.random() * 900000)); }

/** Блок «Постоянная надпись» на листе «Настройки»: I7…I13. */
function buildFooterBlock(sh) {
    sh.getRange('I7').setValue('ПОСТОЯННАЯ НАДПИСЬ под таблицей — видна всегда: на сайте, на листе «Таблица» и при печати. До 5 строк, пустые не выводятся.')
        .setFontWeight('bold').setWrap(true);
    sh.getRange('I8').setValue('Текст надписи').setFontWeight('bold').setBackground(COLOR.head).setFontColor(COLOR.headText);
    sh.getRange('I9:I13').setBackground(COLOR.input).setWrap(true).setNumberFormat('@');
    sh.setColumnWidth(8, 30); sh.setColumnWidth(9, 440);
}

/** Перевод «Настроек» прошлых версий. Возвращает, что поменяли. */
function migrateSettings(sh) {
    var out = {};
    // 1.2: постоянная надпись
    if (!String(sh.getRange('I8').getValue()).trim()) buildFooterBlock(sh);
    // 1.2: №5 больше не «не распределять» по умолчанию и получает долю из примеров
    var cv = sh.getRange(9, 5, 52, 3).getValues();
    for (var i = 0; i < cv.length; i++) {
        if (String(cv[i][0]).trim() === '№5' && !(num(cv[i][1]) > 0) && cv[i][2] === true) {
            sh.getRange(9 + i, 6, 1, 2).setValues([[0.1627, false]]);
            out.nd5 = true;
        }
    }
    // 1.1: второй пароль
    if (String(sh.getRange('A6').getValue()).trim()) return out;
    sh.getRange('A5').setValue(LBL_USER);
    sh.getRange('A6').setValue(LBL_ADMIN);
    sh.getRange('B6').setNumberFormat('@').setValue(randomPwd()).setBackground(COLOR.input);
    if (!String(sh.getRange('B5').getValue()).trim()) sh.getRange('B5').setNumberFormat('@').setValue(randomPwd());
    sh.getRange('C5').setValue('создаёт и сохраняет расчёты').setFontColor(COLOR.auto);
    sh.getRange('C6').setValue('+ журнал: смотрит, открывает, удаляет').setFontColor(COLOR.auto);
    return out;
}

function buildLogSheet(sh) {
    sh.clear();
    sh.getRange(1, 1, 1, 9).setValues([['ID', 'Дата расчёта', 'Сохранено', 'Заголовок', 'Общее', 'Колонки', 'Выводы', 'Данные (не править)', 'Ключ']])
        .setFontWeight('bold').setBackground(COLOR.head).setFontColor(COLOR.headText);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(4, 260); sh.setColumnWidth(6, 320); sh.setColumnWidth(7, 260); sh.setColumnWidth(8, 120);
    sh.hideColumns(9);
}

// ============================================================
// Настройки
// ============================================================

function readSettings() {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SET);
    if (!sh) return JSON.parse(JSON.stringify(VODA_DEFAULTS));
    var top = sh.getRange('B3:B6').getValues();
    var s = {
        title: String(top[0][0] || VODA_DEFAULTS.title),
        decimals: Math.max(0, Math.min(3, parseInt(top[1][0], 10) || 0)),
        userPassword: String(top[2][0] || '').trim(),
        adminPassword: String(top[3][0] || '').trim(),
        rows: [], cols: []
    };
    var rv = sh.getRange(9, 1, 292, 3).getValues();
    for (var i = 0; i < rv.length; i++) {
        var no = String(rv[i][0]).trim(), desc = String(rv[i][1]).trim();
        if (!no && !desc) break;
        s.rows.push({ no: no, desc: desc, w: num(rv[i][2]) || 0 });
    }
    var cv = sh.getRange(9, 5, 52, 3).getValues();
    for (i = 0; i < cv.length; i++) {
        var name = String(cv[i][0]).trim();
        if (!name) break;
        s.cols.push({ name: name, share: num(cv[i][1]) || 0, nd: cv[i][2] === true });
    }
    if (!s.cols.length) s.cols = VODA_DEFAULTS.cols;
    s.footer = sh.getRange('I9:I13').getValues().map(function (r) { return String(r[0]).trim(); })
        .filter(function (t) { return t; });
    return s;
}

function publicSettings(s) {
    return { title: s.title, decimals: s.decimals, rows: s.rows, cols: s.cols, footer: s.footer || [] };
}

function num(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return v;
    var t = String(v).replace(/[\s ]/g, '').replace(',', '.').replace('−', '-');
    if (t === '' || isNaN(+t)) return null;
    return +t;
}

// ============================================================
// Лист «Таблица»
// ============================================================

function layoutFor(settings) {
    var n = settings.rows.length, m = settings.cols.length;
    var rSum = FIRST_ROW + n;
    return {
        n: n, m: m, lastCol: C_COL1 + m - 1,
        cols: settings.cols.map(function (c) { return c.name; }),
        rSum: rSum, rIn: rSum + 1, rNd: rSum + 2, rTot: rSum + 3, rMsg: rSum + 4,
        rNotes: rSum + 7, rFoot: rSum + 11, footer: (settings.footer || []).slice(), id: ''
    };
}

function getLayout(sh) {
    try { var t = sh.getRange(FIRST_ROW - 1, 1).getNote(); return t ? JSON.parse(t) : null; }
    catch (e) { return null; }
}

function readState(sh, L) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var lines = sh.getRange(2, 1, 3, 1).getValues().map(function (r) { return String(r[0]); });
    var dv = sh.getRange(5, 2).getValue();
    var date = dv instanceof Date ? Utilities.formatDate(dv, tz, 'yyyy-MM-dd') : String(dv || '');
    var inp = sh.getRange(L.rIn, C_TOT, 2, L.m + 1).getValues();
    var mv = sh.getRange(5, C_TOT + 1).getValue(), month = '';
    if (mv instanceof Date) month = Utilities.formatDate(mv, tz, 'yyyy-MM');
    else { var mm = /(\d{1,2})[.\-\/](\d{4})/.exec(String(mv || '')); if (mm) month = mm[2] + '-' + ('0' + mm[1]).slice(-2); }
    var st = { id: L.id || '', lines: lines, date: date, month: month, total: num(inp[0][0]), cols: {}, notes: [] };
    for (var j = 0; j < L.m; j++) st.cols[L.cols[j]] = { value: num(inp[0][j + 1]), nd: inp[1][j + 1] === true };
    st.notes = sh.getRange(L.rNotes, 1, 3, 1).getValues().map(function (r) { return String(r[0]); });
    return st;
}

/** Полная перерисовка листа: оформление + ввод + расчёт. */
function writeTable(sh, settings, state) {
    var L = layoutFor(settings);
    L.id = state.id || '';
    var all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    all.breakApart(); all.clearDataValidations(); all.clearNote(); sh.clear();
    var W = L.lastCol;
    if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
    if (sh.getMaxRows() < L.rFoot + 6) sh.insertRowsAfter(sh.getMaxRows(), L.rFoot + 6 - sh.getMaxRows());

    sh.setColumnWidth(C_NO, 45); sh.setColumnWidth(C_DESC, 240);
    for (var j = 0; j <= L.m; j++) sh.setColumnWidth(C_TOT + j, 90);   // «Общее» и все колонки одной ширины

    // заголовок, три строки описания, дата
    sh.getRange(1, 1, 1, W).merge().setValue(settings.title).setFontSize(14).setFontWeight('bold');
    var lines = state.lines || ['', '', ''];
    for (var k = 0; k < 3; k++) sh.getRange(2 + k, 1, 1, W).merge().setValue(lines[k] || '').setBackground(COLOR.input).setWrap(true);
    sh.getRange(5, 1).setValue('Дата:').setFontWeight('bold').setHorizontalAlignment('right');
    var dateCell = sh.getRange(5, 2).setBackground(COLOR.input).setNumberFormat('dd.mm.yyyy').setHorizontalAlignment('left');
    if (state.date && /^\d{4}-\d{2}-\d{2}$/.test(state.date)) {
        var p = state.date.split('-'); dateCell.setValue(new Date(+p[0], +p[1] - 1, +p[2]));
    } else dateCell.setValue(state.date || '');
    dateCell.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
    // шильдик «Отчёт за»: дата 1-го числа месяца в формате «ноябрь 2026»
    sh.getRange(5, C_TOT).setValue('Отчёт за:').setFontWeight('bold').setHorizontalAlignment('right');
    var mCell = sh.getRange(5, C_TOT + 1, 1, 2).merge().getCell(1, 1);
    mCell.setBackground(COLOR.input).setNumberFormat('mmmm yyyy').setFontWeight('bold').setHorizontalAlignment('left');
    var mm = /^(\d{4})-(\d{2})$/.exec(state.month || '');
    mCell.setValue(mm ? new Date(+mm[1], +mm[2] - 1, 1) : '');
    mCell.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true)
        .setHelpText('Месяц отчёта: введите, например, 01.11.2026').build());
    sh.getRange(6, 1, 1, W).merge().setValue('Жёлтые клетки — для ввода. Остальное считается само.')
        .setFontColor(COLOR.auto).setFontStyle('italic').setFontSize(9);

    // шапка таблицы
    var head = ['№', 'Описание', 'Общее'].concat(L.cols);
    sh.getRange(FIRST_ROW - 1, 1, 1, W).setValues([head]).setFontWeight('bold')
        .setBackground(COLOR.head).setFontColor(COLOR.headText).setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW - 1, 1).setNote(JSON.stringify(L));

    // № и описание строк
    if (L.n) {
        sh.getRange(FIRST_ROW, 1, L.n, 1).setNumberFormat('@');
        sh.getRange(FIRST_ROW, 1, L.n, 2).setValues(settings.rows.map(function (r) { return [r.no, r.desc]; }));
        sh.getRange(FIRST_ROW, 1, L.n, 1).setHorizontalAlignment('center');
    }
    // служебные строки
    sh.getRange(L.rSum, 2).setValue('Сумма по строкам');
    sh.getRange(L.rIn, 2).setValue('Ввод итога (пусто = авто)');
    sh.getRange(L.rNd, 2).setValue('Не распределять по строкам');
    sh.getRange(L.rTot, 2).setValue('Общее');
    sh.getRange(L.rSum, 1, 1, W).setBackground(COLOR.sum).setFontStyle('italic');
    sh.getRange(L.rIn, C_TOT, 1, L.m + 1).setBackground(COLOR.input);
    sh.getRange(L.rNd, C_COL1, 1, L.m).insertCheckboxes().setBackground(COLOR.input).setHorizontalAlignment('center');
    sh.getRange(L.rTot, 1, 1, W).setBackground(COLOR.total).setFontWeight('bold');

    // ввод
    var cs = state.cols || {};
    var inRow = [state.total === null || state.total === undefined ? '' : state.total];
    var ndRow = [];
    for (j = 0; j < L.m; j++) {
        var c = cs[L.cols[j]] || {};
        inRow.push(c.value === null || c.value === undefined || c.value === '' ? '' : +c.value);
        ndRow.push(c.nd === undefined || c.nd === null ? !!settings.cols[j].nd : !!c.nd);
    }
    sh.getRange(L.rIn, C_TOT, 1, L.m + 1).setValues([inRow]);
    sh.getRange(L.rNd, C_COL1, 1, L.m).setValues([ndRow]);

    // числовой формат и рамки
    var nf = '#,##0' + (settings.decimals ? '.' + new Array(settings.decimals + 1).join('0') : '');
    sh.getRange(FIRST_ROW, C_TOT, L.rTot - FIRST_ROW + 1, L.m + 1).setNumberFormat(nf).setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW - 1, 1, L.rTot - FIRST_ROW + 2, W)
        .setBorder(true, true, true, true, true, true, COLOR.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(FIRST_ROW - 1, 1, L.rTot - FIRST_ROW + 2, W)
        .setBorder(true, true, true, true, null, null, COLOR.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(FIRST_ROW - 1, C_TOT, L.rTot - FIRST_ROW + 2, 1)
        .setBorder(null, true, null, true, null, null, COLOR.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // строка состояния и выводы
    sh.getRange(L.rMsg, 1, 1, W).merge().setWrap(true).setFontSize(9);
    sh.getRange(L.rNotes - 1, 1).setValue('Выводы:').setFontWeight('bold');
    var notes = state.notes || ['', '', ''];
    for (k = 0; k < 3; k++) sh.getRange(L.rNotes + k, 1, 1, W).merge().setValue(notes[k] || '').setBackground(COLOR.input).setWrap(true);
    // постоянная надпись — из «Настроек», здесь не править
    for (k = 0; k < L.footer.length; k++) sh.getRange(L.rFoot + k, 1, 1, W).merge().setValue(L.footer[k]).setWrap(true).setFontColor('#333333');
    if (L.footer.length) sh.getRange(L.rFoot, 1).setNote('Постоянная надпись. Меняется на листе «Настройки», столбец I.');

    sh.setFrozenRows(0);
    renderValues(sh, L, settings, voda_calc(settings, state));
    return L;
}

/** Пишет только посчитанные числа (быстро, для onEdit). */
function renderValues(sh, L, settings, res) {
    var vals = [];
    for (var i = 0; i < L.n; i++) {
        var row = [res.rowTot[i]];
        for (var j = 0; j < L.m; j++) row.push(res.cols[j].nd ? '' : res.cells[i][j]);
        vals.push(row);
    }
    if (L.n) sh.getRange(FIRST_ROW, C_TOT, L.n, L.m + 1).setValues(vals);
    var sumRow = [res.distributed];
    for (j = 0; j < L.m; j++) sumRow.push(res.cols[j].nd ? '' : res.cols[j].cellSum);
    sh.getRange(L.rSum, C_TOT, 1, L.m + 1).setValues([sumRow]);

    var totRow = [res.total], totColors = [res.totalAuto ? COLOR.auto : '#000000'], totStyle = [res.totalAuto ? 'italic' : 'normal'];
    for (j = 0; j < L.m; j++) {
        totRow.push(res.cols[j].value);
        totColors.push(res.cols[j].auto ? COLOR.auto : '#000000');
        totStyle.push(res.cols[j].auto ? 'italic' : 'normal');
    }
    sh.getRange(L.rTot, C_TOT, 1, L.m + 1).setValues([totRow]).setFontColors([totColors]).setFontStyles([totStyle]);

    var msg = res.ok ? '✓ ' + res.info + ' Серым курсивом — посчитано автоматически.' : '⚠ ' + res.errors.join(' ');
    sh.getRange(L.rMsg, 1).setValue(msg).setBackground(res.ok ? COLOR.ok : COLOR.err);
}

function recalcSheet(sh, L) {
    var settings = readSettings();
    var same = L.n === settings.rows.length && L.m === settings.cols.length &&
        L.cols.join('\u0001') === settings.cols.map(function (c) { return c.name; }).join('\u0001') &&
        JSON.stringify(L.footer || []) === JSON.stringify(settings.footer || []);
    var state = readState(sh, L);
    if (!same) return writeTable(sh, settings, state);
    renderValues(sh, L, settings, voda_calc(settings, state));
    return L;
}

/** Простой триггер: правка жёлтых клеток ввода → пересчёт. */
function onEdit(e) {
    try {
        var sh = e.range.getSheet();
        if (sh.getName() === SH_SET) {          // правка настроек — сразу обновляем «Таблицу»
            var t = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAB), TL = t && getLayout(t);
            if (TL) recalcSheet(t, TL);
            return;
        }
        if (sh.getName() !== SH_TAB) return;
        var L = getLayout(sh);
        if (!L) return;
        var r1 = e.range.getRow(), r2 = e.range.getLastRow();
        if (r2 < L.rIn || r1 > L.rNd) return;
        recalcSheet(sh, L);
    } catch (err) {
        try { SpreadsheetApp.getActiveSpreadsheet().toast(String(err), 'Ошибка пересчёта', 8); } catch (x) { }
    }
}

function tabSheet() {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAB);
    if (!sh || !getLayout(sh)) { setup(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TAB); }
    return sh;
}

function menuRecalc() { var sh = tabSheet(); recalcSheet(sh, getLayout(sh)); }
function menuRebuild() { var sh = tabSheet(); writeTable(sh, readSettings(), readState(sh, getLayout(sh))); }
function menuClearInputs() {
    var sh = tabSheet(), L = getLayout(sh);
    sh.getRange(L.rIn, C_COL1, 1, L.m).clearContent();
    recalcSheet(sh, L);
}
function menuSaveToLog() {
    var sh = tabSheet(), L = getLayout(sh);
    var state = readState(sh, L), settings = readSettings();
    var res = voda_calc(settings, state);
    if (!res.ok) { SpreadsheetApp.getUi().alert('Не сохранено: ' + res.errors.join(' ')); return; }
    state.id = logSave(state, settings, res, true).id;
    L.id = state.id; sh.getRange(FIRST_ROW - 1, 1).setNote(JSON.stringify(L));
    SpreadsheetApp.getActiveSpreadsheet().toast('Сохранено в Журнал', '💧 Вода', 4);
}

// ============================================================
// Журнал
// ============================================================

function logSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SH_LOG);
    if (!sh) { sh = ss.insertSheet(SH_LOG); buildLogSheet(sh); }
    return sh;
}

/** Сохраняет расчёт. Обновить запись можно, только если isAdmin или знаешь её ключ. */
function logSave(state, settings, res, isAdmin) {
    var sh = logSheet();
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var id = state.id || ('R' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000));
    var clean = cleanState(state); clean.id = id;
    var token = String(state.token || '');
    var colsText = res.cols.map(function (c) {
        return c.name + ': ' + voda_format(c.value, res.decimals) + (c.nd ? ' (не распределялась)' : '');
    }).join('; ');
    var row = [id, clean.date, Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm'),
        [voda_monthLabel(clean.month)].concat(clean.lines || []).filter(String).join(' / '), res.total, colsText,
        (clean.notes || []).filter(String).join(' / '), JSON.stringify(clean)];
    var last = sh.getLastRow(), at = 0;
    if (last > 1 && state.id) {
        var ids = sh.getRange(2, 1, last - 1, 9).getValues();
        for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) {
            var key = String(ids[i][8] || '');
            if (!isAdmin && (!key || key !== token)) throw new Error('Эту запись журнала может изменить только администратор. Сохраните как новый расчёт.');
            at = i + 2; token = key; break;
        }
    }
    if (!token) token = Utilities.getUuid();
    row.push(token);
    if (!at) { sh.insertRowBefore(2); at = 2; }   // новые — сверху
    var rg = sh.getRange(at, 1, 1, row.length);
    rg.setFontWeight('normal').setBackground(null).setFontColor('#000000');
    sh.getRange(at, 1, 1, 3).setNumberFormat('@');
    sh.getRange(at, 6, 1, 3).setNumberFormat('@');
    rg.setValues([row]);
    return { id: id, token: token };
}

function logList() {
    var sh = logSheet(), last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 8).getValues().map(function (r) {
        var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
        var f = function (v, p) { return v instanceof Date ? Utilities.formatDate(v, tz, p) : String(v); };
        var month = '', lines = [];
        try { var j = JSON.parse(r[7]); month = j.month || ''; lines = j.lines || []; } catch (e) { }
        return { id: String(r[0]), date: f(r[1], 'yyyy-MM-dd'), month: month, saved: f(r[2], 'dd.MM.yyyy HH:mm'),
            title: lines.filter(String).join(' / '), total: r[4] };
    }).filter(function (x) { return x.id; });
}

function logGet(id) {
    var sh = logSheet(), last = sh.getLastRow();
    if (last < 2) return null;
    var v = sh.getRange(2, 1, last - 1, 8).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][0]) === id) {
        try { return JSON.parse(v[i][7]); } catch (e) { return null; }
    }
    return null;
}

function logDelete(id) {
    var sh = logSheet(), last = sh.getLastRow();
    if (last < 2) return false;
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { sh.deleteRow(i + 2); return true; }
    return false;
}

function cleanState(s) {
    s = s || {};
    var cols = {};
    Object.keys(s.cols || {}).forEach(function (k) {
        var c = s.cols[k] || {};
        cols[String(k)] = { value: num(c.value), nd: !!c.nd };
    });
    var three = function (a) { a = (a || []).slice(0, 3).map(function (x) { return String(x || '').slice(0, 500); }); while (a.length < 3) a.push(''); return a; };
    return {
        id: String(s.id || ''), date: String(s.date || '').slice(0, 10),
        month: /^\d{4}-\d{2}$/.test(String(s.month || '')) ? String(s.month) : '',
        lines: three(s.lines), notes: three(s.notes), total: num(s.total), cols: cols
    };
}

// ============================================================
// API для сайта
//   GET  ?action=load|get|delete|ping&password=…&callback=…  (JSONP)
//   POST {action:'save', password, state, journal:true}      (text/plain)
// ============================================================

function doGet(e) {
    var p = (e && e.parameter) || {};
    var out;
    try { out = handle(p.action, p); } catch (err) { out = { success: false, error: String(err && err.message || err) }; }
    return reply(out, p.callback);
}

function doPost(e) {
    var body = {};
    try {
        if (e && e.postData && e.postData.contents) {
            try { body = JSON.parse(e.postData.contents); } catch (x) { body = {}; }
        }
        if (!body.action && e && e.parameter && e.parameter.payload) body = JSON.parse(e.parameter.payload);
        if (!body.action && e && e.parameter && e.parameter.action) body = e.parameter;
    } catch (err) { body = {}; }
    var out;
    try { out = handle(body.action, body); } catch (err) { out = { success: false, error: String(err && err.message || err) }; }
    return reply(out, null);
}

function reply(obj, callback) {
    var json = JSON.stringify(obj);
    if (callback && /^[A-Za-z_$][\w$]{0,80}$/.test(callback)) {
        return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/** Роль по паролю: 'admin', 'user' или '' (не пустили). */
function roleFor(settings, pw) {
    pw = String(pw || '');
    if (!settings.adminPassword && !settings.userPassword) return 'admin';   // пароли не заданы — открыто всем
    if (settings.adminPassword && pw === settings.adminPassword) return 'admin';
    if (settings.userPassword && pw === settings.userPassword) return 'user';
    return '';
}

function handle(action, p) {
    var settings = readSettings();
    if (action === 'ping') return { success: true, version: VERSION };
    var role = roleFor(settings, p.password);
    if (!role) return { success: false, needPassword: true, error: p.password ? 'Неверный пароль' : 'Нужен пароль' };
    var isAdmin = role === 'admin';
    var denied = { success: false, error: 'Это может только администратор' };

    if (action === 'load') {
        var out = { success: true, role: role, settings: publicSettings(settings), journal: [] };
        if (isAdmin) {
            var sh = tabSheet();
            out.current = readState(sh, getLayout(sh));
            out.journal = logList();
        }
        return out;
    }
    if (action === 'journal') return isAdmin ? { success: true, journal: logList() } : denied;
    if (action === 'get') {
        if (!isAdmin) return denied;
        var item = logGet(String(p.id || ''));
        return item ? { success: true, state: item } : { success: false, error: 'Запись не найдена' };
    }
    if (action === 'delete') {
        if (!isAdmin) return denied;
        var lockD = LockService.getDocumentLock(); lockD.waitLock(20000);
        try { return { success: logDelete(String(p.id || '')), journal: logList() }; } finally { lockD.releaseLock(); }
    }
    if (action === 'save') {
        var raw = typeof p.state === 'string' ? JSON.parse(p.state) : (p.state || {});
        var state = cleanState(raw);
        state.token = String(raw.token || '');
        var res = voda_calc(settings, state);
        if (!res.ok) return { success: false, error: 'Не сохранено: ' + res.errors.join(' ') };
        var lock = LockService.getDocumentLock(); lock.waitLock(20000);
        try {
            var saved = logSave(state, settings, res, isAdmin);
            state.id = saved.id;
            writeTable(tabSheet(), settings, state);
            SpreadsheetApp.flush();
            return { success: true, id: saved.id, token: saved.token, journal: isAdmin ? logList() : [] };
        } finally { lock.releaseLock(); }
    }
    return { success: false, error: 'Неизвестное действие: ' + action };
}

// ============================================================
// Общий с сайтом расчёт
// ============================================================

// ==== CALC START ==== (этот блок одинаковый в index.html и Code.gs)
/*
 * Расчёт таблицы.
 *  settings: { rows:[{no,desc,w}], cols:[{name,share,nd}], decimals }
 *  state:    { total:number|null, cols:{ имя:{ value:number|null, nd:bool } } }
 *
 * 1. Колонки: введённые берутся как есть, остаток (общее − введённые)
 *    раскладывается на НЕвведённые по их долям. Если доли у всех
 *    оставшихся нулевые — поровну между ними.
 *    Если общее не введено — общее = сумма введённых колонок.
 * 2. Каждая колонка без отметки «не распределять» раскладывается по
 *    строкам по весам строк. Округление методом наибольшего остатка,
 *    поэтому сумма клеток колонки ВСЕГДА точно равна итогу колонки.
 * 3. Общее по строке = сумма её клеток. «Не распределять» — колонка
 *    есть только в итоговой строке, в строки не попадает.
 */
function voda_apportion(total, weights) {
    // total — целое (в минимальных единицах), weights — неотрицательные
    var n = weights.length, res = [], i, sumW = 0;
    for (i = 0; i < n; i++) { res.push(0); sumW += Math.max(0, +weights[i] || 0); }
    if (n === 0 || total === 0) return res;
    if (sumW <= 0) return null;
    var sign = total < 0 ? -1 : 1, abs = Math.abs(total), used = 0, rest = [];
    for (i = 0; i < n; i++) {
        var exact = abs * Math.max(0, +weights[i] || 0) / sumW;
        var fl = Math.floor(exact + 1e-9);
        res[i] = fl; used += fl;
        rest.push({ i: i, r: exact - fl, w: weights[i] });
    }
    rest.sort(function (a, b) { return (b.r - a.r) || (b.w - a.w) || (a.i - b.i); });
    for (var k = 0; k < abs - used; k++) res[rest[k % n].i] += 1;
    for (i = 0; i < n; i++) res[i] *= sign;
    return res;
}

function voda_calc(settings, state) {
    var d = Math.max(0, Math.min(3, parseInt(settings.decimals, 10) || 0));
    var unit = Math.pow(10, d);
    var toU = function (x) { return Math.round(x * unit); };
    var fromU = function (x) { return x === null ? null : x / unit; };
    var fmt = function (x) { return voda_format(x / unit, d); };
    var cols = settings.cols, rows = settings.rows;
    var st = state || {}, stc = st.cols || {};
    var errors = [], warnings = [];

    var entered = [], nd = [], sumE = 0, j, i;
    for (j = 0; j < cols.length; j++) {
        var c = stc[cols[j].name] || {};
        var v = (c.value === null || c.value === undefined || c.value === '' || isNaN(+c.value)) ? null : toU(+c.value);
        if (v !== null && v < 0) { errors.push('Колонка ' + cols[j].name + ': число не может быть отрицательным'); v = null; }
        entered.push(v);
        nd.push(c.nd === undefined || c.nd === null ? !!cols[j].nd : !!c.nd);
        if (v !== null) sumE += v;
    }

    var totalEntered = !(st.total === null || st.total === undefined || st.total === '' || isNaN(+st.total));
    var T = totalEntered ? toU(+st.total) : sumE;
    if (T < 0) { errors.push('Общее не может быть отрицательным'); T = 0; }

    var free = [], freeW = [];
    for (j = 0; j < cols.length; j++) if (entered[j] === null) { free.push(j); freeW.push(+cols[j].share || 0); }
    var R = T - sumE;
    var C = entered.slice(), auto = [];
    for (j = 0; j < cols.length; j++) auto.push(entered[j] === null);
    var maxFor = {}; // сколько ещё можно ввести в каждую колонку

    if (totalEntered) {
        for (j = 0; j < cols.length; j++) maxFor[cols[j].name] = fromU(T - sumE + (entered[j] || 0));
    }

    if (R < 0) {
        errors.push('Введённые колонки в сумме ' + fmt(sumE) + ' — это больше общего (' + fmt(T) + ') на ' + fmt(-R) + '. Уменьшите колонки или общее.');
        for (var f = 0; f < free.length; f++) C[free[f]] = 0;
    } else if (free.length === 0) {
        if (R > 0) errors.push('Все колонки введены, в сумме ' + fmt(sumE) + ', а общее ' + fmt(T) + ': не хватает ' + fmt(R) + '.');
    } else {
        var allZero = true;
        for (var z = 0; z < freeW.length; z++) if (freeW[z] > 0) allZero = false;
        if (allZero) for (z = 0; z < freeW.length; z++) freeW[z] = 1;
        var part = voda_apportion(R, freeW);
        for (f = 0; f < free.length; f++) C[free[f]] = part[f];
    }

    var weights = [];
    for (i = 0; i < rows.length; i++) weights.push(Math.max(0, +rows[i].w || 0));
    var cells = [], rowTot = [], colCellSum = [], D = 0, ndSum = 0;
    for (i = 0; i < rows.length; i++) { cells.push([]); rowTot.push(0); }
    for (j = 0; j < cols.length; j++) {
        if (nd[j]) {
            for (i = 0; i < rows.length; i++) cells[i].push(null);
            colCellSum.push(null); ndSum += C[j];
            continue;
        }
        var col = voda_apportion(C[j], weights);
        if (col === null) {
            if (C[j] !== 0) errors.push('У всех строк вес 0 — колонку ' + cols[j].name + ' не по чему распределять (заполните веса на листе «Настройки»)');
            col = []; for (i = 0; i < rows.length; i++) col.push(0);
        }
        var s = 0;
        for (i = 0; i < rows.length; i++) { cells[i].push(col[i]); rowTot[i] += col[i]; s += col[i]; }
        colCellSum.push(s); D += s;
    }

    var info = '';
    if (!errors.length) {
        var autoNames = [];
        for (f = 0; f < free.length; f++) autoNames.push(cols[free[f]].name);
        if (free.length && totalEntered && free.length < cols.length) {
            info = 'Введено колонок: ' + (cols.length - free.length) + ' из ' + cols.length + '. Остаток ' + fmt(R) +
                ' разложен на ' + autoNames.join(', ') + ' по их долям.';
        } else if (free.length === cols.length && totalEntered) {
            info = 'Общее разложено на все колонки по долям. Можно вводить итоги колонок.';
        } else if (!totalEntered) {
            info = 'Общее не введено — считается как сумма колонок.';
        } else {
            info = 'Все колонки введены, сумма сходится с общим.';
        }
        if (ndSum) info += ' Не распределяется по строкам: ' + fmt(ndSum) + '.';
    }

    var out = {
        ok: errors.length === 0, errors: errors, warnings: warnings, info: info,
        decimals: d,
        total: fromU(T), totalAuto: !totalEntered, rest: fromU(R),
        cols: [], cells: [], rowTot: [], distributed: fromU(D), notDistributed: fromU(ndSum)
    };
    for (j = 0; j < cols.length; j++) out.cols.push({
        name: cols[j].name, value: fromU(C[j]), auto: auto[j], nd: nd[j],
        cellSum: colCellSum[j] === null ? null : fromU(colCellSum[j]),
        max: maxFor[cols[j].name] === undefined ? null : maxFor[cols[j].name]
    });
    for (i = 0; i < rows.length; i++) {
        var r = [];
        for (j = 0; j < cols.length; j++) r.push(cells[i][j] === null ? null : fromU(cells[i][j]));
        out.cells.push(r); out.rowTot.push(fromU(rowTot[i]));
    }
    return out;
}

function voda_format(x, d) {
    if (x === null || x === undefined || x === '') return '';
    var s = (+x).toFixed(d || 0), neg = s.charAt(0) === '-';
    if (neg) s = s.slice(1);
    var p = s.split('.'), int = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (neg ? '−' : '') + int + (p[1] ? ',' + p[1] : '');
}

var VODA_MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
/** '2026-11' → 'Ноябрь 2026' */
function voda_monthLabel(ym) {
    var m = /^(\d{4})-(\d{2})$/.exec(String(ym || ''));
    return m ? VODA_MONTHS[+m[2] - 1] + ' ' + m[1] : '';
}

// Значения по умолчанию — взяты из ваших трёх примеров (30 000, 35 000, 10 000)
var VODA_DEFAULTS = {
    title: 'Распределение по потребителям',
    decimals: 0,
    footer: [],
    rows: [
        { no: '1', desc: '', w: 22.6101 }, { no: '2', desc: '', w: 0 }, { no: '3', desc: '', w: 0 },
        { no: '4', desc: '', w: 2.2597 }, { no: '5', desc: '', w: 40.2588 }, { no: '6', desc: '', w: 3.5124 },
        { no: '7', desc: '', w: 3.5565 }, { no: '8', desc: '', w: 1.5385 }, { no: '9', desc: '', w: 1.0030 },
        { no: '10', desc: '', w: 0.3138 }, { no: '11', desc: '', w: 0.1082 }, { no: '12', desc: '', w: 0.0975 },
        { no: '14', desc: '', w: 3.9985 }, { no: '15', desc: '', w: 2.4373 }, { no: '16', desc: '', w: 0.3779 },
        { no: '17', desc: '', w: 0.2631 }, { no: '18', desc: '', w: 14.4168 }, { no: '19', desc: '', w: 0.0975 },
        { no: '20', desc: '', w: 0.0828 }, { no: '21', desc: '', w: 0.0280 }, { no: '22', desc: '', w: 0.0374 },
        { no: '23', desc: '', w: 1.5719 }, { no: '24', desc: '', w: 1.3929 }, { no: '25', desc: '', w: 0.0374 }
    ],
    cols: [
        { name: '№1', share: 17.8533, nd: false }, { name: '№2', share: 17.2747, nd: false },
        { name: '№4', share: 19.9000, nd: false }, { name: '№3', share: 24.7653, nd: false },
        { name: '№6', share: 20.0440, nd: false }, { name: '№5', share: 0.1627, nd: false }
    ]
};
// ==== CALC END ====
