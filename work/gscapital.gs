/** @OnlyCurrentDoc */
/**
 * Capital module for the financial dashboard.
 * Reads only the named range "rangeequity" from the bound spreadsheet.
 */

var CAPITAL_RANGE_NAME = 'rangeequity';
var CAPITAL_REQUIRED_COLUMNS = ['Дата', 'Раздел 1', 'Раздел 2', 'Раздел 3', 'Раздел 4', 'Сумма'];

/**
 * Public endpoint used by capital.html through apiCapital.
 * @param {Object=} params Requested startDate and endDate.
 * @return {Object} Standard {status, data, error} response.
 */
function capitalGetCapitalData(params) {
  try {
    var capitalRequest = capitalNormalizeRequest_(params);
    var capitalSourceRows = capitalReadRows_();
    var capitalDates = capitalUniqueDates_(capitalSourceRows);
    if (!capitalDates.length) {
      return capitalSuccess_({
        kpis: capitalEmptyKpis_(), charts: capitalEmptyCharts_(), tables: {}, filters: {},
        meta: { rangeName: CAPITAL_RANGE_NAME, rows: 0, message: 'В rangeequity нет строк с корректной датой.' }
      });
    }

    var capitalPeriod = capitalResolvePeriod_(capitalRequest, capitalDates);
    var capitalPeriodRows = capitalFilterPeriod_(capitalSourceRows, capitalPeriod.startDate, capitalPeriod.endDate);
    var capitalSnapshotDate = capitalLatestDate_(capitalPeriodRows);
    var capitalSnapshotRows = capitalSnapshotDate ? capitalRowsForDate_(capitalPeriodRows, capitalSnapshotDate) : [];
    var capitalPreviousSnapshotDate = capitalSnapshotDate ? capitalPreviousDate_(capitalDates, capitalSnapshotDate) : null;
    var capitalPreviousSnapshotRows = capitalPreviousSnapshotDate ? capitalRowsForDate_(capitalSourceRows, capitalPreviousSnapshotDate) : [];
    var capitalSeriesDates = capitalUniqueDates_(capitalPeriodRows);
    var capitalKpis = capitalBuildKpis_(capitalSnapshotRows);
    capitalKpis.trends = capitalBuildKpiTrends_(capitalKpis, capitalPreviousSnapshotRows, capitalPreviousSnapshotDate);

    capitalLogInfo_('[Капитал] Срез: ' + (capitalSnapshotDate ? capitalFormatDate_(capitalSnapshotDate) : 'не найден') +
        '; предыдущий срез: ' + (capitalPreviousSnapshotDate ? capitalFormatDate_(capitalPreviousSnapshotDate) : 'не найден') +
        '; ликвидность (краткосрочные обязательства / оборотные активы): ' + capitalKpis.currentRatio);

    return capitalSuccess_({
      kpis: capitalKpis,
      charts: {
        assetsStructure: capitalBuildStructure_(capitalSnapshotRows, 'Активы'),
        liabilitiesStructure: capitalBuildStructure_(capitalSnapshotRows, 'Пассивы'),
        receivables: capitalBuildSectionSeries_(capitalPeriodRows, capitalSeriesDates, 'Раздел 3', 'Дебиторская задолженность'),
        payables: capitalBuildSectionSeries_(capitalPeriodRows, capitalSeriesDates, 'Раздел 3', 'Кредиторская задолженность'),
        capitalMix: capitalBuildCapitalMix_(capitalPeriodRows, capitalSeriesDates),
        currentRatio: capitalBuildRatioSeries_(capitalPeriodRows, capitalSeriesDates, 'currentRatio'),
        independenceRatio: capitalBuildRatioSeries_(capitalPeriodRows, capitalSeriesDates, 'independenceRatio')
      },
      tables: {}, filters: {},
      meta: {
        rangeName: CAPITAL_RANGE_NAME,
        rows: capitalPeriodRows.length,
        period: { startDate: capitalFormatDate_(capitalPeriod.startDate), endDate: capitalFormatDate_(capitalPeriod.endDate) },
        snapshotDate: capitalSnapshotDate ? capitalFormatDate_(capitalSnapshotDate) : null,
        availablePeriod: { startDate: capitalFormatDate_(capitalDates[0]), endDate: capitalFormatDate_(capitalDates[capitalDates.length - 1]) }
      }
    });
  } catch (capitalError) {
    capitalLogError_('capitalGetCapitalData: ' + (capitalError.message || capitalError));
    return capitalError_(capitalError.message || 'Не удалось подготовить данные по капиталу.');
  }
}

/** @param {Object=} params Raw request. @return {Object} Normalized request. @private */
function capitalNormalizeRequest_(params) {
  var capitalParams = params && typeof params === 'object' ? params : {};
  return { startDate: capitalParseDate_(capitalParams.startDate), endDate: capitalParseDate_(capitalParams.endDate) };
}

/** @return {Array<Object>} Source rows. @private */
function capitalReadRows_() {
  var capitalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!capitalSpreadsheet) { throw new Error('Не удалось открыть таблицу, связанную со скриптом.'); }
  var capitalRange = capitalSpreadsheet.getRangeByName(CAPITAL_RANGE_NAME);
  if (!capitalRange) { throw new Error('Не найден именованный диапазон «' + CAPITAL_RANGE_NAME + '».'); }
  var capitalValues = capitalRange.getValues();
  if (!capitalValues.length) { throw new Error('В диапазоне «' + CAPITAL_RANGE_NAME + '» нет строки заголовков.'); }
  var capitalHeaders = capitalValues[0].map(function (capitalHeader) { return String(capitalHeader === null ? '' : capitalHeader).trim(); });
  var capitalIndexes = {};
  CAPITAL_REQUIRED_COLUMNS.forEach(function (capitalColumn) {
    var capitalIndex = capitalHeaders.indexOf(capitalColumn);
    if (capitalIndex === -1) { throw new Error('В rangeequity отсутствует обязательная колонка «' + capitalColumn + '».'); }
    capitalIndexes[capitalColumn] = capitalIndex;
  });
  if (capitalValues.length < 2) { throw new Error('В диапазоне «' + CAPITAL_RANGE_NAME + '» нет строк данных.'); }
  return capitalValues.slice(1).reduce(function (capitalRows, capitalValueRow) {
    if (capitalValueRow.every(function (capitalValue) { return capitalValue === '' || capitalValue === null; })) { return capitalRows; }
    var capitalDate = capitalParseDate_(capitalValueRow[capitalIndexes['Дата']]);
    if (!capitalDate) { return capitalRows; }
    capitalRows.push({
      date: capitalDate,
      level1: capitalCleanText_(capitalValueRow[capitalIndexes['Раздел 1']]),
      level2: capitalCleanText_(capitalValueRow[capitalIndexes['Раздел 2']]),
      level3: capitalCleanText_(capitalValueRow[capitalIndexes['Раздел 3']]),
      level4: capitalCleanText_(capitalValueRow[capitalIndexes['Раздел 4']]),
      amount: capitalNumber_(capitalValueRow[capitalIndexes['Сумма']])
    });
    return capitalRows;
  }, []);
}

/** @param {*} value Cell date. @return {Date|null} Local calendar date. @private */
function capitalParseDate_(value) {
  try {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      var capitalTextDate = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return capitalDateFromParts_(capitalTextDate.substring(0, 4), capitalTextDate.substring(5, 7), capitalTextDate.substring(8, 10));
    }
    if (typeof value !== 'string') { return null; }
    var capitalText = value.trim();
    var capitalMatch = /^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/.exec(capitalText);
    if (capitalMatch) { return capitalDateFromParts_(capitalMatch[1], capitalMatch[2], capitalMatch[3]); }
    capitalMatch = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(capitalText);
    return capitalMatch ? capitalDateFromParts_(capitalMatch[3], capitalMatch[2], capitalMatch[1]) : null;
  } catch (capitalError) { capitalLogError_(capitalError); return null; }
}

/** @param {*} year Year. @param {*} month Month. @param {*} day Day. @return {Date|null} Valid local date. @private */
function capitalDateFromParts_(year, month, day) {
  var capitalDate = new Date(Number(year), Number(month) - 1, Number(day));
  return capitalDate.getFullYear() === Number(year) && capitalDate.getMonth() === Number(month) - 1 && capitalDate.getDate() === Number(day) ? capitalDate : null;
}

/** @param {*} value Numeric cell. @return {number} Number or zero. @private */
function capitalNumber_(value) {
  if (typeof value === 'number' && isFinite(value)) { return value; }
  if (typeof value !== 'string') { return 0; }
  var capitalText = value.replace(/[\s\u00A0]/g, '').replace(',', '.').replace(/[^0-9.\-()]/g, '');
  var capitalNegative = /^\(.*\)$/.test(capitalText);
  capitalText = capitalText.replace(/[()]/g, '');
  var capitalNumber = Number(capitalText);
  return isFinite(capitalNumber) ? (capitalNegative ? -Math.abs(capitalNumber) : capitalNumber) : 0;
}

/** @param {*} value Text value. @return {string} Clean text. @private */
function capitalCleanText_(value) { return String(value === null || value === undefined ? '' : value).replace(/[\u00A0\u200B\uFEFF]/g, ' ').replace(/\s+/g, ' ').trim(); }
/** @param {*} value Text. @return {string} Comparison key. @private */
function capitalKey_(value) { return capitalCleanText_(value).toLowerCase(); }
/** @param {*} left Left. @param {*} right Right. @return {boolean} Equality. @private */
function capitalEquals_(left, right) { return capitalKey_(left) === capitalKey_(right); }

/** @param {Array<Object>} rows Rows. @return {Array<Date>} Unique sorted dates. @private */
function capitalUniqueDates_(rows) {
  var capitalMap = {};
  rows.forEach(function (capitalRow) { capitalMap[capitalRow.date.getTime()] = capitalRow.date; });
  return Object.keys(capitalMap).map(function (capitalKey) { return capitalMap[capitalKey]; }).sort(function (capitalA, capitalB) { return capitalA - capitalB; });
}

/** @param {Array<Date>} dates Sorted dates. @param {Date} date Current date. @return {Date|null} Previous available date. @private */
function capitalPreviousDate_(dates, date) {
  var capitalPrevious = null;
  dates.forEach(function (capitalDate) { if (capitalDate.getTime() < date.getTime()) { capitalPrevious = capitalDate; } });
  return capitalPrevious;
}

/** @param {Object} request Request. @param {Array<Date>} dates Dates. @return {Object} Period. @private */
function capitalResolvePeriod_(request, dates) {
  var capitalStart = request.startDate || new Date(dates[0].getTime());
  var capitalEnd = request.endDate || new Date(dates[dates.length - 1].getTime());
  capitalStart.setHours(0, 0, 0, 0); capitalEnd.setHours(23, 59, 59, 999);
  if (capitalStart.getTime() > capitalEnd.getTime()) { throw new Error('Дата начала периода не может быть больше даты окончания.'); }
  return { startDate: capitalStart, endDate: capitalEnd };
}

/** @param {Array<Object>} rows Rows. @param {Date} start Start. @param {Date} end End. @return {Array<Object>} Rows in range. @private */
function capitalFilterPeriod_(rows, start, end) { return rows.filter(function (capitalRow) { return capitalRow.date >= start && capitalRow.date <= end; }); }
/** @param {Array<Object>} rows Rows. @return {Date|null} Latest date. @private */
function capitalLatestDate_(rows) { var capitalDates = capitalUniqueDates_(rows); return capitalDates.length ? capitalDates[capitalDates.length - 1] : null; }
/** @param {Array<Object>} rows Rows. @param {Date} date Date. @return {Array<Object>} Snapshot rows. @private */
function capitalRowsForDate_(rows, date) { return rows.filter(function (capitalRow) { return capitalRow.date.getTime() === date.getTime(); }); }
/** @param {Array<Object>} rows Rows. @param {string} property Field. @param {string} value Value. @return {number} Sum. @private */
function capitalSumWhere_(rows, property, value) { return rows.filter(function (capitalRow) { return capitalEquals_(capitalRow[property], value); }).reduce(function (capitalSum, capitalRow) { return capitalSum + capitalRow.amount; }, 0); }
/** @param {number} numerator Numerator. @param {number} denominator Denominator. @return {number} Ratio. @private */
function capitalRatio_(numerator, denominator) { return denominator ? numerator / denominator : 0; }

/** @param {Array<Object>} rows Snapshot. @return {Object} KPI values. @private */
function capitalBuildKpis_(rows) {
  var capitalEquity = capitalSumWhere_(rows, 'level2', 'Капитал');
  var capitalCurrentAssets = capitalSumWhere_(rows, 'level2', 'Оборотные активы');
  var capitalShortTermDebt = capitalSumWhere_(rows, 'level3', 'Краткосрочные обязательства');
  var capitalAssets = capitalSumWhere_(rows, 'level1', 'Активы');
  return { equity: capitalEquity, currentRatio: capitalRatio_(capitalShortTermDebt, capitalCurrentAssets), independenceRatio: capitalRatio_(capitalEquity, capitalAssets), currentAssets: capitalCurrentAssets, shortTermDebt: capitalShortTermDebt, assets: capitalAssets };
}

/** @param {Object} current Current KPI values. @param {Array<Object>} previousRows Previous-snapshot rows. @param {Date|null} previousDate Previous snapshot date. @return {Object} KPI changes relative to the previous snapshot. @private */
function capitalBuildKpiTrends_(current, previousRows, previousDate) {
  if (!previousDate) {
    return { equity: capitalTrendEntry_(current.equity, null, null), currentRatio: capitalTrendEntry_(current.currentRatio, null, null), independenceRatio: capitalTrendEntry_(current.independenceRatio, null, null) };
  }
  var capitalPrevious = capitalBuildKpis_(previousRows);
  return {
    equity: capitalTrendEntry_(current.equity, capitalPrevious.equity, capitalFormatDate_(previousDate)),
    currentRatio: capitalTrendEntry_(current.currentRatio, capitalPrevious.currentRatio, capitalFormatDate_(previousDate)),
    independenceRatio: capitalTrendEntry_(current.independenceRatio, capitalPrevious.independenceRatio, capitalFormatDate_(previousDate))
  };
}

/** @param {number} current Current value. @param {number|null} previous Previous value. @param {string|null} previousDate Previous date. @return {Object} Trend payload. @private */
function capitalTrendEntry_(current, previous, previousDate) {
  var capitalHasPrevious = typeof previous === 'number' && isFinite(previous);
  var capitalDirection = !capitalHasPrevious ? null : current > previous ? 'up' : current < previous ? 'down' : 'flat';
  return { current: current, previous: capitalHasPrevious ? previous : null, previousDate: previousDate, delta: capitalHasPrevious ? current - previous : null, relative: capitalHasPrevious && previous !== 0 ? (current - previous) / Math.abs(previous) : null, direction: capitalDirection };
}

/** @param {Array<Object>} rows Snapshot. @param {string} level1 Assets or liabilities. @return {Array<Object>} Pie slices. @private */
function capitalBuildStructure_(rows, level1) {
  var capitalGroups = {};
  rows.filter(function (capitalRow) { return capitalEquals_(capitalRow.level1, level1) && capitalRow.level3; }).forEach(function (capitalRow) { capitalGroups[capitalRow.level3] = (capitalGroups[capitalRow.level3] || 0) + capitalRow.amount; });
  return Object.keys(capitalGroups).map(function (capitalName) { return { label: capitalName, value: capitalGroups[capitalName] }; }).filter(function (capitalItem) { return capitalItem.value !== 0; }).sort(function (capitalA, capitalB) { return Math.abs(capitalB.value) - Math.abs(capitalA.value); });
}

/** @param {Array<Object>} rows Rows. @param {Array<Date>} dates Dates. @param {string} field Field. @param {string} target Target. @return {Array<Object>} Series. @private */
function capitalBuildSectionSeries_(rows, dates, field, target) {
  return dates.map(function (capitalDate) { var capitalDateRows = capitalRowsForDate_(rows, capitalDate); return { label: capitalMonthLabel_(capitalDate), date: capitalFormatDate_(capitalDate), value: capitalSumWhere_(capitalDateRows, field === 'Раздел 3' ? 'level3' : field, target) }; });
}

/** @param {Array<Object>} rows Rows. @param {Array<Date>} dates Dates. @return {Array<Object>} Capital mix. @private */
function capitalBuildCapitalMix_(rows, dates) {
  return dates.map(function (capitalDate) { var capitalDateRows = capitalRowsForDate_(rows, capitalDate), capitalEquity = capitalSumWhere_(capitalDateRows, 'level2', 'Капитал'), capitalDebt = capitalSumWhere_(capitalDateRows, 'level2', 'Обязательства'), capitalTotal = capitalEquity + capitalDebt; return { label: capitalMonthLabel_(capitalDate), date: capitalFormatDate_(capitalDate), equity: capitalEquity, debt: capitalDebt, equityShare: capitalRatio_(capitalEquity, capitalTotal), debtShare: capitalRatio_(capitalDebt, capitalTotal) }; });
}

/** @param {Array<Object>} rows Rows. @param {Array<Date>} dates Dates. @param {string} type Metric. @return {Array<Object>} Ratio series. @private */
function capitalBuildRatioSeries_(rows, dates, type) {
  return dates.map(function (capitalDate) { var capitalKpis = capitalBuildKpis_(capitalRowsForDate_(rows, capitalDate)); return { label: capitalMonthLabel_(capitalDate), date: capitalFormatDate_(capitalDate), value: type === 'currentRatio' ? capitalKpis.currentRatio : capitalKpis.independenceRatio }; });
}

/** @return {Object} Empty KPIs. @private */
function capitalEmptyKpis_() { return { equity: 0, currentRatio: 0, independenceRatio: 0, currentAssets: 0, shortTermDebt: 0, assets: 0 }; }
/** @return {Object} Empty charts. @private */
function capitalEmptyCharts_() { return { assetsStructure: [], liabilitiesStructure: [], receivables: [], payables: [], capitalMix: [], currentRatio: [], independenceRatio: [] }; }
/** @param {Date} date Date. @return {string} DD.MM.YYYY. @private */
function capitalFormatDate_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy'); }
/** @param {Date} date Date. @return {string} Label. @private */
function capitalMonthLabel_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM yyyy'); }
/** @param {*} data Data. @return {Object} Success response. @private */
function capitalSuccess_(data) { var capitalResponse = typeof createSuccessResponse === 'function' ? createSuccessResponse(data) : { status: 'success', data: data, error: null }; return typeof prepareResponse === 'function' ? prepareResponse(capitalResponse) : capitalResponse; }
/** @param {*} message Error. @return {Object} Error response. @private */
function capitalError_(message) { var capitalResponse = typeof createErrorResponse === 'function' ? createErrorResponse(message) : { status: 'error', data: null, error: String(message) }; return typeof prepareResponse === 'function' ? prepareResponse(capitalResponse) : capitalResponse; }
/** @param {*} message Log message. @private */
function capitalLogError_(message) { if (typeof logError === 'function') { logError(message); } else { Logger.log(String(message)); } }
/** @param {*} message Log message. @private */
function capitalLogInfo_(message) { if (typeof logInfo === 'function') { logInfo(message); } else { Logger.log(String(message)); } }
