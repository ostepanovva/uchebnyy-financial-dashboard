/** @OnlyCurrentDoc */

/**
 * Money module. Reads the named range "rangecash" and returns a single
 * serializable payload for money.html. All module functions start with cash.
 */

var CASH_RANGE_NAME = 'rangecash';
var CASH_REQUIRED_COLUMNS = ['Дата', 'Сумма', 'Вид операции', 'Счет', 'Контрагент', 'Статья', 'Раздел ОДДС', 'Направление'];
var CASH_MONTH_NAMES = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

/**
 * Public endpoint used by money.html through google.script.run.
 * @param {Object=} params Page filters and selected period.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 */
function cashGetMoneyData(params) {
  try {
    var cashRequest = cashNormalizeRequest_(params);
    var cashSource = cashReadNamedRange_();
    var cashValidRows = cashSource.rows.filter(function (cashRow) { return cashRow.date !== null; });
    var cashCategorizedRows = cashValidRows.filter(function (cashRow) {
      return cashMatchesFilters_(cashRow, cashRequest);
    });
    var cashPeriodRows = cashCategorizedRows.filter(function (cashRow) {
      return cashIsInPeriod_(cashRow.date, cashRequest.startDate, cashRequest.endDate);
    });
    var cashAnalyticsRows = cashPeriodRows.filter(function (cashRow) { return !cashIsMovement_(cashRow); });
    var cashHistoryRows = cashCategorizedRows.filter(function (cashRow) { return !cashIsMovement_(cashRow); });
    var cashRange = cashResolveVisibleRange_(cashAnalyticsRows, cashRequest);
    var cashMonths = cashBuildMonths_(cashRange.startDate, cashRange.endDate, cashAnalyticsRows);
    var cashBalances = cashBuildBalances_(cashHistoryRows, cashMonths);

    return cashSuccessResponse_({
      kpis: cashBuildKpis_(cashAnalyticsRows, cashRange),
      charts: {
        counterparties: {
          suppliers: cashBuildCounterparties_(cashAnalyticsRows, 'Расход'),
          buyers: cashBuildCounterparties_(cashAnalyticsRows, 'Приход')
        },
        directions: cashBuildDirections_(cashAnalyticsRows),
        monthlyNetFlow: cashBuildMonthlyNetFlow_(cashAnalyticsRows, cashMonths),
        balances: cashBalances,
        expenseStructure: cashBuildExpenseStructure_(cashAnalyticsRows)
      },
      report: cashBuildReport_(cashAnalyticsRows, cashMonths, cashBalances),
      filters: cashBuildFilterOptions_(cashValidRows),
      meta: {
        rangeName: CASH_RANGE_NAME,
        totalSourceRows: cashSource.rows.length,
        invalidDateRows: cashSource.rows.length - cashValidRows.length,
        excludedMovementRows: cashPeriodRows.length - cashAnalyticsRows.length,
        filteredRows: cashAnalyticsRows.length,
        startDate: cashRange.startDate ? cashFormatIsoDate_(cashRange.startDate) : null,
        endDate: cashRange.endDate ? cashFormatIsoDate_(cashRange.endDate) : null
      }
    });
  } catch (cashError) {
    cashLogError_(cashError);
    return cashErrorResponse_(cashError);
  }
}

/** @param {Object=} params Raw client parameters. @return {Object} Safe request. @private */
function cashNormalizeRequest_(params) {
  var cashParams = params && typeof params === 'object' ? params : {};
  var cashStart = cashParseDate_(cashParams.startDate);
  var cashEnd = cashParseDate_(cashParams.endDate);
  if (cashStart) { cashStart.setHours(0, 0, 0, 0); }
  if (cashEnd) { cashEnd.setHours(23, 59, 59, 999); }
  if (cashStart && cashEnd && cashStart.getTime() > cashEnd.getTime()) {
    throw new Error('Дата начала периода не может быть позже даты окончания.');
  }
  return {
    startDate: cashStart,
    endDate: cashEnd,
    account: cashString_(cashParams.account),
    counterparty: cashString_(cashParams.counterparty),
    article: cashString_(cashParams.article),
    direction: cashString_(cashParams.direction)
  };
}

/** @return {{rows: Array<Object>}} Parsed rows from rangecash. @private */
function cashReadNamedRange_() {
  var cashSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cashRange = cashSpreadsheet.getRangeByName(CASH_RANGE_NAME);
  if (!cashRange) {
    throw new Error('Не найден именованный диапазон «' + CASH_RANGE_NAME + '». Создайте его на листе «Деньги» вместе со строкой заголовков.');
  }
  var cashValues = cashRange.getValues();
  if (cashValues.length < 2) {
    return { rows: [] };
  }
  var cashHeaders = cashValues[0].map(cashHeaderKey_);
  CASH_REQUIRED_COLUMNS.forEach(function (cashColumn) {
    if (cashHeaders.indexOf(cashHeaderKey_(cashColumn)) === -1) {
      throw new Error('В диапазоне «' + CASH_RANGE_NAME + '» отсутствует обязательный столбец «' + cashColumn + '».');
    }
  });
  var cashHeaderIndex = {};
  cashHeaders.forEach(function (cashHeader, cashIndex) { cashHeaderIndex[cashHeader] = cashIndex; });
  var cashStartRow = cashRange.getRow() + 1;
  return {
    rows: cashValues.slice(1).reduce(function (cashRows, cashValuesRow, cashIndex) {
      if (cashIsEmptyRow_(cashValuesRow)) { return cashRows; }
      cashRows.push({
        date: cashParseDate_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Дата')),
        amount: cashNumber_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Сумма')),
        operation: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Вид операции')),
        account: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Счет')),
        counterparty: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Контрагент')),
        article: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Статья')),
        section: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Раздел ОДДС')),
        direction: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Направление')),
        description: cashString_(cashValueByHeader_(cashValuesRow, cashHeaderIndex, 'Описание')),
        sourceRow: cashStartRow + cashIndex
      });
      return cashRows;
    }, [])
  };
}

/** @param {Array<*>} row Source row. @return {boolean} Empty state. @private */
function cashIsEmptyRow_(row) {
  return row.every(function (cashValue) { return cashValue === '' || cashValue === null || typeof cashValue === 'undefined'; });
}

/** @param {Array<*>} row Row. @param {Object} index Headers. @param {string} header Header. @return {*} Value. @private */
function cashValueByHeader_(row, index, header) {
  var cashIndex = index[cashHeaderKey_(header)];
  return typeof cashIndex === 'undefined' ? null : row[cashIndex];
}

/** @param {*} value Header. @return {string} Header key. @private */
function cashHeaderKey_(value) { return cashString_(value).toLowerCase().replace(/\s+/g, ' '); }
/** @param {*} value Value. @return {string} String without surrounding whitespace. @private */
function cashString_(value) { return value === null || typeof value === 'undefined' ? '' : String(value).trim(); }
/** @param {*} value Number, possibly text ("-2 231,00", "1 234,56 ₽"). @return {number} Finite number. @private */
function cashNumber_(value) {
  if (typeof value === 'number') { return isFinite(value) ? value : 0; }
  if (typeof value !== 'string') { return 0; }
  var cashText = value.replace(/[\s\u00A0]/g, '').replace(/₽/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var cashParsed = Number(cashText);
  return isFinite(cashParsed) ? cashParsed : 0;
}

/** @return {string} Script time zone, matching gsprofit.gs/gscapital.gs. @private */
function cashTimeZone_() { return Session.getScriptTimeZone() || 'Etc/GMT'; }

/**
 * Parses a spreadsheet date without letting the script's execution time zone
 * shift the calendar day (aligned with profitParseLocalDate_/capitalParseDate_
 * so all three modules agree on which day a given row belongs to).
 * @param {*} value Date input. @return {Date|null} Strictly parsed date. @private */
function cashParseDate_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) { return null; }
    var cashIsoText = Utilities.formatDate(value, cashTimeZone_(), 'yyyy-MM-dd');
    return cashValidatedDate_(+cashIsoText.substring(0, 4), +cashIsoText.substring(5, 7), +cashIsoText.substring(8, 10));
  }
  if (typeof value !== 'string') { return null; }
  var cashInput = value.trim();
  var cashMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(cashInput);
  if (cashMatch) { return cashValidatedDate_(+cashMatch[1], +cashMatch[2], +cashMatch[3]); }
  cashMatch = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(cashInput);
  return cashMatch ? cashValidatedDate_(+cashMatch[3], +cashMatch[2], +cashMatch[1]) : null;
}

/** @return {Date|null} Valid local date. @private */
function cashValidatedDate_(year, month, day) {
  var cashDate = new Date(year, month - 1, day);
  return cashDate.getFullYear() === year && cashDate.getMonth() === month - 1 && cashDate.getDate() === day ? cashDate : null;
}

/** @param {Object} row Row. @param {Object} request Filters. @return {boolean} Match. @private */
function cashMatchesFilters_(row, request) {
  return (!request.account || row.account === request.account) &&
      (!request.counterparty || row.counterparty === request.counterparty) &&
      (!request.article || row.article === request.article) &&
      (!request.direction || row.direction === request.direction);
}

/** @return {boolean} Date belongs to range. @private */
function cashIsInPeriod_(date, startDate, endDate) {
  return date && (!startDate || date.getTime() >= startDate.getTime()) && (!endDate || date.getTime() <= endDate.getTime());
}

/** @param {Object} row Source row. @return {boolean} Internal transfer row. @private */
function cashIsMovement_(row) {
  var cashSection = row.section.toLowerCase();
  return /^5(?:\.|\s|$)/.test(cashSection) || cashSection.indexOf('перемещ') !== -1;
}

/** @param {Object} row Source row. @return {number} Signed cash movement. @private */
function cashSignedAmount_(row) {
  var cashOperation = row.operation.toLowerCase();
  if (cashOperation === 'приход') { return Math.abs(row.amount); }
  if (cashOperation === 'расход') { return -Math.abs(row.amount); }
  if (cashOperation === 'остаток') { return row.amount; }
  return 0;
}

/** @param {Object} row Source row. @param {string} operation Expected operation. @return {boolean} Match. @private */
function cashHasOperation_(row, operation) { return row.operation.toLowerCase() === operation.toLowerCase(); }
/** @param {Object} row Source row. @return {boolean} Deposit account. @private */
function cashIsDeposit_(row) { return row.account.toLowerCase().indexOf('депозит') !== -1; }
/** @param {Object} row Source row. @return {string} Activity category. @private */
function cashActivityType_(row) {
  var cashSection = row.section.toLowerCase();
  if (cashSection.indexOf('основн') !== -1) { return 'operating'; }
  if (cashSection.indexOf('инвести') !== -1) { return 'investing'; }
  if (cashSection.indexOf('финансов') !== -1) { return 'financing'; }
  return 'other';
}

/** @param {Array<Object>} rows Data. @param {Object} request Period selection. @return {Object} Range. @private */
function cashResolveVisibleRange_(rows, request) {
  var cashDates = rows.map(function (cashRow) { return cashRow.date; }).sort(function (a, b) { return a - b; });
  return {
    startDate: request.startDate || (cashDates.length ? new Date(cashDates[0].getTime()) : null),
    endDate: request.endDate || (cashDates.length ? new Date(cashDates[cashDates.length - 1].getTime()) : null)
  };
}

/** @param {Date|null} start Date. @param {Date|null} end Date. @param {Array<Object>} rows Fallback rows. @return {Array<Object>} Months. @private */
function cashBuildMonths_(start, end, rows) {
  if (!start || !end) { return []; }
  var cashMonths = [];
  var cashCursor = new Date(start.getFullYear(), start.getMonth(), 1);
  var cashLast = new Date(end.getFullYear(), end.getMonth(), 1);
  while (cashCursor.getTime() <= cashLast.getTime()) {
    cashMonths.push({ key: cashMonthKey_(cashCursor), label: CASH_MONTH_NAMES[cashCursor.getMonth()] + ' ' + cashCursor.getFullYear(), date: new Date(cashCursor.getTime()) });
    cashCursor.setMonth(cashCursor.getMonth() + 1);
  }
  return cashMonths;
}

/** @param {Date} date Date. @return {string} YYYY-MM. @private */
function cashMonthKey_(date) { return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2); }
/** @param {Date} date Date. @return {string} ISO day. @private */
function cashFormatIsoDate_(date) { return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2) + '-' + ('0' + date.getDate()).slice(-2); }

/** @param {Array<Object>} rows Analytics rows. @param {Object} range Period. @return {Object} KPI payload. @private */
function cashBuildKpis_(rows, range) {
  var cashFlowRows = rows.filter(function (cashRow) { return cashHasOperation_(cashRow, 'Приход') || cashHasOperation_(cashRow, 'Расход'); });
  var cashNet = cashFlowRows.reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0);
  var cashOperatingFlowRows = cashFlowRows.filter(function (cashRow) { return cashActivityType_(cashRow) === 'operating'; });
  var cashOperatingExpense = cashOperatingFlowRows.filter(function (cashRow) { return cashHasOperation_(cashRow, 'Расход'); }).reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0);
  // Остатки не суммируются по месяцам: для Cash Buffer нужен итог по всем
  // счетам на последнюю доступную дату остатка в выбранном периоде.
  var cashEndingBalance = cashEndingBalance_(rows, range.endDate);
  var cashBalance = cashEndingBalance.value;
  var cashDays = range.startDate && range.endDate ? Math.floor((new Date(range.endDate.getFullYear(), range.endDate.getMonth(), range.endDate.getDate()) - new Date(range.startDate.getFullYear(), range.startDate.getMonth(), range.startDate.getDate())) / 86400000) + 1 : 0;
  var cashSections = [
    { key: 'operating', label: 'Основная деятельность' },
    { key: 'investing', label: 'Инвестиционная деятельность' },
    { key: 'financing', label: 'Финансовая деятельность' }
  ].map(function (cashItem) {
    var cashValue = cashFlowRows.filter(function (cashRow) { return cashActivityType_(cashRow) === cashItem.key; }).reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0);
    return { key: cashItem.key, label: cashItem.label, value: cashValue, share: cashNet ? cashValue / cashNet : 0 };
  });
  var cashAverageDailyExpense = cashDays ? Math.abs(cashOperatingExpense) / cashDays : null;
  var cashBufferDays = cashAverageDailyExpense ? Math.abs(cashBalance / cashAverageDailyExpense) : null;
  Logger.log('[Деньги][CashBuffer] Расход с видом операции «Расход»: ' + cashOperatingExpense + '; среднедневной расход: ' + cashAverageDailyExpense + '; чистый денежный поток на ' + (cashEndingBalance.date ? cashFormatIsoDate_(cashEndingBalance.date) : 'нет даты') + ': ' + cashBalance + '; Cash Buffer = |чистый поток / среднедневной расход| = ' + cashBufferDays + '; строк остатка: ' + cashEndingBalance.count);
  return { netCashFlow: cashNet, cashBufferDays: cashBufferDays, cashBufferBalanceDate: cashEndingBalance.date ? cashFormatIsoDate_(cashEndingBalance.date) : null, calendarDays: cashDays, sectionFlow: cashSections };
}

/** @param {Array<Object>} rows Cash rows. @param {Date|null} end Period end. @return {{value:number|null,date:Date|null,count:number}} Last balance snapshot. @private */
function cashEndingBalance_(rows, end) {
  var cashBalanceRows = rows.filter(function (cashRow) {
    return cashHasOperation_(cashRow, 'Остаток') && cashRow.date && (!end || cashRow.date.getTime() <= end.getTime());
  });
  if (!cashBalanceRows.length) {
    var cashFallback = rows.filter(function (cashRow) {
      return cashRow.date && (!end || cashRow.date.getTime() <= end.getTime());
    }).reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0);
    Logger.log('[Деньги][CashBuffer] В источнике нет отдельных строк остатка; использован накопленный остаток операций: ' + cashFallback);
    return { value: cashFallback, date: end ? new Date(end.getTime()) : null, count: 0 };
  }
  var cashBalanceDate = cashBalanceRows.reduce(function (cashLatest, cashRow) {
    return cashRow.date.getTime() > cashLatest.getTime() ? cashRow.date : cashLatest;
  }, cashBalanceRows[0].date);
  var cashRowsAtEnd = cashBalanceRows.filter(function (cashRow) { return cashRow.date.getTime() === cashBalanceDate.getTime(); });
  return {
    value: cashRowsAtEnd.reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0),
    date: new Date(cashBalanceDate.getTime()),
    count: cashRowsAtEnd.length
  };
}

/** @param {Array<Object>} rows Rows. @param {string} operation Operation. @return {Array<Object>} Chart items. @private */
function cashBuildCounterparties_(rows, operation) {
  var cashGroups = {};
  rows.filter(function (cashRow) { return cashHasOperation_(cashRow, operation) && !cashIsDeposit_(cashRow); }).forEach(function (cashRow) {
    var cashKey = cashRow.counterparty || 'Не указан';
    cashGroups[cashKey] = (cashGroups[cashKey] || 0) + Math.abs(cashRow.amount);
  });
  var cashTotal = Object.keys(cashGroups).reduce(function (cashSum, cashKey) { return cashSum + cashGroups[cashKey]; }, 0);
  return Object.keys(cashGroups).map(function (cashKey) {
    return { label: cashKey, value: cashGroups[cashKey], share: cashTotal ? cashGroups[cashKey] / cashTotal : 0, alert: cashTotal ? cashGroups[cashKey] / cashTotal > 0.3 : false };
  }).sort(function (a, b) { return b.value - a.value; });
}

/** @param {Array<Object>} rows Rows. @return {Array<Object>} Chart items. @private */
function cashBuildDirections_(rows) {
  var cashGroups = {};
  rows.forEach(function (cashRow) {
    var cashKey = cashRow.direction || 'Без направления';
    cashGroups[cashKey] = (cashGroups[cashKey] || 0) + cashSignedAmount_(cashRow);
  });
  var cashTotal = Object.keys(cashGroups).reduce(function (cashSum, cashKey) { return cashSum + cashGroups[cashKey]; }, 0);
  return Object.keys(cashGroups).map(function (cashKey) {
    return { label: cashKey, value: cashGroups[cashKey], share: cashTotal ? cashGroups[cashKey] / cashTotal : 0 };
  }).sort(function (a, b) { return Math.abs(b.value) - Math.abs(a.value); });
}

/** @param {Array<Object>} rows Rows. @param {Array<Object>} months Months. @return {Array<Object>} Chart items. @private */
function cashBuildMonthlyNetFlow_(rows, months) {
  var cashValues = {};
  rows.forEach(function (cashRow) { cashValues[cashMonthKey_(cashRow.date)] = (cashValues[cashMonthKey_(cashRow.date)] || 0) + cashSignedAmount_(cashRow); });
  return months.map(function (cashMonth) { return { label: cashMonth.label, value: cashValues[cashMonth.key] || 0 }; });
}

/** @param {Array<Object>} rows All categorical historical rows. @param {Array<Object>} months Visible months. @return {Object} Balance chart. @private */
function cashBuildBalances_(rows, months) {
  var cashSorted = rows.slice().sort(function (a, b) { return a.date - b.date; });
  var cashBalanceDrillId = cashBalanceDrillId_();
  var cashDetails = {};
  cashDetails[cashBalanceDrillId] = {};
  var cashPoints = months.map(function (cashMonth) {
    var cashStart = cashMonth.date.getTime();
    var cashBalanceRows = cashSorted.filter(function (cashRow) {
      return cashRow.date.getTime() < cashStart ||
          (cashHasOperation_(cashRow, 'Остаток') && cashRow.date.getTime() === cashStart);
    });
    cashDetails[cashBalanceDrillId][cashMonth.key] = cashBalanceRows.map(function (cashRow) {
      return cashOperationDetail_(cashRow);
    });
    var cashBalance = cashBalanceRows.reduce(function (cashSum, cashRow) { return cashSum + cashSignedAmount_(cashRow); }, 0);
    return { key: cashMonth.key, label: cashMonth.label, value: cashBalance };
  });
  var cashOperatingExpenses = months.map(function (cashMonth) {
    return cashSorted.filter(function (cashRow) { return cashMonthKey_(cashRow.date) === cashMonth.key && cashHasOperation_(cashRow, 'Расход') && cashActivityType_(cashRow) === 'operating'; }).reduce(function (cashSum, cashRow) { return cashSum + Math.abs(cashRow.amount); }, 0);
  });
  var cashAverage = cashOperatingExpenses.length ? cashOperatingExpenses.reduce(function (a, b) { return a + b; }, 0) / cashOperatingExpenses.length : 0;
  return { points: cashPoints, averageOperatingExpense: cashAverage, details: cashDetails };
}

/** @param {Array<Object>} rows Rows. @return {Array<Object>} Chart items. @private */
function cashBuildExpenseStructure_(rows) {
  var cashGroups = {};
  rows.filter(function (cashRow) { return cashHasOperation_(cashRow, 'Расход'); }).forEach(function (cashRow) {
    var cashKey = cashRow.article || 'Без статьи';
    cashGroups[cashKey] = (cashGroups[cashKey] || 0) + Math.abs(cashRow.amount);
  });
  return Object.keys(cashGroups).map(function (cashKey) { return { label: cashKey, value: cashGroups[cashKey] }; }).sort(function (a, b) { return b.value - a.value; });
}

/** @param {Array<Object>} rows Rows. @return {Object} Select options. @private */
function cashBuildFilterOptions_(rows) {
  return {
    accounts: cashUniqueValues_(rows, 'account'),
    counterparties: cashUniqueValues_(rows, 'counterparty'),
    articles: cashUniqueValues_(rows, 'article'),
    directions: cashUniqueValues_(rows, 'direction')
  };
}

/** @param {Array<Object>} rows Rows. @param {string} key Key. @return {Array<string>} Unique values. @private */
function cashUniqueValues_(rows, key) {
  var cashSet = {};
  rows.forEach(function (cashRow) { if (cashRow[key]) { cashSet[cashRow[key]] = true; } });
  return Object.keys(cashSet).sort(function (a, b) { return a.localeCompare(b, 'ru'); });
}

/** @param {Array<Object>} rows Rows. @param {Array<Object>} months Months. @param {Object} balances Balance chart. @return {Object} Report. @private */
function cashBuildReport_(rows, months, balances) {
  var cashMonthKeys = months.map(function (cashMonth) { return cashMonth.key; });
  var cashGroups = {};
  var cashDetails = {};
  rows.filter(function (cashRow) { return !cashHasOperation_(cashRow, 'Остаток'); }).forEach(function (cashRow) {
    var cashSection = cashRow.section || 'Без раздела ОДДС';
    var cashOperation = cashRow.operation || 'Без вида операции';
    var cashArticle = cashRow.article || 'Без статьи';
    var cashMonthKey = cashMonthKey_(cashRow.date);
    if (!cashGroups[cashSection]) { cashGroups[cashSection] = {}; }
    if (!cashGroups[cashSection][cashOperation]) { cashGroups[cashSection][cashOperation] = {}; }
    if (!cashGroups[cashSection][cashOperation][cashArticle]) { cashGroups[cashSection][cashOperation][cashArticle] = cashEmptyMonthValues_(cashMonthKeys); }
    cashGroups[cashSection][cashOperation][cashArticle][cashMonthKey] += cashSignedAmount_(cashRow);
    cashAddDrillDetail_(cashDetails, cashSectionDrillId_(cashSection), cashMonthKey, cashRow);
    cashAddDrillDetail_(cashDetails, cashOperationDrillId_(cashSection, cashOperation), cashMonthKey, cashRow);
    cashAddDrillDetail_(cashDetails, cashArticleDrillId_(cashSection, cashOperation, cashArticle), cashMonthKey, cashRow);
  });
  var cashSections = Object.keys(cashGroups).sort(function (a, b) { return a.localeCompare(b, 'ru'); }).map(function (cashSection) {
    var cashOperations = Object.keys(cashGroups[cashSection]).sort(function (a, b) { return a.localeCompare(b, 'ru'); }).map(function (cashOperation) {
      var cashArticles = Object.keys(cashGroups[cashSection][cashOperation]).sort(function (a, b) { return a.localeCompare(b, 'ru'); }).map(function (cashArticle) {
        var cashValues = cashGroups[cashSection][cashOperation][cashArticle];
        return { article: cashArticle, values: cashValues, total: cashSumMonthValues_(cashValues), drillId: cashArticleDrillId_(cashSection, cashOperation, cashArticle) };
      });
      return { operation: cashOperation, articles: cashArticles, values: cashSumNodes_(cashArticles, cashMonthKeys), total: cashSumNodes_(cashArticles, cashMonthKeys).total, drillId: cashOperationDrillId_(cashSection, cashOperation) };
    });
    return { section: cashSection, operations: cashOperations, values: cashSumNodes_(cashOperations, cashMonthKeys), total: cashSumNodes_(cashOperations, cashMonthKeys).total, drillId: cashSectionDrillId_(cashSection) };
  });
  var cashBalanceValues = cashEmptyMonthValues_(cashMonthKeys);
  balances.points.forEach(function (cashPoint) { cashBalanceValues[cashPoint.key] = cashPoint.value; });
  cashBalanceValues.total = null;
  Object.keys(balances.details || {}).forEach(function (cashDrillId) { cashDetails[cashDrillId] = balances.details[cashDrillId]; });
  return { months: months, openingBalance: { values: cashBalanceValues, total: null, drillId: cashBalanceDrillId_() }, sections: cashSections, details: cashDetails };
}

/** @param {Object} details Detail map. @param {string} drillId Drilldown key. @param {string} monthKey Month key. @param {Object} row Source row. @private */
function cashAddDrillDetail_(details, drillId, monthKey, row) {
  if (!details[drillId]) { details[drillId] = {}; }
  if (!details[drillId][monthKey]) { details[drillId][monthKey] = []; }
  details[drillId][monthKey].push(cashOperationDetail_(row));
}

/** @param {string} section Section name. @return {string} Drilldown key. @private */
function cashSectionDrillId_(section) { return 'cash:section:' + encodeURIComponent(section); }
/** @param {string} section Section name. @param {string} operation Operation type. @return {string} Drilldown key. @private */
function cashOperationDrillId_(section, operation) { return 'cash:operation:' + encodeURIComponent(section) + '|' + encodeURIComponent(operation); }
/** @param {string} section Section name. @param {string} operation Operation type. @param {string} article Article. @return {string} Drilldown key. @private */
function cashArticleDrillId_(section, operation, article) { return 'cash:article:' + encodeURIComponent(section) + '|' + encodeURIComponent(operation) + '|' + encodeURIComponent(article); }
/** @return {string} Drilldown key. @private */
function cashBalanceDrillId_() { return 'cash:balance:opening'; }

/** @param {Array<string>} monthKeys Keys. @return {Object} Empty values. @private */
function cashEmptyMonthValues_(monthKeys) { return monthKeys.reduce(function (cashObject, cashKey) { cashObject[cashKey] = 0; return cashObject; }, {}); }
/** @param {Object} values Values. @return {number} Total. @private */
function cashSumMonthValues_(values) { return Object.keys(values).reduce(function (cashSum, cashKey) { return cashSum + values[cashKey]; }, 0); }
/** @param {Array<Object>} nodes Tree nodes. @param {Array<string>} keys Keys. @return {Object} Sum values. @private */
function cashSumNodes_(nodes, keys) {
  var cashValues = cashEmptyMonthValues_(keys);
  nodes.forEach(function (cashNode) { keys.forEach(function (cashKey) { cashValues[cashKey] += cashNode.values[cashKey] || 0; }); });
  cashValues.total = cashSumMonthValues_(cashValues);
  return cashValues;
}
/** @param {Object} row Row. @return {Object} Drill item. @private */
function cashOperationDetail_(row) {
  return { date: cashFormatIsoDate_(row.date), description: row.description, account: row.account, counterparty: row.counterparty, direction: row.direction, amount: cashSignedAmount_(row), sourceRow: row.sourceRow };
}

/** @param {*} data Payload. @return {Object} Controlled response. @private */
function cashSuccessResponse_(data) {
  var cashResponse = { status: 'success', data: data, error: null };
  return typeof prepareResponse === 'function' ? prepareResponse(cashResponse) : cashResponse;
}
/** @param {*} error Error. @return {Object} Controlled response. @private */
function cashErrorResponse_(error) {
  var cashResponse = { status: 'error', data: null, error: error && error.message ? error.message : String(error || 'Неизвестная ошибка.') };
  return typeof prepareResponse === 'function' ? prepareResponse(cashResponse) : cashResponse;
}
/** @param {*} error Error. @private */
function cashLogError_(error) { Logger.log('[CASH ERROR] ' + (error && error.stack ? error.stack : String(error))); }
