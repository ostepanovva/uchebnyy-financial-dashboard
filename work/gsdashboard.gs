/** @OnlyCurrentDoc */
/**
 * Dashboard module ("Дэшборд" page).
 *
 * Combines rangecash ("Деньги"), rangeprofit ("Прибыль") and rangeequity
 * ("Капитал") into one summary page. All functions and identifiers in this
 * file start with "dashboard" so they never collide with the money/profit/
 * capital modules, which each read their own named range independently.
 *
 * Key rule (see task spec): "Капитал" only has snapshot dates (e.g. 01.01,
 * 30.06), while "Деньги"/"Прибыль" have day-by-day flows. Every calculation
 * on this page anchors on snapshotDate — the latest "Капитал" date that is
 * not later than the date the user picked. The report start is selected by
 * the user; when it is empty, the earliest available source date is used.
 */

var DASHBOARD_CASH_RANGE_NAME = 'rangecash';
var DASHBOARD_PROFIT_RANGE_NAME = 'rangeprofit';
var DASHBOARD_EQUITY_RANGE_NAME = 'rangeequity';

var DASHBOARD_CASH_REQUIRED_COLUMNS = ['Дата', 'Сумма', 'Вид операции', 'Раздел ОДДС', 'Направление'];
var DASHBOARD_PROFIT_REQUIRED_COLUMNS = ['Дата', 'Сумма', 'Вид операции', 'Направление', 'Раздел ОПиУ'];
var DASHBOARD_EQUITY_REQUIRED_COLUMNS = ['Дата', 'Раздел 2', 'Раздел 3', 'Сумма'];

var DASHBOARD_MONTH_NAMES = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

var DASHBOARD_PROFIT_SECTION_CODES = {
  REVENUE: '1.выручка',
  COST: '2.себестоимость',
  OVERHEAD: '3.общепроизводственные',
  COMMERCIAL: '4.коммерческиерасходы',
  ADMINISTRATIVE: '5.административныерасходы',
  OTHER_INCOME: '6.доходынижеоперационнойприбыли',
  OTHER_EXPENSES: '7.расходынижеоперационнойприбыли'
};

/**
 * Public endpoint used by dashboard.html through apiDashboard (gscore.gs).
 * @param {Object=} params { startDate: 'YYYY-MM-DD', endDate: 'YYYY-MM-DD' }.
 *     The end defaults to the latest available "Капитал" snapshot.
 * @return {Object} Standard {status, data, error} response.
 */
function dashboardGetDashboardData(params) {
  try {
    var dashboardRequest = dashboardNormalizeRequest_(params);
    var dashboardCashSource = dashboardReadCashRows_();
    var dashboardProfitSource = dashboardReadProfitRows_();
    var dashboardEquitySource = dashboardReadEquityRows_();

    var dashboardEquityDates = dashboardUniqueDates_(dashboardEquitySource.rows);
    if (!dashboardEquityDates.length) {
      return dashboardErrorResponse_('В диапазоне «' + DASHBOARD_EQUITY_RANGE_NAME + '» ("Капитал") нет ни одной строки с корректной датой — невозможно определить дату среза капитала.');
    }

    var dashboardRequestedEndDate = dashboardRequest.endDate || dashboardEquityDates[dashboardEquityDates.length - 1];

    var dashboardSnapshotDate = dashboardFindOnOrBeforeDate_(dashboardEquityDates, dashboardRequestedEndDate);
    if (!dashboardSnapshotDate) {
      return dashboardErrorResponse_('В листе «Капитал» нет ни одной записи на ' + dashboardFormatDate_(dashboardRequestedEndDate) + ' или раньше. Самая ранняя доступная дата среза — ' + dashboardFormatDate_(dashboardEquityDates[0]) + '.');
    }

    var dashboardWarning = null;
    if (dashboardSnapshotDate.getTime() < dashboardRequestedEndDate.getTime()) {
      dashboardWarning = 'Данные по капиталу заканчиваются на ' + dashboardFormatDate_(dashboardSnapshotDate) +
          ', это раньше запрошенной даты ' + dashboardFormatDate_(dashboardRequestedEndDate) +
          '. Дашборд показан на последнюю доступную дату среза.';
    }

    // Период отчёта управляется полем «Начало периода». Больше нет скрытой
    // подстановки 1 января: пустое поле берёт первую фактическую дату источников.
    var dashboardFallbackStart = dashboardEarliestSourceDate_(
        [dashboardCashSource.rows, dashboardProfitSource.rows, dashboardEquitySource.rows], dashboardSnapshotDate);
    var dashboardPeriodStart = dashboardRequest.startDate ? new Date(dashboardRequest.startDate.getTime()) :
        new Date((dashboardFallbackStart || dashboardSnapshotDate).getTime());
    if (dashboardPeriodStart.getTime() > dashboardSnapshotDate.getTime()) {
      return dashboardErrorResponse_('Дата начала периода (' + dashboardFormatDate_(dashboardPeriodStart) +
          ') позже даты среза (' + dashboardFormatDate_(dashboardSnapshotDate) + ').');
    }
    dashboardPeriodStart.setHours(0, 0, 0, 0);
    var dashboardPeriodEnd = new Date(dashboardSnapshotDate.getTime());
    dashboardPeriodEnd.setHours(23, 59, 59, 999);

    var dashboardCashPeriodRows = dashboardFilterPeriod_(dashboardCashSource.rows, dashboardPeriodStart, dashboardPeriodEnd);
    var dashboardCashFlowRows = dashboardCashPeriodRows.filter(function (dashboardRow) {
      return !dashboardCashIsMovement_(dashboardRow) &&
          (dashboardIsOperation_(dashboardRow.operation, 'Приход') || dashboardIsOperation_(dashboardRow.operation, 'Расход'));
    });

    var dashboardProfitPeriodRows = dashboardFilterPeriod_(dashboardProfitSource.rows, dashboardPeriodStart, dashboardPeriodEnd);
    var dashboardMonths = dashboardBuildMonths_(dashboardPeriodStart, dashboardPeriodEnd);

    var dashboardKpis = dashboardBuildKpis_(
        dashboardCashFlowRows,
        dashboardCashPeriodRows,
        dashboardProfitPeriodRows,
        dashboardProfitSource.rows,
        dashboardEquitySource.rows,
        dashboardEquityDates,
        dashboardSnapshotDate,
        dashboardPeriodStart
    );

    dashboardKpis.trends = dashboardBuildTrends_(
        dashboardCashSource.rows,
        dashboardProfitSource.rows,
        dashboardEquitySource.rows,
        dashboardEquityDates,
        dashboardSnapshotDate,
        dashboardKpis,
        dashboardPeriodStart
    );

    var dashboardCharts = {
      cashMonthlyNetFlow: dashboardBuildCashMonthlyNetFlow_(dashboardCashFlowRows, dashboardMonths),
      cashByDirection: dashboardBuildCashByDirection_(dashboardCashFlowRows, dashboardMonths),
      revenueAndGrossMargin: dashboardBuildRevenueAndGrossMargin_(dashboardProfitPeriodRows, dashboardMonths),
      grossProfitByDirection: dashboardBuildProfitDirectionSeries_(dashboardProfitPeriodRows, dashboardMonths),
      netProfitAndMargin: dashboardBuildNetProfitAndMargin_(dashboardProfitPeriodRows, dashboardMonths),
      receivables: dashboardBuildEquitySectionSeries_(dashboardEquitySource.rows, dashboardPeriodStart, dashboardPeriodEnd, 'Дебиторская задолженность'),
      payables: dashboardBuildEquitySectionSeries_(dashboardEquitySource.rows, dashboardPeriodStart, dashboardPeriodEnd, 'Кредиторская задолженность'),
      capitalMix: dashboardBuildCapitalMix_(dashboardEquitySource.rows, dashboardPeriodStart, dashboardPeriodEnd),
      currentRatioTrend: dashboardBuildCurrentRatioTrend_(dashboardEquitySource.rows, dashboardPeriodStart, dashboardPeriodEnd)
    };

    return dashboardSuccessResponse_({
      kpis: dashboardKpis,
      charts: dashboardCharts,
      meta: {
        requestedEndDate: dashboardFormatIsoDate_(dashboardRequestedEndDate),
        snapshotDate: dashboardFormatIsoDate_(dashboardSnapshotDate),
        periodStart: dashboardFormatIsoDate_(dashboardPeriodStart),
        periodEnd: dashboardFormatIsoDate_(dashboardPeriodEnd),
        calendarDays: dashboardKpis.calendarDays,
        warning: dashboardWarning
      }
    });
  } catch (dashboardError) {
    dashboardLogError_('dashboardGetDashboardData: ' + (dashboardError && dashboardError.message ? dashboardError.message : dashboardError));
    return dashboardErrorResponse_(dashboardError && dashboardError.message ? dashboardError.message : 'Не удалось подготовить данные дашборда.');
  }
}

/** @param {Object=} params Raw client params. @return {{endDate: Date|null, startDate: Date|null}} Normalized request. @private */
function dashboardNormalizeRequest_(params) {
  var dashboardParams = params && typeof params === 'object' ? params : {};
  return {
    endDate: dashboardParseDate_(dashboardParams.endDate),
    startDate: dashboardParseDate_(dashboardParams.startDate)
  };
}

// ---------------------------------------------------------------------------
// KPI cards
// ---------------------------------------------------------------------------

/**
 * @param {Array<Object>} dashboardCashFlowRows Приход/Расход rows within the period, movements excluded (used for KPI 1 and charts).
 * @param {Array<Object>} dashboardCashPeriodRows ALL "Деньги" rows within the period, unfiltered by operation/movement (used for Cash Buffer, KPI 2).
 * @param {Array<Object>} dashboardProfitPeriodRows Profit rows within the period.
 * @param {Array<Object>} dashboardEquityRows All equity rows (not period-filtered — needed for "on or before" lookups).
 * @param {Array<Date>} dashboardEquityDates Unique sorted equity dates.
 * @param {Date} dashboardSnapshotDate Date D.
 * @param {Date} dashboardPeriodStart Date S (period start — 1 January of D's year, or the user's chosen start).
 * @return {Object} All ten KPI values.
 * @private
 */
function dashboardBuildKpis_(dashboardCashFlowRows, dashboardCashPeriodRows, dashboardProfitPeriodRows, dashboardAllProfitRows, dashboardEquityRows, dashboardEquityDates, dashboardSnapshotDate, dashboardPeriodStart) {
  // 1) Чистый денежный поток (все Приходы/Расходы за период, движения между своими счетами ("5. Перемещение") исключены).
  var dashboardNetCashFlow = dashboardCashFlowRows.reduce(function (dashboardSum, dashboardRow) {
    return dashboardSum + dashboardCashSignedAmount_(dashboardRow);
  }, 0);

  // 2) Cash Buffer.
  // Формула: |чистый денежный поток / среднедневной расход|.
  // Среднедневной расход — строки «Расход» раздела «2. Основная деятельность» за период.
  Logger.log('[Dashboard][CashBuffer] --- начало расчёта ---');
  var dashboardCalendarDays = dashboardCalendarDays_(dashboardPeriodStart, dashboardSnapshotDate);
  Logger.log('[Dashboard][CashBuffer] Период: ' + dashboardFormatDate_(dashboardPeriodStart) + ' — ' + dashboardFormatDate_(dashboardSnapshotDate) + ', дней: ' + dashboardCalendarDays);

  var dashboardOperatingExpenseRows = dashboardCashPeriodRows.filter(function (dashboardRow) {
    return dashboardCashIsOperatingActivity_(dashboardRow) && dashboardIsOperation_(dashboardRow.operation, 'Расход');
  });
  var dashboardOperatingExpenseTotal = dashboardOperatingExpenseRows.reduce(function (dashboardSum, dashboardRow) {
    return dashboardSum + dashboardCashSignedAmount_(dashboardRow);
  }, 0);
  // Остаток — это срез, а не поток. Берём все счета на последней дате
  // остатка внутри периода, а не складываем остатки за все месяцы.
  var dashboardEndingBalance = dashboardCashEndingBalance_(dashboardCashPeriodRows, dashboardSnapshotDate);
  var dashboardBalanceTotal = dashboardEndingBalance.value;
  Logger.log('[Dashboard][CashBuffer] Расходов с видом операции «Расход»: ' + dashboardOperatingExpenseRows.length +
      ' -> итог расхода = ' + dashboardOperatingExpenseTotal);
  Logger.log('[Dashboard][CashBuffer] Строк «Остаток» на ' +
      (dashboardEndingBalance.date ? dashboardFormatDate_(dashboardEndingBalance.date) : 'нет даты') + ': ' +
      dashboardEndingBalance.count + ' -> итог остатка денежных средств = ' + dashboardBalanceTotal);

  var dashboardAverageDailyExpense = dashboardCalendarDays ? Math.abs(dashboardOperatingExpenseTotal) / dashboardCalendarDays : null;
  var dashboardCashBufferDays = dashboardAverageDailyExpense ? Math.abs(dashboardBalanceTotal / dashboardAverageDailyExpense) : null;
  Logger.log('[Dashboard][CashBuffer] Среднедневной расход = |' + dashboardOperatingExpenseTotal + '| / ' + dashboardCalendarDays + ' = ' + dashboardAverageDailyExpense);
  Logger.log('[Dashboard][CashBuffer] Итог: Cash Buffer = |' + dashboardBalanceTotal + ' / ' + dashboardAverageDailyExpense + '| = ' + dashboardCashBufferDays);
  Logger.log('[Dashboard][CashBuffer] --- конец расчёта ---');

  // Снимок капитала на дату D — используется в KPI 3 и 7.
  var dashboardSnapshotRows = dashboardRowsForDate_(dashboardEquityRows, dashboardSnapshotDate);

  // 3) Коэффициент текущей ликвидности (на дату D — последнюю дату среза капитала).
  // Формула по требованиям: Сумма по Разделу 3 = "Краткосрочные обязательства" /
  // Сумма по Разделу 2 = "Оборотные активы".
  Logger.log('[Dashboard][CurrentRatio] --- начало расчёта ---');
  Logger.log('[Dashboard][CurrentRatio] Дата среза (последняя доступная дата капитала <= выбранной): ' + dashboardFormatDate_(dashboardSnapshotDate));
  Logger.log('[Dashboard][CurrentRatio] Строк "Капитал" на дату среза: ' + dashboardSnapshotRows.length);
  var dashboardCurrentAssets = dashboardSumWhere_(dashboardSnapshotRows, 'level2', 'Оборотные активы');
  var dashboardShortTermDebt = dashboardSumWhere_(dashboardSnapshotRows, 'level3', 'Краткосрочные обязательства');
  Logger.log('[Dashboard][CurrentRatio] Сумма "Раздел 2"="Оборотные активы": ' + dashboardCurrentAssets);
  Logger.log('[Dashboard][CurrentRatio] Сумма "Раздел 3"="Краткосрочные обязательства": ' + dashboardShortTermDebt);
  var dashboardCurrentRatio = dashboardCurrentAssets ? dashboardShortTermDebt / dashboardCurrentAssets : null;
  Logger.log('[Dashboard][CurrentRatio] Итог: ' + dashboardShortTermDebt + ' / ' + dashboardCurrentAssets + ' = ' + dashboardCurrentRatio);
  Logger.log('[Dashboard][CurrentRatio] --- конец расчёта ---');

  // 4-6) Прибыль
  var dashboardRevenue = dashboardSumProfitSection_(dashboardProfitPeriodRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
  var dashboardCost = dashboardSumProfitSection_(dashboardProfitPeriodRows, DASHBOARD_PROFIT_SECTION_CODES.COST);
  var dashboardOverhead = dashboardSumProfitSection_(dashboardProfitPeriodRows, DASHBOARD_PROFIT_SECTION_CODES.OVERHEAD);
  var dashboardGrossProfit = dashboardRevenue + dashboardCost + dashboardOverhead;
  var dashboardGrossMargin = dashboardRevenue ? dashboardGrossProfit / dashboardRevenue : null;
  var dashboardNetProfit = dashboardNetProfitForRows_(dashboardProfitPeriodRows);
  var dashboardNetMargin = dashboardRevenue ? dashboardNetProfit / dashboardRevenue : null;

  // 7) Собственный капитал (на дату D)
  var dashboardEquity = dashboardSumWhere_(dashboardSnapshotRows, 'level2', 'Капитал');

  // 8) ROE
  var dashboardEquityStart = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardPeriodStart, 'level2', 'Капитал');
  var dashboardRoe = {
    ratio: 0,
    percent: 0,
    periodStart: dashboardFormatIsoDate_(dashboardPeriodStart),
    periodEnd: dashboardFormatIsoDate_(dashboardSnapshotDate),
    startDate: null,
    endDate: dashboardFormatIsoDate_(dashboardSnapshotDate),
    error: null
  };
  if (!dashboardEquityStart.found) {
    dashboardRoe.error = 'Нет данных по капиталу на ' + dashboardFormatDate_(dashboardPeriodStart) + ' или раньше — невозможно посчитать средний капитал для ROE.';
  } else {
    var dashboardAvgEquity = (dashboardEquityStart.value + dashboardEquity) / 2;
    dashboardRoe.ratio = dashboardAvgEquity ? dashboardNetProfit / dashboardAvgEquity : 0;
    dashboardRoe.percent = dashboardRoe.ratio * 100;
    dashboardRoe.startDate = dashboardFormatIsoDate_(dashboardEquityStart.date);
  }

  // 9-10) Оборачиваемость ДЗ и КЗ всегда считается от 1 января года даты
  // среза D до D, независимо от выбранного пользователем начала отчёта.
  var dashboardTurnoverPeriodStart = new Date(dashboardSnapshotDate.getFullYear(), 0, 1);
  dashboardTurnoverPeriodStart.setHours(0, 0, 0, 0);
  var dashboardTurnoverCalendarDays = dashboardCalendarDays_(dashboardTurnoverPeriodStart, dashboardSnapshotDate);
  var dashboardTurnoverProfitRows = dashboardFilterPeriod_(dashboardAllProfitRows, dashboardTurnoverPeriodStart, dashboardSnapshotDate);
  var dashboardTurnoverRevenue = dashboardSumProfitSection_(dashboardTurnoverProfitRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
  var dashboardTurnoverCostAbs = Math.abs(dashboardSumProfitSection_(dashboardTurnoverProfitRows, DASHBOARD_PROFIT_SECTION_CODES.COST));
  Logger.log('[Dashboard][Turnover] Период: ' + dashboardFormatDate_(dashboardTurnoverPeriodStart) + ' — ' + dashboardFormatDate_(dashboardSnapshotDate) + '; дней: ' + dashboardTurnoverCalendarDays + '; выручка: ' + dashboardTurnoverRevenue + '; себестоимость: ' + dashboardTurnoverCostAbs);

  // 9) Оборачиваемость дебиторской задолженности
  var dashboardReceivablesEnd = dashboardSumWhere_(dashboardSnapshotRows, 'level3', 'Дебиторская задолженность');
  var dashboardReceivablesStart = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardTurnoverPeriodStart, 'level3', 'Дебиторская задолженность');
  var dashboardReceivablesTurnover = {
    ratio: 0,
    days: 0,
    startDate: null,
    endDate: dashboardFormatIsoDate_(dashboardSnapshotDate),
    error: null
  };
  if (!dashboardReceivablesStart.found) {
    dashboardReceivablesTurnover.error = 'Нет данных по дебиторской задолженности на ' + dashboardFormatDate_(dashboardTurnoverPeriodStart) + ' или раньше.';
  } else {
    var dashboardAvgReceivables = (dashboardReceivablesStart.value + dashboardReceivablesEnd) / 2;
    if (dashboardAvgReceivables && dashboardTurnoverRevenue) {
      dashboardReceivablesTurnover.ratio = dashboardTurnoverRevenue / dashboardAvgReceivables;
      dashboardReceivablesTurnover.days = dashboardReceivablesTurnover.ratio ? dashboardTurnoverCalendarDays / dashboardReceivablesTurnover.ratio : 0;
    }
    dashboardReceivablesTurnover.startDate = dashboardFormatIsoDate_(dashboardReceivablesStart.date);
  }

  // 10) Оборачиваемость кредиторской задолженности
  var dashboardPayablesEnd = dashboardSumWhere_(dashboardSnapshotRows, 'level3', 'Кредиторская задолженность');
  var dashboardPayablesStart = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardTurnoverPeriodStart, 'level3', 'Кредиторская задолженность');
  var dashboardPayablesTurnover = {
    ratio: 0,
    days: 0,
    startDate: null,
    endDate: dashboardFormatIsoDate_(dashboardSnapshotDate),
    error: null
  };
  if (!dashboardPayablesStart.found) {
    dashboardPayablesTurnover.error = 'Нет данных по кредиторской задолженности на ' + dashboardFormatDate_(dashboardTurnoverPeriodStart) + ' или раньше.';
  } else {
    var dashboardAvgPayables = (dashboardPayablesStart.value + dashboardPayablesEnd) / 2;
    if (dashboardAvgPayables && dashboardTurnoverCostAbs) {
      dashboardPayablesTurnover.ratio = dashboardTurnoverCostAbs / dashboardAvgPayables;
      dashboardPayablesTurnover.days = dashboardPayablesTurnover.ratio ? dashboardTurnoverCalendarDays / dashboardPayablesTurnover.ratio : 0;
    }
    dashboardPayablesTurnover.startDate = dashboardFormatIsoDate_(dashboardPayablesStart.date);
  }

  return {
    netCashFlow: dashboardNetCashFlow,
    cashBufferDays: dashboardCashBufferDays,
    cashBufferBalanceDate: dashboardEndingBalance.date ? dashboardFormatIsoDate_(dashboardEndingBalance.date) : null,
    calendarDays: dashboardCalendarDays,
    currentRatio: dashboardCurrentRatio,
    revenue: dashboardRevenue,
    netProfit: dashboardNetProfit,
    netMargin: dashboardNetMargin,
    grossMargin: dashboardGrossMargin,
    equity: dashboardEquity,
    roe: dashboardRoe,
    receivablesTurnover: dashboardReceivablesTurnover,
    payablesTurnover: dashboardPayablesTurnover
  };
}

/** @param {Array<Object>} dashboardRows Profit rows (any period). @return {number} Net profit, sections 1-7 only (section 8 is out of scope, see gsprofit.gs). @private */
function dashboardNetProfitForRows_(dashboardRows) {
  var dashboardAllowed = {};
  Object.keys(DASHBOARD_PROFIT_SECTION_CODES).forEach(function (dashboardKey) { dashboardAllowed[DASHBOARD_PROFIT_SECTION_CODES[dashboardKey]] = true; });
  return dashboardRows.filter(function (dashboardRow) { return dashboardAllowed[dashboardSectionKey_(dashboardRow.section)]; })
      .reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardRow.amount; }, 0);
}

// ---------------------------------------------------------------------------
// Month-over-month trend arrows (Чистый денежный поток, Чистая прибыль,
// Рентабельность по ЧП, Собственный капитал)
// ---------------------------------------------------------------------------

/**
 * Compares each of the 4 requested KPIs against the same "month-to-date"
 * window one calendar month earlier (e.g. snapshot 15.09 -> compares
 * 01.09-15.09 against 01.08-15.08; a shorter previous month is clipped to
 * its last day). Uses the FULL, period-unfiltered source rows, since the
 * previous-month window can fall before the dashboard's own period start.
 * @param {Array<Object>} dashboardCashRows All parsed "Деньги" rows.
 * @param {Array<Object>} dashboardProfitRows All parsed "Прибыль" rows.
 * @param {Array<Object>} dashboardEquityRows All parsed "Капитал" rows.
 * @param {Array<Date>} dashboardEquityDates Unique sorted equity dates.
 * @param {Date} dashboardSnapshotDate Date D.
 * @param {Object} dashboardKpis Already-computed current-period KPIs (netCashFlow, netProfit, netMargin, equity).
 * @return {Object} { netCashFlow, netProfit, netMargin, equity } each {current, previous, direction: 'up'|'down'|'flat'|null}.
 * @private
 */
function dashboardBuildTrends_(dashboardCashRows, dashboardProfitRows, dashboardEquityRows, dashboardEquityDates, dashboardSnapshotDate, dashboardKpis, dashboardPeriodStart) {
  var dashboardCurrentMonthStart = new Date(dashboardSnapshotDate.getFullYear(), dashboardSnapshotDate.getMonth(), 1);
  var dashboardCurrentWindowEnd = new Date(dashboardSnapshotDate.getTime());
  dashboardCurrentWindowEnd.setHours(23, 59, 59, 999);
  dashboardCurrentMonthStart.setHours(0, 0, 0, 0);

  var dashboardPrevMonthIndex = dashboardCurrentMonthStart.getMonth() === 0 ? 11 : dashboardCurrentMonthStart.getMonth() - 1;
  var dashboardPrevMonthYear = dashboardCurrentMonthStart.getMonth() === 0 ? dashboardCurrentMonthStart.getFullYear() - 1 : dashboardCurrentMonthStart.getFullYear();
  var dashboardPrevMonthStart = new Date(dashboardPrevMonthYear, dashboardPrevMonthIndex, 1);
  dashboardPrevMonthStart.setHours(0, 0, 0, 0);
  var dashboardPrevMonthLastDay = new Date(dashboardPrevMonthYear, dashboardPrevMonthIndex + 1, 0).getDate();
  var dashboardPrevMonthRefDay = Math.min(dashboardSnapshotDate.getDate(), dashboardPrevMonthLastDay);
  var dashboardPrevWindowEnd = new Date(dashboardPrevMonthYear, dashboardPrevMonthIndex, dashboardPrevMonthRefDay, 23, 59, 59, 999);

  Logger.log('[Dashboard][Trends] Текущий период (МТД): ' + dashboardFormatDate_(dashboardCurrentMonthStart) + ' — ' + dashboardFormatDate_(dashboardCurrentWindowEnd));
  Logger.log('[Dashboard][Trends] Предыдущий период (МТД-1мес): ' + dashboardFormatDate_(dashboardPrevMonthStart) + ' — ' + dashboardFormatDate_(dashboardPrevWindowEnd));

  var dashboardPrevCashRows = dashboardFilterPeriod_(dashboardCashRows, dashboardPrevMonthStart, dashboardPrevWindowEnd).filter(function (dashboardRow) {
    return !dashboardCashIsMovement_(dashboardRow) &&
        (dashboardIsOperation_(dashboardRow.operation, 'Приход') || dashboardIsOperation_(dashboardRow.operation, 'Расход'));
  });
  var dashboardPrevNetCashFlow = dashboardPrevCashRows.reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardCashSignedAmount_(dashboardRow); }, 0);

  var dashboardCurrentMtdCashRows = dashboardFilterPeriod_(dashboardCashRows, dashboardCurrentMonthStart, dashboardCurrentWindowEnd).filter(function (dashboardRow) {
    return !dashboardCashIsMovement_(dashboardRow) &&
        (dashboardIsOperation_(dashboardRow.operation, 'Приход') || dashboardIsOperation_(dashboardRow.operation, 'Расход'));
  });
  var dashboardCurrentMtdNetCashFlow = dashboardCurrentMtdCashRows.reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardCashSignedAmount_(dashboardRow); }, 0);

  var dashboardPrevProfitRows = dashboardFilterPeriod_(dashboardProfitRows, dashboardPrevMonthStart, dashboardPrevWindowEnd);
  var dashboardPrevRevenue = dashboardSumProfitSection_(dashboardPrevProfitRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
  var dashboardPrevNetProfit = dashboardNetProfitForRows_(dashboardPrevProfitRows);
  var dashboardPrevNetMargin = dashboardPrevRevenue ? dashboardPrevNetProfit / dashboardPrevRevenue : null;

  var dashboardCurrentMtdProfitRows = dashboardFilterPeriod_(dashboardProfitRows, dashboardCurrentMonthStart, dashboardCurrentWindowEnd);
  var dashboardCurrentMtdRevenue = dashboardSumProfitSection_(dashboardCurrentMtdProfitRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
  var dashboardCurrentMtdNetProfit = dashboardNetProfitForRows_(dashboardCurrentMtdProfitRows);
  var dashboardCurrentMtdNetMargin = dashboardCurrentMtdRevenue ? dashboardCurrentMtdNetProfit / dashboardCurrentMtdRevenue : null;

  var dashboardPrevEquityLookup = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardPrevWindowEnd, 'level2', 'Капитал');
  var dashboardPrevEquity = dashboardPrevEquityLookup.found ? dashboardPrevEquityLookup.value : null;

  Logger.log('[Dashboard][Trends] Чистый денежный поток: тек.МТД=' + dashboardCurrentMtdNetCashFlow + ', пред.МТД=' + dashboardPrevNetCashFlow);
  Logger.log('[Dashboard][Trends] Чистая прибыль: тек.МТД=' + dashboardCurrentMtdNetProfit + ', пред.МТД=' + dashboardPrevNetProfit);
  Logger.log('[Dashboard][Trends] Рентабельность по ЧП: тек.МТД=' + dashboardCurrentMtdNetMargin + ', пред.МТД=' + dashboardPrevNetMargin);
  Logger.log('[Dashboard][Trends] Собственный капитал: на дату среза=' + dashboardKpis.equity + ', на ' + (dashboardPrevEquityLookup.found ? dashboardFormatDate_(dashboardPrevEquityLookup.date) : 'н/д') + '=' + dashboardPrevEquity);

  var dashboardPreviousPeriod = dashboardPreviousEquivalentPeriod_(dashboardPeriodStart, dashboardSnapshotDate);
  var dashboardPreviousAr = dashboardTurnoverForPeriod_(dashboardProfitRows, dashboardEquityRows, dashboardEquityDates, dashboardPreviousPeriod.start, dashboardPreviousPeriod.end, 'receivables');
  var dashboardPreviousAp = dashboardTurnoverForPeriod_(dashboardProfitRows, dashboardEquityRows, dashboardEquityDates, dashboardPreviousPeriod.start, dashboardPreviousPeriod.end, 'payables');
  Logger.log('[Dashboard][Trends] Предыдущий сопоставимый период для оборачиваемости: ' + dashboardFormatDate_(dashboardPreviousPeriod.start) + ' — ' + dashboardFormatDate_(dashboardPreviousPeriod.end));

  return {
    netCashFlow: dashboardTrendEntry_(dashboardCurrentMtdNetCashFlow, dashboardPrevNetCashFlow),
    netProfit: dashboardTrendEntry_(dashboardCurrentMtdNetProfit, dashboardPrevNetProfit),
    netMargin: dashboardTrendEntry_(dashboardCurrentMtdNetMargin, dashboardPrevNetMargin),
    equity: dashboardTrendEntry_(dashboardKpis.equity, dashboardPrevEquity),
    receivablesTurnover: dashboardTrendEntry_(dashboardKpis.receivablesTurnover.days, dashboardPreviousAr ? dashboardPreviousAr.days : null),
    payablesTurnover: dashboardTrendEntry_(dashboardKpis.payablesTurnover.days, dashboardPreviousAp ? dashboardPreviousAp.days : null)
  };
}

/**
 * Returns the total ending cash balance on the last available balance date.
 * @param {Array<Object>} dashboardRows Cash rows of the selected period.
 * @param {Date} dashboardEnd End of the selected period.
 * @return {{value: number|null, date: Date|null, count: number}}
 * @private
 */
function dashboardCashEndingBalance_(dashboardRows, dashboardEnd) {
  var dashboardBalanceRows = dashboardRows.filter(function (dashboardRow) {
    return dashboardIsOperation_(dashboardRow.operation, 'Остаток') &&
        dashboardRow.date && dashboardRow.date.getTime() <= dashboardEnd.getTime();
  });
  if (!dashboardBalanceRows.length) {
    var dashboardFallback = dashboardRows.filter(function (dashboardRow) {
      return dashboardRow.date && dashboardRow.date.getTime() <= dashboardEnd.getTime();
    }).reduce(function (dashboardSum, dashboardRow) {
      return dashboardSum + dashboardCashSignedAmount_(dashboardRow);
    }, 0);
    Logger.log('[Dashboard][CashBuffer] В источнике нет отдельных строк остатка; использован накопленный остаток операций: ' + dashboardFallback);
    return { value: dashboardFallback, date: new Date(dashboardEnd.getTime()), count: 0 };
  }
  var dashboardBalanceDate = dashboardBalanceRows.reduce(function (dashboardLatest, dashboardRow) {
    return dashboardRow.date.getTime() > dashboardLatest.getTime() ? dashboardRow.date : dashboardLatest;
  }, dashboardBalanceRows[0].date);
  var dashboardRowsAtEnd = dashboardBalanceRows.filter(function (dashboardRow) {
    return dashboardRow.date.getTime() === dashboardBalanceDate.getTime();
  });
  return {
    value: dashboardRowsAtEnd.reduce(function (dashboardSum, dashboardRow) {
      return dashboardSum + dashboardCashSignedAmount_(dashboardRow);
    }, 0),
    date: new Date(dashboardBalanceDate.getTime()),
    count: dashboardRowsAtEnd.length
  };
}

/** @param {number} dashboardCurrent Current-period value. @param {number|null} dashboardPrevious Previous-month value (null when unavailable). @return {{current:number, previous:(number|null), direction:(string|null)}} @private */
function dashboardTrendEntry_(dashboardCurrent, dashboardPrevious) {
  var dashboardDirection = null;
  if (dashboardPrevious !== null && typeof dashboardPrevious !== 'undefined' && typeof dashboardCurrent === 'number') {
    if (dashboardCurrent > dashboardPrevious) { dashboardDirection = 'up'; }
    else if (dashboardCurrent < dashboardPrevious) { dashboardDirection = 'down'; }
    else { dashboardDirection = 'flat'; }
  }
  return {
    current: typeof dashboardCurrent === 'number' ? dashboardCurrent : null,
    previous: dashboardPrevious === null || typeof dashboardPrevious === 'undefined' ? null : dashboardPrevious,
    relative: dashboardPrevious === null || typeof dashboardPrevious === 'undefined' || dashboardPrevious === 0 ? null : (dashboardCurrent - dashboardPrevious) / Math.abs(dashboardPrevious),
    direction: dashboardDirection
  };
}

/** @param {Date} dashboardStart Current-period start. @param {Date} dashboardEnd Current-period end. @return {{start: Date, end: Date}} Previous period of equal calendar length. @private */
function dashboardPreviousEquivalentPeriod_(dashboardStart, dashboardEnd) {
  var dashboardDays = dashboardCalendarDays_(dashboardStart, dashboardEnd);
  var dashboardPreviousEnd = new Date(dashboardStart.getTime());
  dashboardPreviousEnd.setDate(dashboardPreviousEnd.getDate() - 1);
  dashboardPreviousEnd.setHours(23, 59, 59, 999);
  var dashboardPreviousStart = new Date(dashboardPreviousEnd.getTime());
  dashboardPreviousStart.setDate(dashboardPreviousStart.getDate() - dashboardDays + 1);
  dashboardPreviousStart.setHours(0, 0, 0, 0);
  return { start: dashboardPreviousStart, end: dashboardPreviousEnd };
}

/** @param {Array<Object>} dashboardProfitRows All profit rows. @param {Array<Object>} dashboardEquityRows All equity rows. @param {Array<Date>} dashboardEquityDates Equity snapshot dates. @param {Date} dashboardStart Start. @param {Date} dashboardEnd End. @param {'receivables'|'payables'} dashboardKind Turnover kind. @return {{ratio:number,days:number}|null} Previous-period turnover. @private */
function dashboardTurnoverForPeriod_(dashboardProfitRows, dashboardEquityRows, dashboardEquityDates, dashboardStart, dashboardEnd, dashboardKind) {
  var dashboardStartValue = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardStart, 'level3', dashboardKind === 'receivables' ? 'Дебиторская задолженность' : 'Кредиторская задолженность');
  var dashboardEndValue = dashboardFindOnOrBefore_(dashboardEquityRows, dashboardEquityDates, dashboardEnd, 'level3', dashboardKind === 'receivables' ? 'Дебиторская задолженность' : 'Кредиторская задолженность');
  if (!dashboardStartValue.found || !dashboardEndValue.found) { return null; }
  var dashboardProfitPeriodRows = dashboardFilterPeriod_(dashboardProfitRows, dashboardStart, dashboardEnd);
  var dashboardNumerator = dashboardKind === 'receivables' ?
      dashboardSumProfitSection_(dashboardProfitPeriodRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE) :
      Math.abs(dashboardSumProfitSection_(dashboardProfitPeriodRows, DASHBOARD_PROFIT_SECTION_CODES.COST));
  var dashboardAverageBalance = (dashboardStartValue.value + dashboardEndValue.value) / 2;
  if (!dashboardNumerator || !dashboardAverageBalance) { return null; }
  var dashboardRatio = dashboardNumerator / dashboardAverageBalance;
  return { ratio: dashboardRatio, days: dashboardRatio ? dashboardCalendarDays_(dashboardStart, dashboardEnd) / dashboardRatio : 0 };
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

/** @return {Array<Object>} Chart 1: monthly net cash flow. @private */
function dashboardBuildCashMonthlyNetFlow_(dashboardRows, dashboardMonths) {
  var dashboardTotals = {};
  dashboardRows.forEach(function (dashboardRow) {
    var dashboardKey = dashboardMonthKey_(dashboardRow.date);
    dashboardTotals[dashboardKey] = (dashboardTotals[dashboardKey] || 0) + dashboardCashSignedAmount_(dashboardRow);
  });
  return dashboardMonths.map(function (dashboardMonth) {
    return { label: dashboardMonth.label, value: dashboardTotals[dashboardMonth.key] || 0 };
  });
}

/** @return {{months: Array<Object>, series: Array<Object>}} Chart 2: monthly net cash flow by "Направление". @private */
function dashboardBuildCashByDirection_(dashboardRows, dashboardMonths) {
  var dashboardDirections = {};
  dashboardRows.forEach(function (dashboardRow) { if (dashboardRow.direction) { dashboardDirections[dashboardRow.direction] = true; } });
  var dashboardNames = Object.keys(dashboardDirections).sort(function (dashboardA, dashboardB) { return dashboardA.localeCompare(dashboardB, 'ru'); });
  return {
    months: dashboardMonths,
    series: dashboardNames.map(function (dashboardDirection) {
      return {
        label: dashboardDirection,
        values: dashboardMonths.map(function (dashboardMonth) {
          return dashboardRows.filter(function (dashboardRow) {
            return dashboardRow.direction === dashboardDirection && dashboardMonthKey_(dashboardRow.date) === dashboardMonth.key;
          }).reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardCashSignedAmount_(dashboardRow); }, 0);
        })
      };
    })
  };
}

/** @return {Array<Object>} Chart 3: revenue (bars) and gross-profit margin (line), per month. @private */
function dashboardBuildRevenueAndGrossMargin_(dashboardRows, dashboardMonths) {
  return dashboardMonths.map(function (dashboardMonth) {
    var dashboardMonthRows = dashboardRows.filter(function (dashboardRow) { return dashboardMonthKey_(dashboardRow.date) === dashboardMonth.key; });
    var dashboardRevenue = dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
    var dashboardGross = dashboardRevenue + dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.COST) + dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.OVERHEAD);
    return { key: dashboardMonth.key, label: dashboardMonth.label, revenue: dashboardRevenue, grossMargin: dashboardRevenue ? dashboardGross / dashboardRevenue : 0 };
  });
}

/** @return {{months: Array<Object>, series: Array<Object>}} Chart 4: gross profit by direction, stacked. @private */
function dashboardBuildProfitDirectionSeries_(dashboardRows, dashboardMonths) {
  var dashboardDirections = {};
  dashboardRows.forEach(function (dashboardRow) { if (dashboardRow.direction) { dashboardDirections[dashboardRow.direction] = true; } });
  var dashboardNames = Object.keys(dashboardDirections).sort(function (dashboardA, dashboardB) { return dashboardA.localeCompare(dashboardB, 'ru'); });
  return {
    months: dashboardMonths,
    series: dashboardNames.map(function (dashboardDirection) {
      return {
        label: dashboardDirection,
        values: dashboardMonths.map(function (dashboardMonth) {
          var dashboardMonthRows = dashboardRows.filter(function (dashboardRow) {
            return dashboardRow.direction === dashboardDirection && dashboardMonthKey_(dashboardRow.date) === dashboardMonth.key;
          });
          return dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE) +
              dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.COST) +
              dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.OVERHEAD);
        })
      };
    })
  };
}

/** @return {Array<Object>} Chart 5: net profit (bars) and net margin (line), per month. @private */
function dashboardBuildNetProfitAndMargin_(dashboardRows, dashboardMonths) {
  return dashboardMonths.map(function (dashboardMonth) {
    var dashboardMonthRows = dashboardRows.filter(function (dashboardRow) { return dashboardMonthKey_(dashboardRow.date) === dashboardMonth.key; });
    var dashboardRevenue = dashboardSumProfitSection_(dashboardMonthRows, DASHBOARD_PROFIT_SECTION_CODES.REVENUE);
    var dashboardNetProfit = dashboardNetProfitForRows_(dashboardMonthRows);
    return { key: dashboardMonth.key, label: dashboardMonth.label, netProfit: dashboardNetProfit, netMargin: dashboardRevenue ? dashboardNetProfit / dashboardRevenue : 0 };
  });
}

/** @return {Array<Object>} Charts 6/7: receivables or payables per available "Капитал" date within the period. @private */
function dashboardBuildEquitySectionSeries_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd, dashboardLevel3Target) {
  var dashboardDates = dashboardUniqueDates_(dashboardFilterPeriod_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd));
  return dashboardDates.map(function (dashboardDate) {
    var dashboardDateRows = dashboardRowsForDate_(dashboardRows, dashboardDate);
    return {
      label: dashboardFormatDate_(dashboardDate),
      date: dashboardFormatIsoDate_(dashboardDate),
      value: dashboardSumWhere_(dashboardDateRows, 'level3', dashboardLevel3Target)
    };
  });
}

/** @return {Array<Object>} Chart 8: equity vs. debt mix (100% stack) per available "Капитал" date. @private */
function dashboardBuildCapitalMix_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd) {
  var dashboardDates = dashboardUniqueDates_(dashboardFilterPeriod_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd));
  return dashboardDates.map(function (dashboardDate) {
    var dashboardDateRows = dashboardRowsForDate_(dashboardRows, dashboardDate);
    var dashboardEquity = dashboardSumWhere_(dashboardDateRows, 'level2', 'Капитал');
    var dashboardDebt = dashboardSumWhere_(dashboardDateRows, 'level2', 'Обязательства');
    var dashboardTotal = dashboardEquity + dashboardDebt;
    return {
      label: dashboardFormatDate_(dashboardDate),
      date: dashboardFormatIsoDate_(dashboardDate),
      equity: dashboardEquity,
      debt: dashboardDebt,
      equityShare: dashboardTotal ? dashboardEquity / dashboardTotal : 0,
      debtShare: dashboardTotal ? dashboardDebt / dashboardTotal : 0
    };
  });
}

/**
 * Chart 9 ("Коэффициент финансовой независимости" per the task title, computed
 * with the same formula as KPI 3 per the task's explicit instruction: current
 * ratio = Краткосрочные обязательства / Оборотные активы).
 * @return {Array<Object>} Ratio per available "Капитал" date. @private
 */
function dashboardBuildCurrentRatioTrend_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd) {
  var dashboardDates = dashboardUniqueDates_(dashboardFilterPeriod_(dashboardRows, dashboardPeriodStart, dashboardPeriodEnd));
  return dashboardDates.map(function (dashboardDate) {
    var dashboardDateRows = dashboardRowsForDate_(dashboardRows, dashboardDate);
    var dashboardCurrentAssets = dashboardSumWhere_(dashboardDateRows, 'level2', 'Оборотные активы');
    var dashboardShortTermDebt = dashboardSumWhere_(dashboardDateRows, 'level3', 'Краткосрочные обязательства');
    return {
      label: dashboardFormatDate_(dashboardDate),
      date: dashboardFormatIsoDate_(dashboardDate),
      value: dashboardCurrentAssets ? dashboardShortTermDebt / dashboardCurrentAssets : 0
    };
  });
}

// ---------------------------------------------------------------------------
// Data access — rangecash / rangeprofit / rangeequity
// ---------------------------------------------------------------------------

/** @return {{rows: Array<Object>}} Parsed rangecash rows: {date, amount, operation, section, direction}. @private */
function dashboardReadCashRows_() {
  var dashboardRange = dashboardGetRangeByName_(DASHBOARD_CASH_RANGE_NAME);
  var dashboardValues = dashboardRange.getValues();
  if (dashboardValues.length < 2) { return { rows: [] }; }
  var dashboardIndex = dashboardHeaderIndex_(dashboardValues[0], DASHBOARD_CASH_RANGE_NAME, DASHBOARD_CASH_REQUIRED_COLUMNS);
  var dashboardRows = [];
  dashboardValues.slice(1).forEach(function (dashboardValueRow) {
    if (dashboardIsEmptyRow_(dashboardValueRow)) { return; }
    var dashboardDate = dashboardParseDate_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Дата'));
    if (!dashboardDate) { return; }
    dashboardRows.push({
      date: dashboardDate,
      amount: dashboardNumber_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Сумма')),
      operation: dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Вид операции')),
      section: dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Раздел ОДДС')),
      direction: dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Направление'))
    });
  });
  return { rows: dashboardRows };
}

/** @return {{rows: Array<Object>}} Parsed rangeprofit rows: {date, amount, direction, section}. @private */
function dashboardReadProfitRows_() {
  var dashboardRange = dashboardGetRangeByName_(DASHBOARD_PROFIT_RANGE_NAME);
  var dashboardValues = dashboardRange.getValues();
  if (dashboardValues.length < 2) { return { rows: [] }; }
  var dashboardIndex = dashboardHeaderIndex_(dashboardValues[0], DASHBOARD_PROFIT_RANGE_NAME, DASHBOARD_PROFIT_REQUIRED_COLUMNS);
  var dashboardRows = [];
  dashboardValues.slice(1).forEach(function (dashboardValueRow) {
    if (dashboardIsEmptyRow_(dashboardValueRow)) { return; }
    var dashboardDate = dashboardParseDate_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Дата'));
    if (!dashboardDate) { return; }
    var dashboardOperation = dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Вид операции'));
    dashboardRows.push({
      date: dashboardDate,
      amount: dashboardProfitAmount_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Сумма'), dashboardOperation),
      direction: dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Направление')),
      section: dashboardString_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Раздел ОПиУ'))
    });
  });
  return { rows: dashboardRows };
}

/** @return {{rows: Array<Object>}} Parsed rangeequity rows: {date, level2, level3, amount}. @private */
function dashboardReadEquityRows_() {
  var dashboardRange = dashboardGetRangeByName_(DASHBOARD_EQUITY_RANGE_NAME);
  var dashboardValues = dashboardRange.getValues();
  if (dashboardValues.length < 2) { return { rows: [] }; }
  var dashboardIndex = dashboardHeaderIndex_(dashboardValues[0], DASHBOARD_EQUITY_RANGE_NAME, DASHBOARD_EQUITY_REQUIRED_COLUMNS);
  var dashboardRows = [];
  dashboardValues.slice(1).forEach(function (dashboardValueRow) {
    if (dashboardIsEmptyRow_(dashboardValueRow)) { return; }
    var dashboardDate = dashboardParseDate_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Дата'));
    if (!dashboardDate) { return; }
    dashboardRows.push({
      date: dashboardDate,
      level2: dashboardCleanText_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Раздел 2')),
      level3: dashboardCleanText_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Раздел 3')),
      amount: dashboardNumber_(dashboardValueByHeader_(dashboardValueRow, dashboardIndex, 'Сумма'))
    });
  });
  return { rows: dashboardRows };
}

/** @param {string} dashboardRangeName Named range. @return {GoogleAppsScript.Spreadsheet.Range} Range. @private */
function dashboardGetRangeByName_(dashboardRangeName) {
  var dashboardSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboardSpreadsheet) { throw new Error('Не удалось открыть таблицу, связанную со скриптом.'); }
  var dashboardRange = dashboardSpreadsheet.getRangeByName(dashboardRangeName);
  if (!dashboardRange) { throw new Error('Не найден именованный диапазон «' + dashboardRangeName + '».'); }
  return dashboardRange;
}

/** @param {Array<*>} dashboardHeaderRow Header row. @param {string} dashboardRangeName Named range (for error text). @param {Array<string>} dashboardRequiredColumns Required column names. @return {Object<string, number>} Header key -> column index. @private */
function dashboardHeaderIndex_(dashboardHeaderRow, dashboardRangeName, dashboardRequiredColumns) {
  var dashboardHeaders = dashboardHeaderRow.map(dashboardHeaderKey_);
  dashboardRequiredColumns.forEach(function (dashboardColumn) {
    if (dashboardHeaders.indexOf(dashboardHeaderKey_(dashboardColumn)) === -1) {
      throw new Error('В диапазоне «' + dashboardRangeName + '» отсутствует обязательная колонка «' + dashboardColumn + '».');
    }
  });
  var dashboardIndex = {};
  dashboardHeaders.forEach(function (dashboardHeader, dashboardI) { dashboardIndex[dashboardHeader] = dashboardI; });
  return dashboardIndex;
}

/** @param {Array<*>} dashboardRow Row. @param {Object} dashboardIndex Header index. @param {string} dashboardHeader Header name. @return {*} Cell value. @private */
function dashboardValueByHeader_(dashboardRow, dashboardIndex, dashboardHeader) {
  var dashboardI = dashboardIndex[dashboardHeaderKey_(dashboardHeader)];
  return typeof dashboardI === 'undefined' ? null : dashboardRow[dashboardI];
}

/** @param {Array<*>} dashboardRow Row. @return {boolean} Whether every cell is empty. @private */
function dashboardIsEmptyRow_(dashboardRow) {
  return dashboardRow.every(function (dashboardValue) { return dashboardValue === '' || dashboardValue === null || typeof dashboardValue === 'undefined'; });
}

// ---------------------------------------------------------------------------
// Small shared helpers (deliberately self-contained — see gsmoney.gs /
// gsprofit.gs / gscapital.gs, which each read their own range independently)
// ---------------------------------------------------------------------------

/** @return {string} Script time zone. @private */
function dashboardTimeZone_() { return Session.getScriptTimeZone() || 'Etc/GMT'; }

/** @param {*} dashboardValue Header text. @return {string} Normalized header key. @private */
function dashboardHeaderKey_(dashboardValue) { return dashboardString_(dashboardValue).toLowerCase().replace(/\s+/g, ' '); }

/** @param {*} dashboardValue Value. @return {string} Trimmed string. @private */
function dashboardString_(dashboardValue) { return dashboardValue === null || typeof dashboardValue === 'undefined' ? '' : String(dashboardValue).trim(); }

/** @param {*} dashboardValue Value. @return {string} Text with collapsed internal whitespace (guards against pasted NBSP/zero-width chars). @private */
function dashboardCleanText_(dashboardValue) {
  return String(dashboardValue === null || typeof dashboardValue === 'undefined' ? '' : dashboardValue)
      .replace(/[ ​﻿]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
}

/** @param {string} dashboardSection Раздел ОПиУ text. @return {string} Comparable key ("1.выручка"). @private */
function dashboardSectionKey_(dashboardSection) { return dashboardCleanText_(dashboardSection).toLowerCase().replace(/\s+/g, ''); }

/** @param {string} dashboardOperation Operation. @param {string} dashboardExpected Expected value. @return {boolean} Case/whitespace-insensitive match. @private */
function dashboardIsOperation_(dashboardOperation, dashboardExpected) {
  return dashboardCleanText_(dashboardOperation).toLowerCase() === dashboardCleanText_(dashboardExpected).toLowerCase();
}

/**
 * Parses a spreadsheet date the same timezone-safe way as gsprofit.gs /
 * gscapital.gs (Date cells are re-based on the script's own timezone so this
 * module and the others always agree on the calendar day), plus the
 * YYYY-MM-DD / DD.MM.YYYY string formats used by the date picker and by
 * hand-entered sheet rows. Returns null (never "today") when the value is
 * missing or unparsable.
 * @param {*} dashboardValue Source value.
 * @return {Date|null} Local calendar date or null.
 * @private
 */
function dashboardParseDate_(dashboardValue) {
  if (Object.prototype.toString.call(dashboardValue) === '[object Date]') {
    if (isNaN(dashboardValue.getTime())) { return null; }
    var dashboardIsoText = Utilities.formatDate(dashboardValue, dashboardTimeZone_(), 'yyyy-MM-dd');
    return dashboardValidatedDate_(+dashboardIsoText.substring(0, 4), +dashboardIsoText.substring(5, 7), +dashboardIsoText.substring(8, 10));
  }
  if (typeof dashboardValue !== 'string') { return null; }
  var dashboardText = dashboardValue.trim();
  if (!dashboardText) { return null; }
  var dashboardMatch = /^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/.exec(dashboardText);
  if (dashboardMatch) { return dashboardValidatedDate_(+dashboardMatch[1], +dashboardMatch[2], +dashboardMatch[3]); }
  dashboardMatch = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(dashboardText);
  return dashboardMatch ? dashboardValidatedDate_(+dashboardMatch[3], +dashboardMatch[2], +dashboardMatch[1]) : null;
}

/** @return {Date|null} Valid local date, or null for an impossible calendar date (e.g. 31.02). @private */
function dashboardValidatedDate_(dashboardYear, dashboardMonth, dashboardDay) {
  var dashboardDate = new Date(dashboardYear, dashboardMonth - 1, dashboardDay);
  return dashboardDate.getFullYear() === dashboardYear && dashboardDate.getMonth() === dashboardMonth - 1 && dashboardDate.getDate() === dashboardDay ? dashboardDate : null;
}

/** @param {*} dashboardValue Number, possibly text ("-2 231,00", "1 234,56 ₽"). @return {number} Finite number. @private */
function dashboardNumber_(dashboardValue) {
  if (typeof dashboardValue === 'number') { return isFinite(dashboardValue) ? dashboardValue : 0; }
  if (typeof dashboardValue !== 'string') { return 0; }
  var dashboardText = dashboardValue.replace(/[\s ]/g, '').replace(/₽/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var dashboardParsed = Number(dashboardText);
  return isFinite(dashboardParsed) ? dashboardParsed : 0;
}

/** @param {*} dashboardValue Raw amount. @param {string} dashboardOperation "Приход"/"Расход"/other. @return {number} Signed amount (same convention as gsprofit.gs). @private */
function dashboardProfitAmount_(dashboardValue, dashboardOperation) {
  var dashboardNumberValue;
  if (typeof dashboardValue === 'number') {
    dashboardNumberValue = dashboardValue;
  } else {
    var dashboardText = String(dashboardValue === null || typeof dashboardValue === 'undefined' ? '' : dashboardValue)
        .replace(/\s/g, '').replace(/₽/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
    dashboardNumberValue = Number(dashboardText);
  }
  if (!isFinite(dashboardNumberValue)) { return 0; }
  if (dashboardIsOperation_(dashboardOperation, 'Расход')) { return -Math.abs(dashboardNumberValue); }
  if (dashboardIsOperation_(dashboardOperation, 'Приход')) { return Math.abs(dashboardNumberValue); }
  return dashboardNumberValue;
}

/** @param {Object} dashboardRow Cash row. @return {boolean} Whether it is an internal transfer ("5. Перемещение"). @private */
function dashboardCashIsMovement_(dashboardRow) {
  var dashboardSection = dashboardRow.section.toLowerCase();
  return /^5(?:\.|\s|$)/.test(dashboardSection) || dashboardSection.indexOf('перемещ') !== -1;
}

/** @param {Object} dashboardRow Cash row. @return {string} 'operating'|'investing'|'financing'|'other'. @private */
function dashboardCashActivityType_(dashboardRow) {
  var dashboardSection = dashboardRow.section.toLowerCase();
  if (dashboardSection.indexOf('основн') !== -1) { return 'operating'; }
  if (dashboardSection.indexOf('инвести') !== -1) { return 'investing'; }
  if (dashboardSection.indexOf('финансов') !== -1) { return 'financing'; }
  return 'other';
}

/** @param {Object} dashboardRow Cash row. @return {number} Signed cash movement (Приход +, Расход -, Остаток as-is). @private */
function dashboardCashSignedAmount_(dashboardRow) {
  var dashboardOperation = dashboardRow.operation.toLowerCase();
  if (dashboardOperation === 'приход') { return Math.abs(dashboardRow.amount); }
  if (dashboardOperation === 'расход') { return -Math.abs(dashboardRow.amount); }
  if (dashboardOperation === 'остаток') { return dashboardRow.amount; }
  return 0;
}

/** @param {Array<Object>} dashboardRows Profit rows. @param {string} dashboardCode Normalized section code. @return {number} Sum of .amount for that section. @private */
function dashboardSumProfitSection_(dashboardRows, dashboardCode) {
  return dashboardRows.filter(function (dashboardRow) { return dashboardSectionKey_(dashboardRow.section) === dashboardCode; })
      .reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardRow.amount; }, 0);
}

/** @param {Array<Object>} dashboardRows Equity rows. @param {string} dashboardField 'level2' or 'level3'. @param {string} dashboardTarget Target text. @return {number} Sum of .amount matching the field. @private */
function dashboardSumWhere_(dashboardRows, dashboardField, dashboardTarget) {
  var dashboardKey = dashboardCleanText_(dashboardTarget).toLowerCase();
  return dashboardRows.filter(function (dashboardRow) { return dashboardCleanText_(dashboardRow[dashboardField]).toLowerCase() === dashboardKey; })
      .reduce(function (dashboardSum, dashboardRow) { return dashboardSum + dashboardRow.amount; }, 0);
}

/** @param {Array<Object>} dashboardRows Any parsed rows with a .date field. @return {Array<Date>} Unique dates, ascending. @private */
function dashboardUniqueDates_(dashboardRows) {
  var dashboardMap = {};
  dashboardRows.forEach(function (dashboardRow) { dashboardMap[dashboardRow.date.getTime()] = dashboardRow.date; });
  return Object.keys(dashboardMap).map(function (dashboardKey) { return dashboardMap[dashboardKey]; })
      .sort(function (dashboardA, dashboardB) { return dashboardA - dashboardB; });
}

/** @param {Array<Object>} dashboardRows Rows. @param {Date} dashboardDate Target date. @return {Array<Object>} Rows on that exact date. @private */
function dashboardRowsForDate_(dashboardRows, dashboardDate) {
  return dashboardRows.filter(function (dashboardRow) { return dashboardRow.date.getTime() === dashboardDate.getTime(); });
}

/** @param {Array<Object>} dashboardRows Rows. @param {Date} dashboardStart Start (inclusive). @param {Date} dashboardEnd End (inclusive). @return {Array<Object>} Rows within [start, end]. @private */
function dashboardFilterPeriod_(dashboardRows, dashboardStart, dashboardEnd) {
  return dashboardRows.filter(function (dashboardRow) {
    return dashboardRow.date && dashboardRow.date.getTime() >= dashboardStart.getTime() && dashboardRow.date.getTime() <= dashboardEnd.getTime();
  });
}

/** @param {Array<Date>} dashboardSortedDates Ascending unique dates. @param {Date} dashboardTarget Target date. @return {Date|null} Latest date <= target, or null. @private */
function dashboardFindOnOrBeforeDate_(dashboardSortedDates, dashboardTarget) {
  var dashboardCandidate = null;
  for (var dashboardI = 0; dashboardI < dashboardSortedDates.length; dashboardI++) {
    if (dashboardSortedDates[dashboardI].getTime() <= dashboardTarget.getTime()) {
      dashboardCandidate = dashboardSortedDates[dashboardI];
    } else {
      break;
    }
  }
  return dashboardCandidate;
}

/**
 * Finds the latest equity date on or before dashboardTarget and sums a field
 * on that date. Used for "capital at period start" (S may not be an exact
 * snapshot date — spec: take the nearest earlier one).
 * @return {{found: boolean, date: Date|null, value: number}}
 * @private
 */
function dashboardFindOnOrBefore_(dashboardRows, dashboardSortedDates, dashboardTarget, dashboardField, dashboardValueTarget) {
  var dashboardCandidate = dashboardFindOnOrBeforeDate_(dashboardSortedDates, dashboardTarget);
  if (!dashboardCandidate) { return { found: false, date: null, value: 0 }; }
  var dashboardDateRows = dashboardRowsForDate_(dashboardRows, dashboardCandidate);
  return { found: true, date: dashboardCandidate, value: dashboardSumWhere_(dashboardDateRows, dashboardField, dashboardValueTarget) };
}

/** @param {Date} dashboardStart Start. @param {Date} dashboardEnd End. @return {number} Inclusive calendar-day count. @private */
function dashboardCalendarDays_(dashboardStart, dashboardEnd) {
  var dashboardStartDay = new Date(dashboardStart.getFullYear(), dashboardStart.getMonth(), dashboardStart.getDate());
  var dashboardEndDay = new Date(dashboardEnd.getFullYear(), dashboardEnd.getMonth(), dashboardEnd.getDate());
  return Math.floor((dashboardEndDay - dashboardStartDay) / 86400000) + 1;
}

/** @param {Date} dashboardStart Start. @param {Date} dashboardEnd End. @return {Array<Object>} Calendar months spanning [start, end]. @private */
function dashboardBuildMonths_(dashboardStart, dashboardEnd) {
  var dashboardMonths = [];
  var dashboardCursor = new Date(dashboardStart.getFullYear(), dashboardStart.getMonth(), 1);
  var dashboardLast = new Date(dashboardEnd.getFullYear(), dashboardEnd.getMonth(), 1);
  while (dashboardCursor.getTime() <= dashboardLast.getTime()) {
    dashboardMonths.push({
      key: dashboardMonthKey_(dashboardCursor),
      label: DASHBOARD_MONTH_NAMES[dashboardCursor.getMonth()] + ' ' + dashboardCursor.getFullYear(),
      date: new Date(dashboardCursor.getTime())
    });
    dashboardCursor.setMonth(dashboardCursor.getMonth() + 1);
  }
  return dashboardMonths;
}

/** @param {Date} dashboardDate Date. @return {string} "YYYY-MM" key. @private */
function dashboardMonthKey_(dashboardDate) { return dashboardDate.getFullYear() + '-' + ('0' + (dashboardDate.getMonth() + 1)).slice(-2); }

/** @param {Date} dashboardDate Date. @return {string} "YYYY-MM-DD". @private */
function dashboardFormatIsoDate_(dashboardDate) {
  return dashboardDate.getFullYear() + '-' + ('0' + (dashboardDate.getMonth() + 1)).slice(-2) + '-' + ('0' + dashboardDate.getDate()).slice(-2);
}

/** @param {Date} dashboardDate Date. @return {string} "DD.MM.YYYY" for messages and chart labels. @private */
function dashboardFormatDate_(dashboardDate) { return Utilities.formatDate(dashboardDate, dashboardTimeZone_(), 'dd.MM.yyyy'); }

// ---------------------------------------------------------------------------
// Response helpers (mirrors gsmoney.gs / gsprofit.gs / gscapital.gs)
// ---------------------------------------------------------------------------

/** @param {*} dashboardData Payload. @return {Object} Controlled success response. @private */
function dashboardSuccessResponse_(dashboardData) {
  var dashboardResponse = typeof createSuccessResponse === 'function' ? createSuccessResponse(dashboardData) : { status: 'success', data: dashboardData, error: null };
  return typeof prepareResponse === 'function' ? prepareResponse(dashboardResponse) : dashboardResponse;
}

/** @param {*} dashboardMessage Error text. @return {Object} Controlled error response. @private */
function dashboardErrorResponse_(dashboardMessage) {
  var dashboardResponse = typeof createErrorResponse === 'function' ? createErrorResponse(dashboardMessage) : { status: 'error', data: null, error: String(dashboardMessage) };
  return typeof prepareResponse === 'function' ? prepareResponse(dashboardResponse) : dashboardResponse;
}

/** @param {*} dashboardMessage Log payload. @private */
function dashboardLogError_(dashboardMessage) {
  if (typeof logError === 'function') { logError(dashboardMessage); } else { Logger.log(String(dashboardMessage)); }
}

/** @param {Object} dashboardRow Cash row. @return {boolean} Whether the row belongs to «2. Основная деятельность». @private */
function dashboardCashIsOperatingActivity_(dashboardRow) {
  var dashboardSection = dashboardSectionKey_(dashboardRow.section);
  return /^2(?:\.|$)/.test(dashboardSection) || dashboardSection.indexOf('основн') !== -1;
}

/** @param {Array<Array<Object>>} dashboardSources Source row arrays. @param {Date} dashboardEnd Latest permitted date. @return {Date|null} Earliest actual source date on or before end. @private */
function dashboardEarliestSourceDate_(dashboardSources, dashboardEnd) {
  var dashboardEarliest = null;
  dashboardSources.forEach(function (dashboardRows) {
    dashboardRows.forEach(function (dashboardRow) {
      if (dashboardRow.date && dashboardRow.date.getTime() <= dashboardEnd.getTime() &&
          (!dashboardEarliest || dashboardRow.date.getTime() < dashboardEarliest.getTime())) {
        dashboardEarliest = dashboardRow.date;
      }
    });
  });
  return dashboardEarliest;
}
