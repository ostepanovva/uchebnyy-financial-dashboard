/** @OnlyCurrentDoc */
/**
 * Profit module for the financial dashboard.
 * Reads only the named range "rangeprofit" and only the declared columns.
 * All functions in this file start with "profit" to avoid module conflicts.
 */

var PROFIT_RANGE_NAME = 'rangeprofit';
var PROFIT_REQUIRED_COLUMNS = [
  'Дата', 'Сумма', 'Направление', 'Наименование проекта',
  'Вид операции', 'Статья ОПиУ', 'Раздел ОПиУ'
];
var PROFIT_SECTION_CODES = {
  REVENUE: '1.выручка',
  COST: '2.себестоимость',
  OVERHEAD: '3.общепроизводственные',
  COMMERCIAL: '4.коммерческиерасходы',
  ADMINISTRATIVE: '5.административныерасходы',
  OTHER_INCOME: '6.доходынижеоперационнойприбыли',
  OTHER_EXPENSES: '7.расходынижеоперационнойприбыли'
};

/**
 * Public API called from profit.html through google.script.run.
 * @param {Object=} params Date period and page filters.
 * @return {Object} Unified response: {status, data, error}.
 */
function profitGetProfitData(params) {
  try {
    var profitRequest = profitNormalizeRequest_(params);
    var profitSource = profitReadNamedRange_();
    if (!profitSource.rows.length) {
      return profitSuccessResponse_({
        kpis: profitEmptyKpis_(),
        charts: profitEmptyCharts_(),
        tables: { construction: [], repair: [] },
        filters: profitBuildFilterOptions_([]),
        meta: { rangeName: PROFIT_RANGE_NAME, rows: 0, message: 'В rangeprofit нет строк с валидной датой.' }
      });
    }

    var profitPeriod = profitResolvePeriod_(profitRequest, profitSource.rows);
    var profitPeriodRows = profitFilterByPeriod_(profitSource.rows, profitPeriod.startDate, profitPeriod.endDate);
    var profitRows = profitFilterRows_(profitPeriodRows, profitRequest.filters);
    var profitMonths = profitBuildMonths_(profitPeriod.startDate, profitPeriod.endDate);
    var profitMonthly = profitBuildMonthlyMetrics_(profitRows, profitMonths);

    return profitSuccessResponse_({
      kpis: profitBuildKpis_(profitRows),
      charts: {
        revenueAndGrossMargin: profitMonthly,
        revenueByDirection: profitBuildDirectionSeries_(profitRows, profitMonths, 'revenue'),
        grossProfitByDirection: profitBuildDirectionSeries_(profitRows, profitMonths, 'grossProfit'),
        netProfitAndMargin: profitMonthly.map(function (profitItem) {
          return {
            key: profitItem.key,
            label: profitItem.label,
            netProfit: profitItem.netProfit,
            netMargin: profitItem.netMargin
          };
        })
      },
      tables: {
        construction: profitBuildProjectTable_(profitRows, 'Стройка'),
        repair: profitBuildProjectTable_(profitRows, 'Ремонт')
      },
      filters: profitBuildFilterOptions_(profitSource.rows),
      meta: {
        rangeName: PROFIT_RANGE_NAME,
        rows: profitRows.length,
        sourceRows: profitSource.rows.length,
        period: {
          startDate: profitFormatDate_(profitPeriod.startDate),
          endDate: profitFormatDate_(profitPeriod.endDate)
        }
      }
    });
  } catch (profitError) {
    profitLogError_('profitGetProfitData: ' + profitError.message);
    return profitErrorResponse_(profitError.message || 'Не удалось подготовить данные по прибыли.');
  }
}

/** @param {Object=} params Raw request. @return {Object} Normalized request. @private */
function profitNormalizeRequest_(params) {
  var profitParams = params && typeof params === 'object' ? params : {};
  var profitFilters = profitParams.filters && typeof profitParams.filters === 'object' ? profitParams.filters : {};
  return {
    startDate: profitParseLocalDate_(profitParams.startDate),
    endDate: profitParseLocalDate_(profitParams.endDate),
    filters: {
      projects: profitStringArray_(profitFilters.projects),
      directions: profitStringArray_(profitFilters.directions),
      articles: profitStringArray_(profitFilters.articles),
      sections: profitStringArray_(profitFilters.sections)
    }
  };
}

/** @param {*} value Value. @return {Array<string>} String array. @private */
function profitStringArray_(value) {
  return (Array.isArray(value) ? value : []).map(function (profitItem) {
    return String(profitItem === null || profitItem === undefined ? '' : profitItem).trim();
  }).filter(function (profitItem) { return profitItem !== ''; });
}

/** @return {Object} Named-range rows and source metadata. @private */
function profitReadNamedRange_() {
  var profitSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var profitRange = profitSpreadsheet.getRangeByName(PROFIT_RANGE_NAME);
  if (!profitRange) {
    throw new Error('Не найден именованный диапазон "' + PROFIT_RANGE_NAME + '".');
  }
  var profitValues = profitRange.getValues();
  if (!profitValues.length) {
    return { rows: [] };
  }
  var profitHeaders = profitValues[0].map(function (profitHeader) { return String(profitHeader).trim(); });
  var profitIndexes = profitHeaderIndexes_(profitHeaders);
  var profitRows = [];

  profitValues.slice(1).forEach(function (profitValueRow, profitIndex) {
    if (profitValueRow.every(function (profitCell) { return profitCell === '' || profitCell === null; })) {
      return;
    }
    var profitDate = profitParseLocalDate_(profitValueRow[profitIndexes['Дата']]);
    if (!profitDate) {
      profitLogInfo_('Пропущена строка с невалидной датой: ' + (profitRange.getRow() + profitIndex + 1));
      return;
    }
    var profitOperation = profitCleanText_(profitValueRow[profitIndexes['Вид операции']]);
    profitRows.push({
      date: profitDate,
      amount: profitNormalizeAmount_(profitValueRow[profitIndexes['Сумма']], profitOperation),
      direction: profitCleanText_(profitValueRow[profitIndexes['Направление']]),
      project: profitCleanText_(profitValueRow[profitIndexes['Наименование проекта']]),
      operation: profitOperation,
      article: profitCleanText_(profitValueRow[profitIndexes['Статья ОПиУ']]),
      section: profitCleanText_(profitValueRow[profitIndexes['Раздел ОПиУ']]),
      sourceRow: profitRange.getRow() + profitIndex + 1
    });
  });
  return { rows: profitRows };
}

/** @param {Array<string>} headers Header row. @return {Object<string, number>} Header indexes. @private */
function profitHeaderIndexes_(headers) {
  var profitIndexes = {};
  PROFIT_REQUIRED_COLUMNS.forEach(function (profitHeader) {
    var profitIndex = headers.indexOf(profitHeader);
    if (profitIndex === -1) {
      throw new Error('В rangeprofit отсутствует обязательная колонка «' + profitHeader + '».');
    }
    profitIndexes[profitHeader] = profitIndex;
  });
  return profitIndexes;
}

/**
 * Parses a spreadsheet date or text date without UTC day shifting.
 * For ISO input the date part is used directly: 2023-01-01T00:00:00Z stays 01.01.2023.
 * @param {*} value Source date.
 * @return {Date|null} Local calendar-day Date or null.
 * @private
 */
function profitParseLocalDate_(value) {
  try {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      var profitSheetDate = Utilities.formatDate(value, profitTimeZone_(), 'yyyy-MM-dd');
      return profitDateFromParts_(profitSheetDate.substring(0, 4), profitSheetDate.substring(5, 7), profitSheetDate.substring(8, 10));
    }
    if (typeof value !== 'string') {
      return null;
    }
    var profitText = value.trim();
    var profitMatch = profitText.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (profitMatch) {
      return profitDateFromParts_(profitMatch[3], profitMatch[2], profitMatch[1]);
    }
    profitMatch = profitText.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
    if (profitMatch) {
      return profitDateFromParts_(profitMatch[1], profitMatch[2], profitMatch[3]);
    }
    return null;
  } catch (profitError) {
    profitLogError_('profitParseLocalDate: ' + profitError.message);
    return null;
  }
}

/** @param {string|number} year Year. @param {string|number} month Month. @param {string|number} day Day. @return {Date|null} Local Date. @private */
function profitDateFromParts_(year, month, day) {
  var profitYear = Number(year);
  var profitMonth = Number(month);
  var profitDay = Number(day);
  var profitDate = new Date(profitYear, profitMonth - 1, profitDay, 0, 0, 0, 0);
  if (profitDate.getFullYear() !== profitYear || profitDate.getMonth() !== profitMonth - 1 || profitDate.getDate() !== profitDay) {
    return null;
  }
  return profitDate;
}

/** @return {string} Server script time zone. @private */
function profitTimeZone_() {
  return Session.getScriptTimeZone() || 'Etc/GMT';
}

/** @param {*} value Raw amount. @param {string} operation Operation. @return {number} Normalized amount. @private */
function profitNormalizeAmount_(value, operation) {
  var profitNumber;
  if (typeof value === 'number') {
    profitNumber = value;
  } else {
    var profitText = String(value === null || value === undefined ? '' : value)
      .replace(/\s/g, '').replace(/₽/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
    profitNumber = Number(profitText);
  }
  if (!isFinite(profitNumber)) {
    return 0;
  }
  if (profitOperationIs_(operation, 'Расход')) {
    return -Math.abs(profitNumber);
  }
  if (profitOperationIs_(operation, 'Приход')) {
    return Math.abs(profitNumber);
  }
  return profitNumber;
}

/** @param {string} operation Operation. @param {string} expected Expected operation. @return {boolean} Match. @private */
function profitOperationIs_(operation, expected) {
  return profitCleanText_(operation).toLowerCase() === profitCleanText_(expected).toLowerCase();
}

/**
 * Produces a stable value for comparisons and group keys. Google Sheets can
 * contain non-breaking or zero-width spaces after copy/paste; they must not
 * make the same project or direction look like a different value.
 *
 * @param {*} value Value.
 * @return {string} Normalized text.
 * @private
 */
function profitCleanText_(value) {
  return String(value === null || value === undefined ? '' : value)
      .replace(/[\u00A0\u200B\uFEFF]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
}

/** @param {Object} request Request. @param {Array<Object>} rows Rows. @return {Object} Period. @private */
function profitResolvePeriod_(request, rows) {
  var profitToday = new Date();
  var profitStart = request.startDate || new Date(profitToday.getFullYear(), profitToday.getMonth(), 1);
  var profitEnd = request.endDate || new Date(profitToday.getFullYear(), profitToday.getMonth() + 1, 0);
  profitStart.setHours(0, 0, 0, 0);
  profitEnd.setHours(23, 59, 59, 999);
  if (profitStart.getTime() > profitEnd.getTime()) {
    throw new Error('Дата начала периода не может быть больше даты окончания.');
  }
  return { startDate: profitStart, endDate: profitEnd };
}

/** @param {Array<Object>} rows Rows. @param {Date} start Start. @param {Date} end End. @return {Array<Object>} Filtered rows. @private */
function profitFilterByPeriod_(rows, start, end) {
  var profitStart = new Date(start.getTime());
  var profitEnd = new Date(end.getTime());
  profitStart.setHours(0, 0, 0, 0);
  profitEnd.setHours(23, 59, 59, 999);
  return rows.filter(function (profitRow) {
    return profitRow.date.getTime() >= profitStart.getTime() && profitRow.date.getTime() <= profitEnd.getTime();
  });
}

/** @param {Array<Object>} rows Rows. @param {Object} filters Page filters. @return {Array<Object>} Filtered rows. @private */
function profitFilterRows_(rows, filters) {
  return rows.filter(function (profitRow) {
    return profitMatchesFilter_(profitRow.project, filters.projects) &&
      profitMatchesFilter_(profitRow.direction, filters.directions) &&
      profitMatchesFilter_(profitRow.article, filters.articles) &&
      profitMatchesFilter_(profitRow.section, filters.sections);
  });
}

/** @param {string} value Value. @param {Array<string>} selected Selected filters. @return {boolean} Match. @private */
function profitMatchesFilter_(value, selected) { return !selected.length || selected.indexOf(value) !== -1; }

/** @param {Date} start Start. @param {Date} end End. @return {Array<Object>} Months. @private */
function profitBuildMonths_(start, end) {
  var profitCursor = new Date(start.getFullYear(), start.getMonth(), 1);
  var profitLast = new Date(end.getFullYear(), end.getMonth(), 1);
  var profitMonths = [];
  while (profitCursor.getTime() <= profitLast.getTime()) {
    profitMonths.push({
      key: profitMonthKey_(profitCursor),
      label: Utilities.formatDate(profitCursor, profitTimeZone_(), 'MMM yyyy').replace('.', '')
    });
    profitCursor = new Date(profitCursor.getFullYear(), profitCursor.getMonth() + 1, 1);
  }
  return profitMonths;
}

/** @param {Date} date Date. @return {string} YYYY-MM key. @private */
function profitMonthKey_(date) { return date.getFullYear() + '-' + ('0' + (date.getMonth() + 1)).slice(-2); }

/** @param {Array<Object>} rows Rows. @return {Object} KPIs. @private */
function profitBuildKpis_(rows) {
  var profitRevenue = profitSumSection_(rows, PROFIT_SECTION_CODES.REVENUE);
  var profitCost = profitSumSection_(rows, PROFIT_SECTION_CODES.COST);
  var profitOverhead = profitSumSection_(rows, PROFIT_SECTION_CODES.OVERHEAD);
  var profitGross = profitRevenue + profitCost + profitOverhead;
  var profitAllExpenses = rows.filter(function (profitRow) { return profitOperationIs_(profitRow.operation, 'Расход'); })
    .reduce(function (profitSum, profitRow) { return profitSum + Math.abs(profitRow.amount); }, 0);
  var profitFixedExpenses = rows.filter(function (profitRow) {
    return profitOperationIs_(profitRow.operation, 'Расход') &&
      (profitSectionKey_(profitRow.section) === PROFIT_SECTION_CODES.COMMERCIAL || profitSectionKey_(profitRow.section) === PROFIT_SECTION_CODES.ADMINISTRATIVE);
  }).reduce(function (profitSum, profitRow) { return profitSum + Math.abs(profitRow.amount); }, 0);
  var profitNet = profitNetProfit_(rows);
  return {
    revenue: profitRevenue,
    cost: profitCost,
    costShare: profitRatio_(Math.abs(profitCost), profitRevenue),
    grossMargin: profitRatio_(profitGross, profitRevenue),
    fixedCostShare: profitRatio_(profitFixedExpenses, profitAllExpenses),
    netProfit: profitNet,
    netMargin: profitRatio_(profitNet, profitRevenue)
  };
}

/** @return {Object} Empty KPI object. @private */
function profitEmptyKpis_() { return { revenue: 0, cost: 0, costShare: 0, grossMargin: 0, fixedCostShare: 0, netProfit: 0, netMargin: 0 }; }
/** @return {Object} Empty chart object. @private */
function profitEmptyCharts_() { return { revenueAndGrossMargin: [], revenueByDirection: { months: [], series: [] }, grossProfitByDirection: { months: [], series: [] }, netProfitAndMargin: [] }; }

/** @param {Array<Object>} rows Rows. @param {string} sectionCode Normalized code. @return {number} Sum. @private */
function profitSumSection_(rows, sectionCode) {
  return rows.filter(function (profitRow) { return profitSectionKey_(profitRow.section) === sectionCode; })
    .reduce(function (profitSum, profitRow) { return profitSum + profitRow.amount; }, 0);
}

/** @param {string} section Section. @return {string} Normalized section key. @private */
function profitSectionKey_(section) { return profitCleanText_(section).toLowerCase().replace(/\s+/g, ''); }
/** @param {number} dividend Dividend. @param {number} divisor Divisor. @return {number} Ratio. @private */
function profitRatio_(dividend, divisor) { return divisor ? dividend / divisor : 0; }

/** @param {Array<Object>} rows Rows. @return {number} Net profit. @private */
function profitNetProfit_(rows) {
  var profitAllowed = {};
  Object.keys(PROFIT_SECTION_CODES).forEach(function (profitKey) { profitAllowed[PROFIT_SECTION_CODES[profitKey]] = true; });
  return rows.filter(function (profitRow) { return profitAllowed[profitSectionKey_(profitRow.section)]; })
    .reduce(function (profitSum, profitRow) { return profitSum + profitRow.amount; }, 0);
}

/** @param {Array<Object>} rows Rows. @param {Array<Object>} months Months. @return {Array<Object>} Monthly metrics. @private */
function profitBuildMonthlyMetrics_(rows, months) {
  return months.map(function (profitMonth) {
    var profitMonthRows = rows.filter(function (profitRow) { return profitMonthKey_(profitRow.date) === profitMonth.key; });
    var profitRevenue = profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.REVENUE);
    var profitGross = profitRevenue + profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.COST) + profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.OVERHEAD);
    var profitNet = profitNetProfit_(profitMonthRows);
    return {
      key: profitMonth.key,
      label: profitMonth.label,
      revenue: profitRevenue,
      grossMargin: profitRatio_(profitGross, profitRevenue),
      netProfit: profitNet,
      netMargin: profitRatio_(profitNet, profitRevenue)
    };
  });
}

/** @param {Array<Object>} rows Rows. @param {Array<Object>} months Months. @param {string} metric Metric. @return {Object} Stacked-series data. @private */
function profitBuildDirectionSeries_(rows, months, metric) {
  var profitDirections = {};
  rows.forEach(function (profitRow) { if (profitRow.direction) { profitDirections[profitRow.direction] = true; } });
  var profitNames = Object.keys(profitDirections).sort(function (profitA, profitB) { return profitA.localeCompare(profitB, 'ru'); });
  return {
    months: months,
    series: profitNames.map(function (profitDirection) {
      return {
        label: profitDirection,
        values: months.map(function (profitMonth) {
          var profitMonthRows = rows.filter(function (profitRow) { return profitRow.direction === profitDirection && profitMonthKey_(profitRow.date) === profitMonth.key; });
          if (metric === 'revenue') {
            return profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.REVENUE);
          }
          return profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.REVENUE) + profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.COST) + profitSumSection_(profitMonthRows, PROFIT_SECTION_CODES.OVERHEAD);
        })
      };
    })
  };
}

/** @param {Array<Object>} rows Rows. @param {string} targetDirection Direction. @return {Array<Object>} Project table rows. @private */
function profitBuildProjectTable_(rows, targetDirection) {
  var profitProjects = {};
  rows.filter(function (profitRow) {
    return profitDirectionIs_(profitRow.direction, targetDirection) && profitRow.project;
  }).forEach(function (profitRow) {
    var profitSection = profitSectionKey_(profitRow.section);
    if (profitSection !== PROFIT_SECTION_CODES.REVENUE && profitSection !== PROFIT_SECTION_CODES.COST && profitSection !== PROFIT_SECTION_CODES.OVERHEAD) {
      return;
    }
    if (!profitProjects[profitRow.project]) {
      profitProjects[profitRow.project] = { revenue: 0, grossProfit: 0 };
    }
    if (profitSection === PROFIT_SECTION_CODES.REVENUE) {
      profitProjects[profitRow.project].revenue += profitRow.amount;
    }
    profitProjects[profitRow.project].grossProfit += profitRow.amount;
  });
  return Object.keys(profitProjects).map(function (profitProject) {
    var profitItem = profitProjects[profitProject];
    return {
      project: profitProject,
      revenue: profitItem.revenue,
      grossProfit: profitItem.grossProfit,
      margin: profitRatio_(profitItem.grossProfit, profitItem.revenue)
    };
  }).sort(function (profitA, profitB) { return profitB.margin - profitA.margin || profitA.project.localeCompare(profitB.project, 'ru'); });
}

/** @param {string} direction Direction. @param {string} expected Expected direction. @return {boolean} Match. @private */
function profitDirectionIs_(direction, expected) { return profitCleanText_(direction).toLowerCase() === profitCleanText_(expected).toLowerCase(); }

/** @param {Array<Object>} rows Rows. @return {Object} Filter options. @private */
function profitBuildFilterOptions_(rows) {
  return {
    projects: profitUniqueValues_(rows, 'project'),
    directions: profitUniqueValues_(rows, 'direction'),
    articles: profitUniqueValues_(rows, 'article'),
    sections: profitUniqueValues_(rows, 'section')
  };
}

/** @param {Array<Object>} rows Rows. @param {string} key Object key. @return {Array<string>} Values. @private */
function profitUniqueValues_(rows, key) {
  var profitSet = {};
  rows.forEach(function (profitRow) { if (profitRow[key]) { profitSet[profitRow[key]] = true; } });
  return Object.keys(profitSet).sort(function (profitA, profitB) { return profitA.localeCompare(profitB, 'ru'); });
}

/** @param {Date} date Date. @return {string} DD.MM.YYYY. @private */
function profitFormatDate_(date) { return Utilities.formatDate(date, profitTimeZone_(), 'dd.MM.yyyy'); }

/** @param {*} data API data. @return {Object} Success response. @private */
function profitSuccessResponse_(data) {
  var profitResponse = typeof createSuccessResponse === 'function' ? createSuccessResponse(data) : { status: 'success', data: data, error: null };
  return typeof prepareResponse === 'function' ? prepareResponse(profitResponse) : profitResponse;
}
/** @param {string} message Error text. @return {Object} Error response. @private */
function profitErrorResponse_(message) {
  var profitResponse = typeof createErrorResponse === 'function' ? createErrorResponse(message) : { status: 'error', data: null, error: message };
  return typeof prepareResponse === 'function' ? prepareResponse(profitResponse) : profitResponse;
}
/** @param {*} message Debug message. @private */
function profitLogDebug_(message) { if (typeof logDebug === 'function') { logDebug(message); } else { console.log('[PROFIT][DEBUG] ' + message); } }
/** @param {*} message Info message. @private */
function profitLogInfo_(message) { if (typeof logInfo === 'function') { logInfo(message); } else { console.log('[PROFIT][INFO] ' + message); } }
/** @param {*} message Error message. @private */
function profitLogError_(message) { if (typeof logError === 'function') { logError(message); } else { console.error('[PROFIT][ERROR] ' + message); } }
