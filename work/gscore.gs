/**
 * Shared configuration, data utilities and API contracts.
 * This file is intentionally independent from the implementations of the
 * dashboard, money, capital and profit modules.
 */

var CONFIG = Object.freeze({
  SHEETS: Object.freeze({
    MONEY: 'Деньги',
    CAPITAL: 'Капитал',
    PROFIT: 'Прибыль'
  }),
  DATE_FORMAT: 'dd.MM.yyyy',
  DISPLAY_DATE_FORMAT: 'DD.MM.YYYY',
  CURRENCY_SYMBOL: '₽',
  DECIMAL_PLACES: 2,
  THOUSANDS_SEPARATOR: ' ',
  PERIODS: Object.freeze({
    currentMonth: {
      name: 'Текущий месяц',
      getDates: function () {
        var now = new Date();
        return createPeriodDates_(now.getFullYear(), now.getMonth(), now.getFullYear(), now.getMonth() + 1, 0);
      }
    },
    lastMonth: {
      name: 'Прошлый месяц',
      getDates: function () {
        var now = new Date();
        return createPeriodDates_(now.getFullYear(), now.getMonth() - 1, now.getFullYear(), now.getMonth(), 0);
      }
    },
    currentQuarter: {
      name: 'Текущий квартал',
      getDates: function () {
        var now = new Date();
        var quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
        return createPeriodDates_(now.getFullYear(), quarterStartMonth, now.getFullYear(), quarterStartMonth + 3, 0);
      }
    },
    currentYear: {
      name: 'Текущий год',
      getDates: function () {
        var now = new Date();
        return createPeriodDates_(now.getFullYear(), 0, now.getFullYear() + 1, 0, 0);
      }
    }
  })
});

/**
 * Creates the date boundaries for a period. The end date is the last
 * millisecond of the previous month when endDay is zero.
 *
 * @param {number} startYear Start year.
 * @param {number} startMonth Zero-based start month.
 * @param {number} endYear End year.
 * @param {number} endMonth Zero-based month after the period.
 * @param {number} endDay Day offset for Date constructor.
 * @return {{startDate: Date, endDate: Date}} Normalized date range.
 * @private
 */
function createPeriodDates_(startYear, startMonth, endYear, endMonth, endDay) {
  var startDate = new Date(startYear, startMonth, 1);
  var endDate = new Date(endYear, endMonth, endDay);
  endDate.setHours(23, 59, 59, 999);
  return { startDate: startDate, endDate: endDate };
}

/**
 * Creates the standard successful response returned by module APIs.
 *
 * @param {*} data Response payload.
 * @return {{status: string, data: *, error: null}} Serializable response.
 */
function createSuccessResponse(data) {
  return prepareResponse({
    status: 'success',
    data: data,
    error: null
  });
}

/**
 * Creates the standard controlled error response returned by module APIs.
 *
 * @param {*} message Human-readable error description.
 * @return {{status: string, data: null, error: string}} Serializable response.
 */
function createErrorResponse(message) {
  return prepareResponse({
    status: 'error',
    data: null,
    error: getErrorMessage_(message)
  });
}

/**
 * Returns a sheet from the spreadsheet bound to this script.
 * Returns null when the active spreadsheet or requested sheet is unavailable.
 *
 * @param {string} sheetName Exact spreadsheet sheet name.
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null} Sheet or null.
 */
function getSheetByName(sheetName) {
  try {
    if (typeof sheetName !== 'string' || !sheetName.trim()) {
      logError('getSheetByName: sheetName must be a non-empty string.');
      return null;
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      logError('getSheetByName: active spreadsheet is unavailable.');
      return null;
    }

    return spreadsheet.getSheetByName(sheetName) || null;
  } catch (error) {
    logError(error);
    return null;
  }
}

/**
 * Reads a sheet whose first row contains column names into generic objects.
 * Blank data rows are omitted. Values are returned as provided by getValues(),
 * so spreadsheet dates remain Date instances and numbers remain numbers.
 *
 * @param {string} sheetName Exact spreadsheet sheet name.
 * @return {Array<Object>} Rows mapped to their header names, or an empty array
 *     when the sheet is missing, empty, or cannot be read.
 */
function readSheetAsObjects(sheetName) {
  try {
    var sheet = getSheetByName(sheetName);
    if (!sheet) {
      logError('readSheetAsObjects: sheet not found: ' + sheetName);
      return [];
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    var headers = values[0].map(function (header) {
      return header === null || typeof header === 'undefined' ? '' : String(header).trim();
    });

    return values.slice(1).reduce(function (rows, row) {
      if (isEmptyRow_(row)) {
        return rows;
      }

      var item = {};
      headers.forEach(function (header, index) {
        // A blank header has no reliable property name and is intentionally skipped.
        if (header) {
          item[header] = normalizeValue(row[index]);
        }
      });
      rows.push(item);
      return rows;
    }, []);
  } catch (error) {
    logError(error);
    return [];
  }
}

/**
 * Applies only safe, explicit normalization rules to a cell value.
 * Nullish values and invalid dates/numbers become null; strings are trimmed.
 * Other value types are preserved without coercion.
 *
 * @param {*} value Source value.
 * @return {*} Normalized value.
 */
function normalizeValue(value) {
  try {
    if (value === null || typeof value === 'undefined') {
      return null;
    }
    if (value instanceof Date) {
      return isValidDate_(value) ? new Date(value.getTime()) : null;
    }
    if (typeof value === 'number') {
      return Number.isFinite(value) ? value : null;
    }
    if (typeof value === 'string') {
      return value.trim();
    }
    return value;
  } catch (error) {
    logError(error);
    return null;
  }
}

/**
 * Parses only supported date formats: Date, YYYY-MM-DD and DD.MM.YYYY.
 * It never substitutes an invalid value with the current date.
 *
 * @param {*} value Date value to parse.
 * @return {Date|null} Parsed local date at its original time, or null.
 */
function parseDate(value) {
  try {
    if (value instanceof Date) {
      return isValidDate_(value) ? new Date(value.getTime()) : null;
    }
    if (typeof value !== 'string') {
      return null;
    }

    var input = value.trim();
    var match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(input);
    if (match) {
      return createValidatedDate_(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    match = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(input);
    if (match) {
      return createValidatedDate_(Number(match[3]), Number(match[2]), Number(match[1]));
    }

    return null;
  } catch (error) {
    logError(error);
    return null;
  }
}

/**
 * Filters objects by dates in a named column. Only rows with a valid date in
 * dateColumn are returned. Omitting a boundary leaves that side of the range
 * open; an explicitly supplied invalid boundary returns an empty array.
 *
 * @param {Array<Object>} rows Rows to filter.
 * @param {string} dateColumn Property that contains a date.
 * @param {*} start Range start (inclusive), supported by parseDate.
 * @param {*} end Range end (inclusive), supported by parseDate.
 * @return {Array<Object>} Filtered row copies.
 */
function filterByDateRange(rows, dateColumn, start, end) {
  try {
    if (!Array.isArray(rows) || typeof dateColumn !== 'string' || !dateColumn) {
      return [];
    }

    var startDate = parseDate(start);
    var endDate = parseDate(end);
    if ((hasValue_(start) && !startDate) || (hasValue_(end) && !endDate)) {
      logError('filterByDateRange: invalid range boundary.');
      return [];
    }

    if (startDate) {
      startDate.setHours(0, 0, 0, 0);
    }
    if (endDate) {
      endDate.setHours(23, 59, 59, 999);
    }
    if (startDate && endDate && startDate.getTime() > endDate.getTime()) {
      logError('filterByDateRange: start date is after end date.');
      return [];
    }

    return rows.filter(function (row) {
      var rowDate = row && typeof row === 'object' ? parseDate(row[dateColumn]) : null;
      if (!rowDate) {
        return false;
      }
      var timestamp = rowDate.getTime();
      return (!startDate || timestamp >= startDate.getTime()) &&
          (!endDate || timestamp <= endDate.getTime());
    });
  } catch (error) {
    logError(error);
    return [];
  }
}

/**
 * Filters rows by an object of conditions. A condition can be an exact value,
 * an array of acceptable exact values, or a predicate function(value, row).
 *
 * @param {Array<Object>} rows Rows to filter.
 * @param {Object<string, *>=} conditions Conditions by column name.
 * @return {Array<Object>} Matching row copies.
 */
function filterByConditions(rows, conditions) {
  try {
    if (!Array.isArray(rows)) {
      return [];
    }
    if (!conditions || typeof conditions !== 'object' || Array.isArray(conditions)) {
      return rows.slice();
    }

    var columns = Object.keys(conditions);
    return rows.filter(function (row) {
      if (!row || typeof row !== 'object') {
        return false;
      }
      return columns.every(function (columnName) {
        var expected = conditions[columnName];
        var actual = normalizeValue(row[columnName]);
        if (typeof expected === 'function') {
          return Boolean(expected(actual, row));
        }
        if (Array.isArray(expected)) {
          return expected.some(function (value) {
            return areValuesEqual_(actual, normalizeValue(value));
          });
        }
        return areValuesEqual_(actual, normalizeValue(expected));
      });
    });
  } catch (error) {
    logError(error);
    return [];
  }
}

/**
 * Returns a new array sorted by a column. Original rows and their order are
 * not mutated. The order may be 'asc' or 'desc'; any other value means 'asc'.
 *
 * @param {Array<Object>} rows Rows to sort.
 * @param {string} columnName Property used for sorting.
 * @param {string=} order Sort order: 'asc' or 'desc'.
 * @return {Array<Object>} Sorted copy of rows.
 */
function sortByColumn(rows, columnName, order) {
  try {
    if (!Array.isArray(rows) || typeof columnName !== 'string' || !columnName) {
      return [];
    }
    var direction = String(order || 'asc').toLowerCase() === 'desc' ? -1 : 1;
    return rows.slice().sort(function (left, right) {
      return compareValues_(left && left[columnName], right && right[columnName]) * direction;
    });
  } catch (error) {
    logError(error);
    return [];
  }
}

/**
 * Groups rows into an object keyed by a column value. Missing values are kept
 * under the empty-string key rather than being discarded.
 *
 * @param {Array<Object>} rows Rows to group.
 * @param {string} columnName Property used as the group key.
 * @return {Object<string, Array<Object>>} Grouped rows.
 */
function groupByColumn(rows, columnName) {
  try {
    if (!Array.isArray(rows) || typeof columnName !== 'string' || !columnName) {
      return {};
    }
    return rows.reduce(function (groups, row) {
      if (!row || typeof row !== 'object') {
        return groups;
      }
      var key = getGroupKey_(row[columnName]);
      if (!Object.prototype.hasOwnProperty.call(groups, key)) {
        groups[key] = [];
      }
      groups[key].push(row);
      return groups;
    }, {});
  } catch (error) {
    logError(error);
    return {};
  }
}

/**
 * Sums finite numeric values from a column. Non-numeric values are ignored.
 * Sheet values read with getValues() preserve numeric cells as numbers.
 *
 * @param {Array<Object>} rows Rows to sum.
 * @param {string} columnName Property containing numeric values.
 * @return {number} Sum, or zero for invalid input.
 */
function sumColumn(rows, columnName) {
  try {
    if (!Array.isArray(rows) || typeof columnName !== 'string' || !columnName) {
      return 0;
    }
    return rows.reduce(function (sum, row) {
      var value = row && typeof row === 'object' ? row[columnName] : null;
      return typeof value === 'number' && Number.isFinite(value) ? sum + value : sum;
    }, 0);
  } catch (error) {
    logError(error);
    return 0;
  }
}

/**
 * Formats a finite number using the currency settings in CONFIG.
 *
 * @param {*} value Number to format.
 * @return {string} Currency value, for example "1 234.50 ₽".
 */
function formatCurrency(value) {
  try {
    var amount = typeof value === 'number' && Number.isFinite(value) ? value : 0;
    var sign = amount < 0 ? '-' : '';
    var parts = Math.abs(amount).toFixed(CONFIG.DECIMAL_PLACES).split('.');
    var integerPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, CONFIG.THOUSANDS_SEPARATOR);
    return sign + integerPart + '.' + parts[1] + ' ' + CONFIG.CURRENCY_SYMBOL;
  } catch (error) {
    logError(error);
    return '0.00 ' + CONFIG.CURRENCY_SYMBOL;
  }
}

/**
 * Recursively makes a value safe for transmission to HTML Service clients.
 * Undefined properties are removed; non-finite numbers become zero; dates
 * become ISO strings. Circular references are omitted.
 *
 * @param {*} obj Value to prepare.
 * @return {*} Serializable value.
 */
function prepareResponse(obj) {
  try {
    return prepareResponseValue_(obj, []);
  } catch (error) {
    logError(error);
    return {
      status: 'error',
      data: null,
      error: 'Не удалось подготовить ответ сервера.'
    };
  }
}

/**
 * Unified API entry point for the money module.
 *
 * @param {Object=} params Module parameters reserved for future use.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 */
function apiMoney(params) {
  try {
    if (typeof cashGetMoneyData !== 'function') {
      return createErrorResponse('Модуль «Деньги» не подключен. Проверьте файл gsmoney.gs.');
    }
    return prepareResponse(cashGetMoneyData(params));
  } catch (error) {
    logError(error);
    return createErrorResponse(error);
  }
}

/**
 * API contract stub for the capital module.
 *
 * @param {Object=} params Module parameters reserved for future use.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 */
function apiCapital(params) {
  return createModuleStubResponse_('capital', params);
}

/**
 * Unified API entry point for the profit module.
 *
 * @param {Object=} params Module parameters reserved for future use.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 */
function apiProfit(params) {
  try {
    if (typeof profitGetProfitData !== 'function') {
      return createErrorResponse('Модуль «Прибыль» не подключен. Проверьте файл gsprofit.gs.');
    }
    return prepareResponse(profitGetProfitData(params));
  } catch (error) {
    logError(error);
    return createErrorResponse(error);
  }
}

/**
 * API contract stub for the dashboard module.
 *
 * @param {Object=} params Module parameters reserved for future use.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 */
function apiDashboard(params) {
  return createModuleStubResponse_('dashboard', params);
}

/**
 * Creates the shared, non-business API payload shape for a module.
 *
 * @param {string} moduleName Module identifier.
 * @param {Object=} params Request parameters.
 * @return {{status: string, data: Object|null, error: string|null}} API response.
 * @private
 */
function createModuleStubResponse_(moduleName, params) {
  try {
    logDebug('API stub requested: ' + moduleName);
    return createSuccessResponse({
      kpis: {},
      tables: {},
      charts: {},
      meta: {
        module: moduleName,
        implemented: false,
        params: typeof params === 'undefined' ? {} : params
      }
    });
  } catch (error) {
    logError(error);
    return createErrorResponse(error);
  }
}

/**
 * Writes a debug message to the Apps Script execution log.
 *
 * @param {*} message Message or value to log.
 */
function logDebug(message) {
  log_("DEBUG", message);
}

/**
 * Writes an informational message to the Apps Script execution log.
 *
 * @param {*} message Message or value to log.
 */
function logInfo(message) {
  log_('INFO', message);
}

/**
 * Writes an error message to the Apps Script execution log.
 *
 * @param {*} message Error or value to log.
 */
function logError(message) {
  log_('ERROR', message);
}

/**
 * @param {Array<*>} row Candidate spreadsheet row.
 * @return {boolean} Whether every value is empty after normalization.
 * @private
 */
function isEmptyRow_(row) {
  return row.every(function (value) {
    var normalized = normalizeValue(value);
    return normalized === null || normalized === '';
  });
}

/**
 * @param {*} value Value to test.
 * @return {boolean} Whether the value was explicitly supplied.
 * @private
 */
function hasValue_(value) {
  return value !== null && typeof value !== 'undefined' && value !== '';
}

/**
 * @param {number} year Full year.
 * @param {number} month One-based month.
 * @param {number} day Day of month.
 * @return {Date|null} Valid local date or null.
 * @private
 */
function createValidatedDate_(year, month, day) {
  var date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
      ? date
      : null;
}

/**
 * @param {*} value Candidate date.
 * @return {boolean} Whether value is a valid Date.
 * @private
 */
function isValidDate_(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * @param {*} left Left value.
 * @param {*} right Right value.
 * @return {boolean} Whether values are equal without type coercion.
 * @private
 */
function areValuesEqual_(left, right) {
  if (left instanceof Date && right instanceof Date) {
    return left.getTime() === right.getTime();
  }
  return left === right;
}

/**
 * @param {*} left Left value.
 * @param {*} right Right value.
 * @return {number} Comparison result.
 * @private
 */
function compareValues_(left, right) {
  var normalizedLeft = normalizeValue(left);
  var normalizedRight = normalizeValue(right);
  var leftEmpty = normalizedLeft === null || normalizedLeft === '';
  var rightEmpty = normalizedRight === null || normalizedRight === '';
  if (leftEmpty && rightEmpty) {
    return 0;
  }
  if (leftEmpty) {
    return 1;
  }
  if (rightEmpty) {
    return -1;
  }
  if (normalizedLeft instanceof Date && normalizedRight instanceof Date) {
    return normalizedLeft.getTime() - normalizedRight.getTime();
  }
  if (typeof normalizedLeft === 'number' && typeof normalizedRight === 'number') {
    return normalizedLeft - normalizedRight;
  }
  return String(normalizedLeft).localeCompare(String(normalizedRight), 'ru', {
    numeric: true,
    sensitivity: 'base'
  });
}

/**
 * @param {*} value Group value.
 * @return {string} Stable object key.
 * @private
 */
function getGroupKey_(value) {
  var normalized = normalizeValue(value);
  if (normalized === null) {
    return '';
  }
  if (normalized instanceof Date) {
    return normalized.toISOString();
  }
  return String(normalized);
}

/**
 * @param {*} value Value to serialize.
 * @param {Array<Object>} ancestors Values in the current traversal path.
 * @return {*} Serializable value or undefined when it must be removed.
 * @private
 */
function prepareResponseValue_(value, ancestors) {
  if (typeof value === 'undefined' || typeof value === 'function') {
    return undefined;
  }
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  if (value instanceof Date) {
    return isValidDate_(value) ? value.toISOString() : null;
  }
  if (typeof value !== 'object') {
    return String(value);
  }
  if (ancestors.indexOf(value) !== -1) {
    return undefined;
  }

  var nextAncestors = ancestors.concat([value]);
  if (Array.isArray(value)) {
    return value.reduce(function (preparedItems, item) {
      var preparedItem = prepareResponseValue_(item, nextAncestors);
      if (typeof preparedItem !== 'undefined') {
        preparedItems.push(preparedItem);
      }
      return preparedItems;
    }, []);
  }

  return Object.keys(value).reduce(function (preparedObject, key) {
    var preparedValue = prepareResponseValue_(value[key], nextAncestors);
    if (typeof preparedValue !== 'undefined') {
      preparedObject[key] = preparedValue;
    }
    return preparedObject;
  }, {});
}

/**
 * @param {*} message Log payload.
 * @return {string} Safe log string.
 * @private
 */
function getErrorMessage_(message) {
  if (message instanceof Error) {
    return message.message || String(message);
  }
  if (typeof message === 'string' && message.trim()) {
    return message;
  }
  if (message === null || typeof message === 'undefined') {
    return 'Неизвестная ошибка.';
  }
  try {
    return String(message);
  } catch (error) {
    return 'Неизвестная ошибка.';
  }
}

/**
 * @param {string} level Log level.
 * @param {*} message Log payload.
 * @private
 */
function log_(level, message) {
  try {
    Logger.log('[' + level + '] ' + getErrorMessage_(message));
  } catch (error) {
    // Logging must never interrupt a controlled response.
  }
}
