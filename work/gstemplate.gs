/**
 * Web-app shell for the financial management dashboard.
 *
 * This stage deliberately does not read Google Sheets. Sheet access and
 * calculations will be added in the next stages.
 */

var APP_ROUTES = Object.freeze({
  dashboard: {
    file: 'dashboard',
    title: 'Финансовый дашборд'
  },
  money: {
    file: 'money',
    title: 'Деньги'
  },
  capital: {
    file: 'capital',
    title: 'Капитал'
  },
  profit: {
    file: 'profit',
    title: 'Прибыль'
  }
});

/**
 * Serves a complete page for the requested route.
 * Example: .../exec?page=money
 *
 * @param {Object} e Web-app request event.
 * @return {GoogleAppsScript.HTML.HtmlOutput} Page output.
 */
function doGet(e) {
  var route = getRoute_(e);
  var template = HtmlService.createTemplateFromFile('template');

  // Apps Script serves the page from a temporary googleusercontent.com
  // sandbox URL, not from the .../exec address in the browser bar. Relative
  // links (href="?page=...") resolve against that sandbox URL and silently
  // break navigation, so every link must be built from the real, stable
  // web-app URL instead.
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.currentPage = route.id;
  template.pageTitle = route.title;
  template.pageFile = route.file;

  // The period picked on the Дэшборд page is carried forward as ?start=&end=
  // query params (see gsdashboard.gs / dashboard.html) so every other page
  // opens pre-filtered to the same period. periodQuery is pre-built here
  // (already URL-encoded) and appended, via <?= ?>, to every nav link in
  // template.html — the auto-escaping there turns its "&" into "&amp;",
  // which is what a literal "&" must look like inside an href attribute.
  var periodStart = e && e.parameter && typeof e.parameter.start === 'string' ? e.parameter.start.trim() : '';
  var periodEnd = e && e.parameter && typeof e.parameter.end === 'string' ? e.parameter.end.trim() : '';
  var periodQuery = '';
  if (periodStart) { periodQuery += '&start=' + encodeURIComponent(periodStart); }
  if (periodEnd) { periodQuery += '&end=' + encodeURIComponent(periodEnd); }
  template.periodStart = periodStart;
  template.periodEnd = periodEnd;
  template.periodQuery = periodQuery;

  return template
      .evaluate()
      .setTitle(route.title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns the HTML content of a partial file for a template.
 * The caller controls the filename through APP_ROUTES only.
 *
 * @param {string} filename HTML filename without extension.
 * @return {string} Raw HTML from the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Resolves and validates the page query parameter.
 * Unknown or missing values always open the dashboard.
 *
 * @param {Object} e Web-app request event.
 * @return {{id: string, file: string, title: string}} Route definition.
 */
function getRoute_(e) {
  var requestedPage = e && e.parameter ? e.parameter.page : '';
  var pageId = typeof requestedPage === 'string'
      ? requestedPage.toLowerCase().trim()
      : '';

  if (!Object.prototype.hasOwnProperty.call(APP_ROUTES, pageId)) {
    pageId = 'dashboard';
  }

  return {
    id: pageId,
    file: APP_ROUTES[pageId].file,
    title: APP_ROUTES[pageId].title
  };
}