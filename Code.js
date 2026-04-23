/**
 * Creates the menu item when the spreadsheet is opened.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open Architecture Exporter', 'showSidebar')
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the Google Sheet.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Architecture Exporter')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ══════════════════════════════════════════
   HELPERS
   ══════════════════════════════════════════ */

function getColLetter(col) {
  var letter = '';
  var temp = col;
  while (temp > 0) {
    var mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

function cellA1(row, col) {
  return getColLetter(col) + row;
}

function getValidationSignature(rule) {
  if (!rule) return null;
  var type = rule.getCriteriaType();
  var typeStr = type ? type.toString() : 'UNKNOWN';
  var sig = typeStr;
  try { sig += '|' + JSON.stringify(rule.getCriteriaValues()); } catch (_) { }
  return sig;
}

function mergeCellsToRanges(cells) {
  var byCols = {};
  for (var i = 0; i < cells.length; i++) {
    var c = cells[i];
    if (!byCols[c.col]) byCols[c.col] = [];
    byCols[c.col].push(c.row);
  }
  var ranges = [];
  var colKeys = Object.keys(byCols);
  for (var k = 0; k < colKeys.length; k++) {
    var col = Number(colKeys[k]);
    var rows = byCols[col];
    rows.sort(function (a, b) { return a - b; });
    var colLetter = getColLetter(col);
    var start = rows[0], end = rows[0];
    for (var j = 1; j < rows.length; j++) {
      if (rows[j] === end + 1) { end = rows[j]; }
      else {
        ranges.push(start === end ? colLetter + start : colLetter + start + ':' + colLetter + end);
        start = rows[j]; end = rows[j];
      }
    }
    ranges.push(start === end ? colLetter + start : colLetter + start + ':' + colLetter + end);
  }
  return ranges.join(', ');
}

/**
 * Extracts the source/criteria detail from a validation rule.
 */
function extractValidationSource(rule) {
  var info = {};
  try {
    var values = rule.getCriteriaValues();
    if (values && values.length > 0) {
      var processed = [];
      for (var i = 0; i < values.length; i++) {
        var v = values[i];
        if (v && typeof v === 'object' && typeof v.getA1Notation === 'function') {
          try {
            processed.push(v.getSheet().getName() + '!' + v.getA1Notation());
          } catch (_) {
            processed.push(v.getA1Notation());
          }
        } else if (v !== null && v !== undefined && v !== '') {
          processed.push(v.toString());
        }
      }
      if (processed.length > 0) {
        info.criteriaValues = processed;
      }
    }
  } catch (_) { }

  try { info.allowInvalid = rule.getAllowInvalid(); } catch (_) { }
  try {
    var helpText = rule.getHelpText();
    if (helpText) info.helpText = helpText;
  } catch (_) { }

  return info;
}

/**
 * Safely extracts chart type from an EmbeddedChart.
 * Tries multiple approaches so that unusual or unsupported chart types
 * don't throw and kill the entire export.
 */
function safeGetChartType(chart) {
  // Approach 1: chart.modify().getChartType() — standard path
  try {
    var builder = chart.modify();
    if (builder && typeof builder.getChartType === 'function') {
      var ct = builder.getChartType();
      if (ct) return ct.toString();
    }
  } catch (_) { }

  // Approach 2: inspect the chart's options spec for a chartType key
  try {
    var options = chart.getOptions();
    if (options) {
      // EmbeddedChartBuilder options sometimes expose chartType as an option
      var fromOptions = options.get('chartType');
      if (fromOptions) return fromOptions.toString();
    }
  } catch (_) { }

  // Approach 3: inspect the chart's as-JSON representation
  try {
    var spec = chart.getAs('application/json');
    if (spec) {
      var parsed = JSON.parse(spec.getDataAsString());
      if (parsed && parsed.chartType) return parsed.chartType.toString();
    }
  } catch (_) { }

  return 'UNSUPPORTED_TYPE';
}

/**
 * Safely extracts data ranges from an EmbeddedChart.
 */
function safeGetChartRanges(chart) {
  try {
    var ranges = chart.getRanges();
    if (!ranges || ranges.length === 0) return [];
    var result = [];
    for (var i = 0; i < ranges.length; i++) {
      try {
        result.push(ranges[i].getSheet().getName() + '!' + ranges[i].getA1Notation());
      } catch (_) {
        try { result.push(ranges[i].getA1Notation()); } catch (_2) { }
      }
    }
    return result;
  } catch (_) {
    return ['Unable to extract ranges'];
  }
}

/* ══════════════════════════════════════════
   MAIN EXPORT FUNCTION
   ══════════════════════════════════════════ */

function exportArchitecture(options) {
  try {
    options = options || {};

    // Defaults for boolean toggles
    var boolKeys = ['namedRanges', 'formulas', 'dataValidation', 'conditionalFormatting',
      'charts', 'headers', 'sampleData', 'mergedCells', 'sheetMetadata',
      'numberFormats', 'notes', 'crossSheetDeps'];
    for (var d = 0; d < boolKeys.length; d++) {
      if (options[boolKeys[d]] === undefined) options[boolKeys[d]] = true;
    }

    // Row limit: how many rows to scan for formulas, validation, notes, number formats
    var rowLimit = (typeof options.rowLimit === 'number' && options.rowLimit > 0)
      ? Math.floor(options.rowLimit)
      : 100;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var architecture = {
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      exportedAt: new Date().toISOString(),
      rowLimitApplied: rowLimit,
      sheetNames: [],
      detailsBySheet: {}
    };

    // ── Named Ranges ──
    if (options.namedRanges) {
      architecture.namedRanges = [];
      var namedRanges = ss.getNamedRanges();
      for (var i = 0; i < namedRanges.length; i++) {
        var nr = namedRanges[i];
        var rng = nr.getRange();
        architecture.namedRanges.push({
          name: nr.getName(),
          range: rng.getSheet().getName() + '!' + rng.getA1Notation()
        });
      }
    }

    var sheets = ss.getSheets();

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var sheetName = sheet.getName();
      architecture.sheetNames.push(sheetName);

      var det = {};

      // ── Sheet Metadata ──
      if (options.sheetMetadata) {
        det.metadata = {
          index: sheet.getIndex(),
          isHidden: sheet.isSheetHidden(),
          tabColor: sheet.getTabColor(),
          frozenRows: sheet.getFrozenRows(),
          frozenColumns: sheet.getFrozenColumns(),
          maxRows: sheet.getMaxRows(),
          maxColumns: sheet.getMaxColumns(),
          lastRowWithContent: sheet.getLastRow(),
          lastColumnWithContent: sheet.getLastColumn()
        };
      }

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();

      // Skip empty sheets gracefully
      if (lastRow === 0 || lastCol === 0) {
        architecture.detailsBySheet[sheetName] = det;
        continue;
      }

      var dataRange = sheet.getRange(1, 1, lastRow, lastCol);
      var startRow = 1;
      var startCol = 1;
      var numRows = lastRow;
      var numCols = lastCol;

      // Effective scan rows: the lesser of rowLimit and actual data rows
      var scanRows = Math.min(rowLimit, numRows);

      // ── Headers (row 1 values) ──
      if (options.headers && numRows > 0 && numCols > 0) {
        var headerRange = sheet.getRange(startRow, startCol, 1, numCols);
        var headerValues = headerRange.getValues()[0];
        det.headers = [];
        for (var h = 0; h < headerValues.length; h++) {
          var val = headerValues[h];
          if (val !== '' && val !== null) {
            det.headers.push({
              column: cellA1(startRow, startCol + h),
              value: val.toString()
            });
          }
        }
      }

      // ── Sample Data (first 5 data rows after header) ──
      if (options.sampleData && numRows > 1 && numCols > 0) {
        var sampleRowCount = Math.min(5, numRows - 1);
        var sampleRange = sheet.getRange(startRow + 1, startCol, sampleRowCount, numCols);
        var sampleValues = sampleRange.getValues();
        det.sampleData = {
          range: cellA1(startRow + 1, startCol) + ':' + cellA1(startRow + sampleRowCount, startCol + numCols - 1),
          rows: []
        };
        for (var sr = 0; sr < sampleValues.length; sr++) {
          var rowObj = {};
          for (var sc = 0; sc < sampleValues[sr].length; sc++) {
            var cellVal = sampleValues[sr][sc];
            if (cellVal !== '' && cellVal !== null) {
              var colKey = (det.headers && det.headers[sc]) ? det.headers[sc].value : getColLetter(startCol + sc);
              rowObj[colKey] = cellVal;
            }
          }
          if (Object.keys(rowObj).length > 0) {
            det.sampleData.rows.push(rowObj);
          }
        }
      }

      // ── Number Formats (scan up to rowLimit rows to detect date/currency/% patterns) ──
      if (options.numberFormats && numRows > 0 && numCols > 0) {
        var fmtRowCount = Math.min(scanRows, numRows);
        var fmtRange = sheet.getRange(startRow, startCol, fmtRowCount, numCols);
        var formats = fmtRange.getNumberFormats();
        var formatMap = {};
        for (var fr = 0; fr < formats.length; fr++) {
          for (var fc = 0; fc < formats[fr].length; fc++) {
            var fmt = formats[fr][fc];
            if (fmt && fmt !== '' && fmt !== '0.###############') {
              var colLtr = getColLetter(startCol + fc);
              if (!formatMap[colLtr]) formatMap[colLtr] = fmt;
            }
          }
        }
        if (Object.keys(formatMap).length > 0) {
          det.numberFormats = [];
          var fmtKeys = Object.keys(formatMap);
          for (var fk = 0; fk < fmtKeys.length; fk++) {
            var headerName = '';
            if (det.headers) {
              for (var hh = 0; hh < det.headers.length; hh++) {
                if (det.headers[hh].column.replace(/[0-9]/g, '') === fmtKeys[fk]) {
                  headerName = det.headers[hh].value;
                  break;
                }
              }
            }
            det.numberFormats.push({
              column: fmtKeys[fk],
              columnHeader: headerName || null,
              format: formatMap[fmtKeys[fk]]
            });
          }
        }
      }

      // ── Merged Cells ──
      if (options.mergedCells) {
        try {
          var mergedRanges = sheet.getRange(1, 1, numRows, numCols).getMergedRanges();
          if (mergedRanges.length > 0) {
            det.mergedCells = [];
            for (var m = 0; m < mergedRanges.length; m++) {
              det.mergedCells.push(mergedRanges[m].getA1Notation());
            }
          }
        } catch (_) { }
      }

      // ── Notes / Comments (scanned up to rowLimit rows) ──
      if (options.notes && scanRows > 0 && numCols > 0) {
        try {
          var noteScanRange = sheet.getRange(startRow, startCol, scanRows, numCols);
          var notes = noteScanRange.getNotes();
          var notesList = [];
          for (var nr2 = 0; nr2 < notes.length; nr2++) {
            for (var nc = 0; nc < notes[nr2].length; nc++) {
              if (notes[nr2][nc]) {
                notesList.push({
                  cell: cellA1(startRow + nr2, startCol + nc),
                  note: notes[nr2][nc]
                });
              }
            }
          }
          if (notesList.length > 0) det.notes = notesList;
        } catch (_) { }
      }

      // ── Formulas (scanned up to rowLimit rows) ──
      if (options.formulas && scanRows > 0 && numCols > 0) {
        det.formulas = [];
        var formulaScanRange = sheet.getRange(startRow, startCol, scanRows, numCols);
        var allFormulas = formulaScanRange.getFormulas();
        for (var r = 0; r < allFormulas.length; r++) {
          var fRow = allFormulas[r];
          for (var c = 0; c < fRow.length; c++) {
            if (fRow[c]) {
              det.formulas.push({
                cell: cellA1(startRow + r, startCol + c),
                formula: fRow[c]
              });
            }
          }
        }
        if (scanRows < numRows) {
          det.formulasScanNote = 'Scanned rows 1–' + scanRows + ' of ' + numRows + ' total';
        }
      }

      // ── Cross-Sheet Dependencies (extracted from formulas) ──
      if (options.crossSheetDeps && options.formulas && det.formulas) {
        var depsSet = {};
        var sheetRefPattern = /(?:'([^']+)'|([A-Za-z0-9_]+))!/g;
        for (var fd = 0; fd < det.formulas.length; fd++) {
          var formula = det.formulas[fd].formula;
          var match;
          while ((match = sheetRefPattern.exec(formula)) !== null) {
            var refSheet = match[1] || match[2];
            if (refSheet !== sheetName) {
              depsSet[refSheet] = true;
            }
          }
          sheetRefPattern.lastIndex = 0;
        }
        var depKeys = Object.keys(depsSet);
        if (depKeys.length > 0) {
          det.crossSheetDependencies = depKeys;
        }
      }

      // ── Data Validation (scanned up to rowLimit rows, with full source info) ──
      if (options.dataValidation) {
        det.dataValidation = [];
        var dvScanRows = Math.min(scanRows, numRows);

        if (dvScanRows > 0 && numCols > 0) {
          var dvRange = sheet.getRange(startRow, startCol, dvScanRows, numCols);
          var validations = dvRange.getDataValidations();
          var validationGroups = {};

          for (var vr = 0; vr < dvScanRows; vr++) {
            var rowRules = validations[vr];
            for (var vc = 0; vc < numCols; vc++) {
              var rule = rowRules[vc];
              if (!rule) continue;
              var sig = getValidationSignature(rule);
              if (!sig) continue;

              if (!validationGroups[sig]) {
                var vType = rule.getCriteriaType();
                validationGroups[sig] = {
                  type: vType ? vType.toString() : 'UNKNOWN',
                  cells: [],
                  source: extractValidationSource(rule)
                };
              }
              validationGroups[sig].cells.push({ row: startRow + vr, col: startCol + vc });
            }
          }

          var vKeys = Object.keys(validationGroups);
          for (var vk = 0; vk < vKeys.length; vk++) {
            var group = validationGroups[vKeys[vk]];
            var entry = {
              range: mergeCellsToRanges(group.cells),
              criteriaType: group.type
            };
            if (group.source.criteriaValues) entry.criteriaValues = group.source.criteriaValues;
            if (group.source.helpText) entry.helpText = group.source.helpText;
            if (group.source.allowInvalid !== undefined) entry.allowInvalid = group.source.allowInvalid;
            det.dataValidation.push(entry);
          }

          if (dvScanRows < numRows) {
            det.dataValidationScanNote = 'Scanned rows 1–' + dvScanRows + ' of ' + numRows + ' total';
          }
        }
      }

      // ── Conditional Formatting ──
      if (options.conditionalFormatting) {
        det.conditionalFormatting = [];
        try {
          var cfRules = sheet.getConditionalFormatRules();
          for (var ci = 0; ci < cfRules.length; ci++) {
            var cfRule = cfRules[ci];
            var cfType = 'UNKNOWN';
            var cfDetails = {};

            try {
              var boolCond = cfRule.getBooleanCondition();
              if (boolCond) {
                cfType = boolCond.getCriteriaType().toString();
                try {
                  var cfVals = boolCond.getCriteriaValues();
                  if (cfVals && cfVals.length > 0) {
                    cfDetails.criteriaValues = [];
                    for (var cv = 0; cv < cfVals.length; cv++) {
                      cfDetails.criteriaValues.push(cfVals[cv] !== null ? cfVals[cv].toString() : null);
                    }
                  }
                } catch (_) { }
              } else {
                try {
                  if (cfRule.getGradientCondition()) cfType = 'GRADIENT_COLOR_SCALE';
                } catch (_) { }
              }
            } catch (_) { }

            try {
              var cfRanges = cfRule.getRanges().map(function (r) {
                return r.getA1Notation();
              }).join(', ');
              var cfEntry = { ranges: cfRanges, type: cfType };
              if (cfDetails.criteriaValues) cfEntry.criteriaValues = cfDetails.criteriaValues;
              det.conditionalFormatting.push(cfEntry);
            } catch (_) { }
          }
        } catch (_) { }
      }

      // ── Charts (defensive per-property extraction) ──
      if (options.charts) {
        det.charts = [];
        try {
          var charts = sheet.getCharts();
          for (var chi = 0; chi < charts.length; chi++) {
            var chart = charts[chi];
            var chartInfo = {};

            // ID
            try { chartInfo.chartId = chart.getChartId(); }
            catch (_) { chartInfo.chartId = null; }

            // Type — uses multi-attempt helper
            chartInfo.chartType = safeGetChartType(chart);

            // Data ranges — sheet-qualified where possible
            chartInfo.dataRanges = safeGetChartRanges(chart);

            // Position
            try {
              var pos = chart.getContainerInfo();
              chartInfo.position = {
                anchorColumn: pos.getAnchorColumn(),
                anchorRow: pos.getAnchorRow(),
                offsetX: pos.getOffsetX(),
                offsetY: pos.getOffsetY()
              };
            } catch (_) { }

            det.charts.push(chartInfo);
          }
        } catch (chartErr) {
          det.chartsError = chartErr.message;
        }
      }

      architecture.detailsBySheet[sheetName] = det;
    }

    // ── File Generation ──
    var jsonOutput = JSON.stringify(architecture, null, 2);
    var timestamp = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd-HH-mm-ss');
    var fileName = ss.getName() + '-' + timestamp + '.json';

    var folderName = 'ArchEx Output';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var file = folder.createFile(fileName, jsonOutput, MimeType.PLAIN_TEXT);

    var result = { success: true, url: file.getUrl(), name: fileName, content: jsonOutput };

    // ── Optional: Download spreadsheet as Excel ──
    if (options.downloadXlsx) {
      try {
        var xlsxUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
        var xlsxBlob = UrlFetchApp.fetch(xlsxUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
        }).getBlob().setName(ss.getName() + '.xlsx');
        var xlsxFile = folder.createFile(xlsxBlob);
        result.xlsxUrl = 'https://drive.google.com/uc?export=download&id=' + xlsxFile.getId();
        result.xlsxDriveUrl = xlsxFile.getUrl();
      } catch (xlsxError) {
        result.xlsxError = xlsxError.message;
      }
    }

    return result;
  } catch (error) {
    return { success: false, error: error.message + ' — ' + error.stack };
  }
}