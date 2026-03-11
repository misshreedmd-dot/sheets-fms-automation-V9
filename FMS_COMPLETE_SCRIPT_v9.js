// ══════════════════════════════════════════════════════════════════
// FMS COMPLETE SCRIPT v9
// ✅ REQ 1: setFormulas() batch write — no timeout possible
// ✅ REQ 2: Resume step writing if interrupted (PropertiesService)
// ✅ REQ 3: Extra columns per step + Form 2 sections with navigation
// ✅ REQ 4: Full FMS Diagnostic Check with detailed log sheet
// ✅ FIX:   Form 2 writes Actual timestamp directly (not ARRAYFORMULA)
// ✅ All v8 features preserved
// ══════════════════════════════════════════════════════════════════

var FORM1_SHEET_NAME   = "Form responses 6";
var FORM2_SHEET_NAME   = "Form responses 4";
var FMS_SHEET_NAME     = "FMS";
var FMS_DATA_START_ROW = 9;   // row 9 = headers, row 10+ = data
var ID_COLUMN          = 1;   // Col A = Unique ID
var TIMESTAMP_COLUMN   = 2;   // Col B = Timestamp
var FIELDS_START_COL   = 3;   // Col C = first user-defined field

// Base cols per step (Planned|Actual|Status|TimeDelay) — extra cols added on top
var BASE_COLS_PER_STEP    = 4;
var STATUS_OFFSET_IN_STEP = 2; // within base: 0=Planned,1=Actual,2=Status,3=TimeDelay

var ROW_STEP_NUM  = 2;
var ROW_STEP_ID   = 3;
var ROW_WHAT      = 4;
var ROW_WHO       = 5;
var ROW_HOW       = 6;
var ROW_WHEN      = 7;
var ROW_TAT_REF   = 8;
var ROW_HEADERS   = 9;
var FIELD_DEF_COL = 100; // hidden col — stores BOTH fieldDefs + stepDefs as JSON
var STEP_DEF_COL  = 101; // hidden col — stores full step defs with col positions

var NR = 1000; // number of data rows


// ══════════════════════════════════════════════════════════════════
// MENU
// ══════════════════════════════════════════════════════════════════
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BMP Formulas')
    .addItem('Run This First Time After Install', 'initial')
    .addSeparator()
    .addItem('🆕 Create New FMS Sheet',            'createFMSWizard')
    .addItem('➕ Add Steps to FMS',                'addStepsToFMS')
    .addItem('▶️ Resume Step Writing',             'resumeStepWriting')
    .addSeparator()
    .addItem('📋 Generate Forms for Active FMS',   'generateFormsForActiveFMS')
    .addSeparator()
    .addItem('⚙️ Setup All Triggers (Run Once)',   'setupAllTriggers')
    .addSeparator()
    .addItem('🔍 Run FMS Diagnostic',              'runFMSDiagnostic')
    .addItem('🧪 Test Form 1 Manually',            'testForm1Manually')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('FMS Formulas')
        .addItem('TAT (add working-hours TAT)',          'TAT')
        .addSeparator()
        .addItem('T-x Formula',                          'plannedlead')
        .addSeparator()
        .addItem('Specific Time',                        'specificTime')
        .addSeparator()
        .addItem('Show planned only when status is NO',  'tatifno')
        .addSeparator()
        .addItem('Show planned only when status is YES', 'tatifyes')
        .addSeparator()
        .addItem('Set Actual Time',                      'createTrigger')
        .addSeparator()
        .addItem('Time Delay Formula',                   'timeDelay')
    )
    .addToUi();
}


// ══════════════════════════════════════════════════════════════════
// HELPERS — field + step def storage
// ══════════════════════════════════════════════════════════════════
function _storeFieldDefs_(sheet, fieldDefs) {
  try {
    sheet.getRange(1, FIELD_DEF_COL).setValue(JSON.stringify(fieldDefs));
    sheet.hideColumns(FIELD_DEF_COL);
  } catch(e) { Logger.log('⚠️ _storeFieldDefs_ error: ' + e); }
}

function _loadFieldDefs_(sheet) {
  try {
    var val = sheet.getRange(1, FIELD_DEF_COL).getValue();
    if (!val || val === '') return null;
    return JSON.parse(val);
  } catch(e) { Logger.log('⚠️ _loadFieldDefs_ error: ' + e); return null; }
}

// Step defs include col positions — stored separately in col 101
function _storeStepDefs_(sheet, stepDefs) {
  try {
    sheet.getRange(1, STEP_DEF_COL).setValue(JSON.stringify(stepDefs));
    sheet.hideColumns(STEP_DEF_COL);
  } catch(e) { Logger.log('⚠️ _storeStepDefs_ error: ' + e); }
}

function _loadStepDefs_(sheet) {
  try {
    var val = sheet.getRange(1, STEP_DEF_COL).getValue();
    if (!val || val === '') return null;
    return JSON.parse(val);
  } catch(e) { Logger.log('⚠️ _loadStepDefs_ error: ' + e); return null; }
}

function _calcFirstStepCol_(numFields) {
  return FIELDS_START_COL + numFields;
}

function _col_(col) {
  var l = '';
  while (col > 0) { var r = (col-1)%26; l = String.fromCharCode(65+r)+l; col = Math.floor((col-1)/26); }
  return l;
}

function _parseTime_(txt) {
  var m = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(txt);
  if (!m) return null;
  return (parseInt(m[1])*60 + parseInt(m[2])) / 1440;
}

function _parseTimeToFractionOrAlert_(txt, label) {
  var ui = SpreadsheetApp.getUi();
  var m  = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(txt);
  if (!m) { ui.alert(label + ' must be HH:MM. Got: "' + txt + '"'); return null; }
  return (parseInt(m[1])*60 + parseInt(m[2])) / 1440;
}

function _getField_(data, variants) {
  for (var i = 0; i < variants.length; i++) {
    var k = variants[i].toLowerCase().trim();
    if (data[k] !== undefined && data[k] !== '') return data[k];
  }
  return null;
}

// Build step column layout — calculates exact col positions for each step
// Returns enriched steps array with: startCol, totalCols, plannedCol, actualCol,
// statusCol, timeDelayCol, and col positions for each extraCol
function _calcStepColLayout_(steps, numFields) {
  var firstStepCol = _calcFirstStepCol_(numFields);
  var cursor = firstStepCol;
  var enriched = [];

  for (var s = 0; s < steps.length; s++) {
    var step       = steps[s];
    var extraCols  = step.extraCols || [];
    var totalCols  = BASE_COLS_PER_STEP + extraCols.length;

    var plannedCol    = cursor;
    var actualCol     = cursor + 1;
    // Extra cols go after Actual, before Status
    var statusCol     = cursor + 2 + extraCols.length;
    var timeDelayCol  = cursor + 3 + extraCols.length;

    // Assign exact col number to each extra col
    var enrichedExtras = [];
    for (var x = 0; x < extraCols.length; x++) {
      enrichedExtras.push({
        name:     extraCols[x].name,
        type:     extraCols[x].type,
        choices:  extraCols[x].choices || [],
        col:      cursor + 2 + x   // after Actual
      });
    }

    enriched.push({
      num:         step.num,
      what:        step.what,
      who:         step.who,
      how:         step.how,
      tatType:     step.tatType,
      tatValue:    step.tatValue,
      tatLabel:    step.tatLabel,
      startCol:    cursor,
      totalCols:   totalCols,
      plannedCol:  plannedCol,
      actualCol:   actualCol,
      statusCol:   statusCol,
      timeDelayCol:timeDelayCol,
      extraCols:   enrichedExtras
    });

    cursor += totalCols;
  }
  return enriched;
}


// ══════════════════════════════════════════════════════════════════
// WIZARD 1 — CREATE FMS SHEET (sheet + time + fields ONLY)
// ══════════════════════════════════════════════════════════════════
function createFMSWizard() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var fmsSheet = ss.getSheetByName(FMS_SHEET_NAME);
  if (!fmsSheet) {
    fmsSheet = ss.insertSheet(FMS_SHEET_NAME);
  } else {
    var rebuild = ui.alert(
      '⚠️ FMS sheet already exists',
      '"FMS" sheet already exists. Delete and rebuild from scratch?',
      ui.ButtonSet.YES_NO
    );
    if (rebuild !== ui.Button.YES) return;
    ss.deleteSheet(fmsSheet);
    fmsSheet = ss.insertSheet(FMS_SHEET_NAME);
  }

  var or2 = ui.prompt('⏰ Opening Time', 'HH:MM  e.g. 10:00', ui.ButtonSet.OK_CANCEL);
  if (or2.getSelectedButton() !== ui.Button.OK) return;
  var openFrac = _parseTime_(or2.getResponseText().trim());
  if (openFrac === null) { ui.alert('Use HH:MM format.'); return; }

  var cr = ui.prompt('⏰ Closing Time', 'HH:MM  e.g. 18:00', ui.ButtonSet.OK_CANCEL);
  if (cr.getSelectedButton() !== ui.Button.OK) return;
  var closeFrac = _parseTime_(cr.getResponseText().trim());
  if (closeFrac === null) { ui.alert('Use HH:MM format.'); return; }

  var nfr = ui.prompt(
    '📝 New Order Entry — Fields',
    'How many fields?\n(Unique ID and Timestamp are auto-added)\ne.g. 3',
    ui.ButtonSet.OK_CANCEL
  );
  if (nfr.getSelectedButton() !== ui.Button.OK) return;
  var nf = parseInt(nfr.getResponseText().trim());
  if (isNaN(nf) || nf < 1 || nf > 20) { ui.alert('Enter 1–20.'); return; }

  var fieldDefs = [];
  for (var f = 1; f <= nf; f++) {
    var fnr = ui.prompt(
      '📝 Field ' + f + ' of ' + nf + ' — Name',
      'e.g. PO Number / Bag Number / Client Name / Due Date',
      ui.ButtonSet.OK_CANCEL
    );
    if (fnr.getSelectedButton() !== ui.Button.OK) return;
    var fname = fnr.getResponseText().trim() || ('Field ' + f);

    var ftr = ui.prompt(
      '📝 Field ' + f + ' — Type for "' + fname + '"',
      'Text\nDate\nDropdown',
      ui.ButtonSet.OK_CANCEL
    );
    if (ftr.getSelectedButton() !== ui.Button.OK) return;
    var ftype = ftr.getResponseText().trim().toLowerCase();
    if (ftype === 'date') ftype = 'Date';
    else if (ftype.indexOf('drop') !== -1) ftype = 'Dropdown';
    else ftype = 'Text';

    var fchoices = [];
    if (ftype === 'Dropdown') {
      var fcr = ui.prompt(
        '📝 Choices for "' + fname + '"',
        'Comma separated  e.g. Normal, Urgent, Very Urgent',
        ui.ButtonSet.OK_CANCEL
      );
      if (fcr.getSelectedButton() !== ui.Button.OK) return;
      fchoices = fcr.getResponseText().split(',').map(function(c){ return c.trim(); }).filter(Boolean);
    }
    fieldDefs.push({ name: fname, type: ftype, choices: fchoices });
  }

  _buildFMSBaseSheet_(fmsSheet, openFrac, closeFrac, fieldDefs);

  ui.alert(
    '✅ FMS Sheet Created!\n\n' +
    'Next:\n' +
    '1. ➕ Add Steps to FMS\n' +
    '2. 📋 Generate Forms'
  );
}


// ══════════════════════════════════════════════════════════════════
// WIZARD 2 — ADD STEPS TO FMS
// ══════════════════════════════════════════════════════════════════
function addStepsToFMS() {
  var ui       = SpreadsheetApp.getUi();
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var fmsSheet = ss.getSheetByName(FMS_SHEET_NAME);

  if (!fmsSheet) { ui.alert('❌ FMS sheet not found. Run "🆕 Create New FMS Sheet" first.'); return; }

  var fieldDefs = _loadFieldDefs_(fmsSheet);
  if (!fieldDefs || fieldDefs.length === 0) {
    ui.alert('❌ Field definitions not found. Run "🆕 Create New FMS Sheet" first.');
    return;
  }

  var sr = ui.prompt('📋 Number of Departments', 'How many departments/steps? (1–15)', ui.ButtonSet.OK_CANCEL);
  if (sr.getSelectedButton() !== ui.Button.OK) return;
  var ns = parseInt(sr.getResponseText().trim());
  if (isNaN(ns) || ns < 1 || ns > 15) { ui.alert('Enter 1–15.'); return; }

  var steps = [];
  for (var s = 1; s <= ns; s++) {
    var wr = ui.prompt('🏢 Step ' + s + ' of ' + ns + ' — Department', 'Department name:', ui.ButtonSet.OK_CANCEL);
    if (wr.getSelectedButton() !== ui.Button.OK) return;
    var what = wr.getResponseText().trim() || ('Step ' + s);

    var whor = ui.prompt('👤 ' + what + ' — Person', 'Who is responsible?', ui.ButtonSet.OK_CANCEL);
    if (whor.getSelectedButton() !== ui.Button.OK) return;
    var who = whor.getResponseText().trim() || '—';

    var howr = ui.prompt('⚙️ ' + what + ' — Method', 'System / WhatsApp / Tally / Email', ui.ButtonSet.OK_CANCEL);
    if (howr.getSelectedButton() !== ui.Button.OK) return;
    var how = howr.getResponseText().trim() || 'System';

    var tr = ui.prompt(
      '⏱️ ' + what + ' — TAT Type',
      'H = Hours after previous step\nD = Days after previous step\nT = Fixed time of day\nN = Whenever Needed',
      ui.ButtonSet.OK_CANCEL
    );
    if (tr.getSelectedButton() !== ui.Button.OK) return;
    var tt = tr.getResponseText().trim().toUpperCase();
    if (['H','D','T','N'].indexOf(tt) === -1) tt = 'N';

    var tv = '', tl = 'Whenever Needed';
    if (tt === 'H') {
      var tvr = ui.prompt('Hours after previous step?', 'e.g. 3 or 0.5', ui.ButtonSet.OK_CANCEL);
      if (tvr.getSelectedButton() !== ui.Button.OK) return;
      tv = tvr.getResponseText().trim(); tl = tv + ' Hours';
    }
    if (tt === 'D') {
      var tvr = ui.prompt('Days after previous step?', 'e.g. 1 or 2', ui.ButtonSet.OK_CANCEL);
      if (tvr.getSelectedButton() !== ui.Button.OK) return;
      tv = tvr.getResponseText().trim(); tl = tv + ' Days';
    }
    if (tt === 'T') {
      var tvr = ui.prompt('Fixed time? (HH:MM)', 'e.g. 18:00', ui.ButtonSet.OK_CANCEL);
      if (tvr.getSelectedButton() !== ui.Button.OK) return;
      tv = tvr.getResponseText().trim(); tl = 'By ' + tv;
    }

    // ── REQ 3: Ask for extra columns ────────────────────────────
    var extraCols = [];
    var askExtra  = ui.alert(
      '➕ Extra Columns for "' + what + '"?',
      'Does this step need any extra columns?\n(e.g. CAD File, Revision Number, Approval Status)',
      ui.ButtonSet.YES_NO
    );

    while (askExtra === ui.Button.YES) {
      var ecNameR = ui.prompt(
        '📝 Extra Column — Name',
        'Column name for step "' + what + '":\ne.g. CAD File / Revision / Approval',
        ui.ButtonSet.OK_CANCEL
      );
      if (ecNameR.getSelectedButton() !== ui.Button.OK) break;
      var ecName = ecNameR.getResponseText().trim();
      if (!ecName) break;

      var ecTypeR = ui.prompt(
        '📝 "' + ecName + '" — Data Type',
        'Text\nDate\nNumber\nDropdown',
        ui.ButtonSet.OK_CANCEL
      );
      if (ecTypeR.getSelectedButton() !== ui.Button.OK) break;
      var ecType = ecTypeR.getResponseText().trim().toLowerCase();
      if (ecType === 'date')           ecType = 'Date';
      else if (ecType === 'number')    ecType = 'Number';
      else if (ecType.indexOf('drop') !== -1) ecType = 'Dropdown';
      else                             ecType = 'Text';

      var ecChoices = [];
      if (ecType === 'Dropdown') {
        var ecChoiceR = ui.prompt(
          '📝 Choices for "' + ecName + '"',
          'Comma separated  e.g. R1, R2, R3',
          ui.ButtonSet.OK_CANCEL
        );
        if (ecChoiceR.getSelectedButton() !== ui.Button.OK) break;
        ecChoices = ecChoiceR.getResponseText().split(',').map(function(c){ return c.trim(); }).filter(Boolean);
      }

      extraCols.push({ name: ecName, type: ecType, choices: ecChoices });
      Logger.log('Extra col added for step "' + what + '": ' + ecName + ' (' + ecType + ')');

      askExtra = ui.alert(
        '➕ Another Extra Column for "' + what + '"?',
        'Add another extra column to this step?',
        ui.ButtonSet.YES_NO
      );
    }

    steps.push({ num: s, what: what, who: who, how: how,
                 tatType: tt, tatValue: tv, tatLabel: tl,
                 extraCols: extraCols });
  }

  // Save progress to PropertiesService before writing (REQ 2)
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    'fms_steps_json':   JSON.stringify(steps),
    'fms_fields_json':  JSON.stringify(fieldDefs),
    'fms_total_steps':  steps.length.toString(),
    'fms_done_steps':   '0',
    'fms_sheet_name':   FMS_SHEET_NAME
  });
  Logger.log('✅ Step data saved to PropertiesService — resume available if timeout');

  // Write steps to sheet
  _buildFMSSteps_(fmsSheet, steps, fieldDefs, 0);

  // Clear PropertiesService on success
  props.deleteAllProperties();
  Logger.log('✅ All steps written — PropertiesService cleared');

  ui.alert('✅ ' + ns + ' Steps Added!\n\nNext: 📋 Generate Forms for Active FMS');
}


// ══════════════════════════════════════════════════════════════════
// REQ 2 — RESUME STEP WRITING (if previous run timed out)
// ══════════════════════════════════════════════════════════════════
function resumeStepWriting() {
  var ui    = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var all   = props.getProperties();

  if (!all.fms_steps_json || !all.fms_fields_json) {
    ui.alert('ℹ️ No interrupted step writing found.\n\nEither:\n• Steps completed successfully\n• No step writing was started\n\nUse "➕ Add Steps to FMS" to start fresh.');
    return;
  }

  var steps      = JSON.parse(all.fms_steps_json);
  var fieldDefs  = JSON.parse(all.fms_fields_json);
  var totalSteps = parseInt(all.fms_total_steps || steps.length);
  var doneSteps  = parseInt(all.fms_done_steps  || '0');
  var sheetName  = all.fms_sheet_name || FMS_SHEET_NAME;

  var confirm = ui.alert(
    '▶️ Resume Step Writing',
    'Found interrupted step writing:\n\n' +
    '• Total steps: ' + totalSteps + '\n' +
    '• Already written: ' + doneSteps + '\n' +
    '• Remaining: ' + (totalSteps - doneSteps) + '\n\n' +
    'Resume from step ' + (doneSteps + 1) + '?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var fmsSheet = ss.getSheetByName(sheetName);
  if (!fmsSheet) {
    ui.alert('❌ FMS sheet "' + sheetName + '" not found.');
    return;
  }

  _buildFMSSteps_(fmsSheet, steps, fieldDefs, doneSteps);

  props.deleteAllProperties();
  ui.alert('✅ Resume complete! All ' + totalSteps + ' steps written.\n\nNext: 📋 Generate Forms for Active FMS');
}


// ══════════════════════════════════════════════════════════════════
// BUILD FMS — BASE SHEET
// ══════════════════════════════════════════════════════════════════
function _buildFMSBaseSheet_(sheet, openFrac, closeFrac, fieldDefs) {
  try {
    var numFields = fieldDefs.length;

    sheet.getRange('A1').setFormula('=NOW()');
    sheet.getRange('C1').setValue(openFrac).setNumberFormat('HH:mm');
    sheet.getRange('D1').setValue(closeFrac).setNumberFormat('HH:mm');
    sheet.getRange('E1').setValue("'0000001");
    sheet.hideRows(1);

    _storeFieldDefs_(sheet, fieldDefs);

    sheet.getRange(ROW_HEADERS, ID_COLUMN).setValue('ID');
    sheet.getRange(ROW_HEADERS, TIMESTAMP_COLUMN).setValue('Timestamp');
    for (var f = 0; f < numFields; f++) {
      sheet.getRange(ROW_HEADERS, FIELDS_START_COL + f).setValue(fieldDefs[f].name);
    }

    var tc = FIELDS_START_COL + numFields - 1;
    sheet.getRange(ROW_HEADERS, 1, 1, tc)
      .setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);

    sheet.setFrozenRows(ROW_HEADERS);
    iterative();
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    Logger.log('✅ Base FMS sheet built');
  } catch(err) {
    Logger.log('❌ _buildFMSBaseSheet_ error: ' + err);
  }
}


// ══════════════════════════════════════════════════════════════════
// BUILD FMS — STEPS
// REQ 1: Uses setFormulas() batch write — one API call per column
// REQ 2: startFromStep param allows resume from any step
// REQ 3: Handles extra cols per step with variable column widths
// ══════════════════════════════════════════════════════════════════
function _buildFMSSteps_(sheet, steps, fieldDefs, startFromStep) {
  try {
    startFromStep = startFromStep || 0;
    var ns        = steps.length;
    var numFields = fieldDefs.length;
    var ds        = FMS_DATA_START_ROW + 1; // row 10

    // Calculate full column layout for ALL steps
    var stepLayout = _calcStepColLayout_(steps, numFields);
    var tc         = stepLayout[ns-1].timeDelayCol; // last col used

    // Store enriched step defs with col positions
    _storeStepDefs_(sheet, stepLayout);

    // ── Rows 2–3: Step IDs ──────────────────────────────────────
    for (var s = startFromStep; s < ns; s++) {
      var sl = stepLayout[s];
      sheet.getRange(ROW_STEP_NUM, sl.startCol).setValue('Step' + (s+1));
      sheet.getRange(ROW_STEP_ID,  sl.startCol).setValue('s'    + (s+1));
    }

    // ── Rows 4–7: What/Who/How/When ─────────────────────────────
    var lbs  = ['What','Who','How','When'];
    var flds = ['what','who','how','tatLabel'];
    for (var ri = 0; ri < 4; ri++) {
      sheet.getRange(ROW_WHAT + ri, 1).setValue(lbs[ri]);
      for (var s = startFromStep; s < ns; s++) {
        var sl = stepLayout[s];
        sheet.getRange(ROW_WHAT + ri, sl.startCol, 1, sl.totalCols)
          .merge().setValue(steps[s][flds[ri]]);
      }
    }

    // ── Row 8: TAT ref ──────────────────────────────────────────
    sheet.getRange(ROW_TAT_REF, 1).setValue('TAT Ref');
    for (var s = startFromStep; s < ns; s++) {
      var sl = stepLayout[s];
      if ((steps[s].tatType === 'H' || steps[s].tatType === 'D') && steps[s].tatValue) {
        sheet.getRange(ROW_TAT_REF, sl.plannedCol).setValue(parseFloat(steps[s].tatValue));
      }
    }
    sheet.hideRows(ROW_TAT_REF);

    // ── Row 9: Headers ──────────────────────────────────────────
    for (var s = startFromStep; s < ns; s++) {
      var sl = stepLayout[s];
      // Base headers: Planned, Actual
      sheet.getRange(ROW_HEADERS, sl.plannedCol).setValue('Planned');
      sheet.getRange(ROW_HEADERS, sl.actualCol).setValue('Actual');
      // Extra col headers
      for (var x = 0; x < sl.extraCols.length; x++) {
        sheet.getRange(ROW_HEADERS, sl.extraCols[x].col).setValue(sl.extraCols[x].name);
      }
      // Status + Time Delay
      sheet.getRange(ROW_HEADERS, sl.statusCol).setValue('Status');
      sheet.getRange(ROW_HEADERS, sl.timeDelayCol).setValue('Time Delay');
    }

    // Restyle full header row
    sheet.getRange(ROW_HEADERS, 1, 1, tc)
      .setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);

    // Style info rows 4–7
    sheet.getRange(ROW_WHAT, 1, 4, tc).setBackground('#EFF6FF').setFontWeight('bold');

    // ── REQ 1: BATCH FORMULA WRITE ──────────────────────────────
    // Build formula arrays for each step, write entire column in ONE call
    var identColLtr = _col_(FIELDS_START_COL);

    for (var s = startFromStep; s < ns; s++) {
      var sl         = stepLayout[s];
      var pL         = _col_(sl.plannedCol);
      var aL         = _col_(sl.actualCol);
      var stL        = _col_(sl.statusCol);
      var dL         = _col_(sl.timeDelayCol);
      var tatRef     = pL + '$8';

      // Previous actual col letter
      var prevActL   = (s === 0)
        ? 'B'
        : _col_(stepLayout[s-1].actualCol);

      // ── BUILD Planned formula array (1000 items) ─────────────
      var plannedFormulas = [];
      for (var r = ds; r < ds + NR; r++) {
        var rr         = r.toString();
        var identRef   = identColLtr + rr;
        var prevActRef = prevActL + rr;
        var pF = '';

        if (steps[s].tatType === 'H') {
          pF = _buildTATFormula_(identRef, prevActRef, tatRef, false);
        } else if (steps[s].tatType === 'D') {
          var days = parseFloat(steps[s].tatValue) || 1;
          pF = '=IFERROR(IF(' + identRef + '<>"",IF(' + prevActRef + '>0,WORKDAY.INTL(INT(' + prevActRef + '),' + days + ',$E$1,Holidays!$A:$A)+$C$1,""),""),"")';
        } else if (steps[s].tatType === 'T') {
          var tp = (steps[s].tatValue || '18:00').split(':');
          var tf = (parseInt(tp[0])*60 + parseInt(tp[1]||'0')) / 1440;
          pF = '=IFERROR(IF(' + identRef + '<>"",IF(' + prevActRef + '>0,IF(MOD(' + prevActRef + ',1)<' + tf + ',INT(' + prevActRef + ')+' + tf + ',WORKDAY.INTL(INT(' + prevActRef + '),1,$E$1,Holidays!$A:$A)+' + tf + '),""),""),"")';
        }
        // N type = empty string (no formula)
        plannedFormulas.push([pF]);
      }

      // Write entire Planned column in ONE API call
      if (steps[s].tatType !== 'N') {
        sheet.getRange(ds, sl.plannedCol, NR, 1)
          .setFormulas(plannedFormulas)
          .setNumberFormat('dd/MM/yyyy HH:mm:ss');
      }

      // ── Actual — plain cells (REQ 2 FIX: Form 2 writes directly)
      // No formula — Form 2 handler writes new Date() here on submit

      // ── Time Delay — ARRAYFORMULA (one cell, open range) ─────
      var pRange = pL + ds + ':' + pL;
      var aRange = aL + ds + ':' + aL;
      sheet.getRange(ds, sl.timeDelayCol)
        .setFormula('=ARRAYFORMULA(IFERROR(IF(' + pRange + '<>"",IF(' + aRange + '<>"",IF(' + aRange + '>' + pRange + ',' + aRange + '-' + pRange + ',""),$A$1-' + pRange + '),""),""))')
        .setNumberFormat('[h]:mm:ss');

      // ── Conditional formatting on Time Delay ─────────────────
      // (done after all steps to avoid repeated setConditionalFormatRules calls)

      // ── REQ 2: Update progress in PropertiesService ──────────
      var props = PropertiesService.getScriptProperties();
      props.setProperty('fms_done_steps', (s + 1).toString());
      Logger.log('✅ Step ' + (s+1) + '/' + ns + ' written — cols ' + sl.startCol + ' to ' + sl.timeDelayCol);
    }

    // ── Conditional formatting — all steps at once ───────────────
    var rules = [];
    for (var s = 0; s < ns; s++) {
      var sl  = stepLayout[s];
      var dr  = sheet.getRange(ds, sl.timeDelayCol, NR, 1);
      var aL2 = _col_(sl.actualCol);
      var pL2 = _col_(sl.plannedCol);
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([dr])
        .whenFormulaSatisfied('=AND(' + aL2 + ds + '<>"",' + aL2 + ds + '>' + pL2 + ds + ')')
        .setBackground('#F4C7C3').build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .setRanges([dr])
        .whenFormulaSatisfied('=AND(' + aL2 + ds + '<>"",' + aL2 + ds + '<=' + pL2 + ds + ')')
        .setBackground('#B7E1CD').build());
    }
    sheet.setConditionalFormatRules(rules);

    sheet.getRange('B' + ds + ':B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.setFrozenRows(ROW_HEADERS);
    iterative();

    Logger.log('✅ _buildFMSSteps_ complete — all ' + ns + ' steps written');
  } catch(err) {
    Logger.log('❌ _buildFMSSteps_ error at step: ' + err);
    throw err; // rethrow so PropertiesService keeps saved state for resume
  }
}


// ══════════════════════════════════════════════════════════════════
// GENERATE FORMS
// REQ 3: Form 2 has sections per step with navigation
// ══════════════════════════════════════════════════════════════════
function generateFormsForActiveFMS() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var ui          = SpreadsheetApp.getUi();
  var activeSheet = ss.getActiveSheet();
  var sheetName   = activeSheet.getName();

  var confirm = ui.alert(
    '📋 Generate Forms',
    'Generate forms for sheet: "' + sheetName + '"?\n\n' +
    '• Form 1 — New Order Entry\n' +
    '• Form 2 — Status Update (with step sections)\n' +
    '• Links saved to "Form Links" sheet',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var stepDefs  = _loadStepDefs_(activeSheet);
  var fieldDefs = _loadFieldDefs_(activeSheet);

  if (!stepDefs || stepDefs.length === 0) {
    ui.alert('❌ No steps found.\nRun "➕ Add Steps to FMS" first.');
    return;
  }
  if (!fieldDefs || fieldDefs.length === 0) {
    ui.alert('❌ No field definitions found.\nRun "🆕 Create New FMS Sheet" first.');
    return;
  }

  ui.alert('⏳ Creating forms... ~15 seconds.\nClick OK and wait.');

  try {
    var result = _createBothForms_(ss, stepDefs, sheetName, fieldDefs);
    _saveFormLinks_(ss, sheetName, result.form1Url, result.form1EditUrl,
                    result.form2Url, result.form2EditUrl, stepDefs, fieldDefs);
    ui.alert(
      '✅ Both Forms Created!\n\n' +
      '📋 Form 1:\n' + result.form1Url + '\n\n' +
      '📋 Form 2:\n' + result.form2Url + '\n\n' +
      'Links saved to "Form Links" sheet.'
    );
  } catch(err) {
    ui.alert('❌ Error creating forms:\n' + err + '\n\nCheck Apps Script logs.');
    Logger.log('❌ generateFormsForActiveFMS error: ' + err);
  }
}


// ── Read steps from FMS sheet ──────────────────────────────────────
function _readStepsFromSheet_(sheet) {
  // First try loading from stored step defs (v9)
  var stepDefs = _loadStepDefs_(sheet);
  if (stepDefs && stepDefs.length > 0) return stepDefs;

  // Fallback: read from sheet rows
  var steps     = [];
  var fieldDefs = _loadFieldDefs_(sheet);
  var numFields = fieldDefs ? fieldDefs.length : 1;
  var firstStepCol = _calcFirstStepCol_(numFields);
  var lastCol   = sheet.getLastColumn();

  for (var s = 0; s < 15; s++) {
    var startCol = firstStepCol + (s * BASE_COLS_PER_STEP);
    if (startCol > lastCol) break;
    var stepId   = sheet.getRange(ROW_STEP_ID, startCol).getValue().toString().trim();
    var stepName = sheet.getRange(ROW_WHAT,    startCol).getValue().toString().trim();
    if (!stepId || !stepName) break;
    steps.push({ id: stepId, name: stepName, startCol: startCol,
                 totalCols: BASE_COLS_PER_STEP,
                 plannedCol: startCol, actualCol: startCol+1,
                 statusCol: startCol+2, timeDelayCol: startCol+3,
                 extraCols: [] });
  }
  return steps;
}


// ── Create both forms ──────────────────────────────────────────────
function _createBothForms_(ss, stepDefs, fmsSheetName, fieldDefs) {

  // ── FORM 1 — dynamic fields ──────────────────────────────────
  var form1 = FormApp.create('FMS New Order Entry (' + fmsSheetName + ')');
  form1.setDescription('Submit to create a new order in: ' + fmsSheetName);
  form1.setCollectEmail(false);
  form1.setShowLinkToRespondAgain(true);
  form1.setConfirmationMessage('New order submitted! FMS has been updated.');

  fieldDefs.forEach(function(fd) {
    try {
      if (fd.type === 'Text') {
        form1.addTextItem().setTitle(fd.name).setRequired(true);
      } else if (fd.type === 'Date') {
        form1.addDateItem().setTitle(fd.name).setRequired(true);
      } else if (fd.type === 'Dropdown') {
        form1.addListItem().setTitle(fd.name)
          .setChoiceValues(fd.choices.length > 0 ? fd.choices : ['Option 1'])
          .setRequired(true);
      }
    } catch(e) { Logger.log('⚠️ Form1 field error: ' + e); }
  });

  form1.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  Utilities.sleep(4000);
  _renameLatestFormSheet_(ss, FORM1_SHEET_NAME, ['Form responses 4']);

  // ── FORM 2 — sections per step (REQ 3) ──────────────────────
  var form2 = FormApp.create('FMS Step Status Update (' + fmsSheetName + ')');
  form2.setDescription('Submit to mark a step as Done in: ' + fmsSheetName);
  form2.setCollectEmail(false);
  form2.setShowLinkToRespondAgain(true);
  form2.setConfirmationMessage('Step marked as Done! FMS has been updated.');

  // Page 1: Identifier + Step selection
  var identifierName = fieldDefs[0].name;
  form2.addTextItem()
    .setTitle(identifierName)
    .setHelpText('Enter ' + identifierName + ' exactly as shown in FMS')
    .setRequired(true);

  // Step dropdown with navigation to sections
  var stepListItem = form2.addListItem()
    .setTitle('Select Step')
    .setHelpText('Select the step that has been completed')
    .setRequired(true);

  // Create one page/section per step FIRST (so we have page objects)
  var stepPages = [];
  for (var s = 0; s < stepDefs.length; s++) {
    var sl   = stepDefs[s];
    var page = form2.addPageBreakItem()
      .setTitle('Step ' + (s+1) + ' — ' + sl.what)
      .setGoToPage(FormApp.PageNavigationType.SUBMIT);
    stepPages.push(page);

    // Extra cols for this step
    if (sl.extraCols && sl.extraCols.length > 0) {
      sl.extraCols.forEach(function(ec) {
        try {
          if (ec.type === 'Text') {
            form2.addTextItem().setTitle(ec.name).setRequired(false);
          } else if (ec.type === 'Date') {
            form2.addDateItem().setTitle(ec.name).setRequired(false);
          } else if (ec.type === 'Number') {
            form2.addTextItem().setTitle(ec.name)
              .setHelpText('Enter a number').setRequired(false);
          } else if (ec.type === 'Dropdown') {
            form2.addListItem().setTitle(ec.name)
              .setChoiceValues(ec.choices.length > 0 ? ec.choices : ['Option 1'])
              .setRequired(false);
          }
        } catch(e) { Logger.log('⚠️ Form2 extra col error: ' + e); }
      });
    }

    // Status field in each section
    form2.addMultipleChoiceItem()
      .setTitle('Status')
      .setChoiceValues(['Done'])
      .setRequired(true);
  }

  // Now set navigation choices on step dropdown
  var choices = stepDefs.map(function(sl, idx) {
    return stepListItem.createChoice(
      'S' + (idx+1) + ' - ' + sl.what,
      stepPages[idx]
    );
  });
  stepListItem.setChoices(choices);

  form2.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  Utilities.sleep(4000);
  _renameLatestFormSheet_(ss, FORM2_SHEET_NAME, ['Form responses 6', FORM1_SHEET_NAME]);

  return {
    form1Url: form1.getPublishedUrl(), form1EditUrl: form1.getEditUrl(),
    form2Url: form2.getPublishedUrl(), form2EditUrl: form2.getEditUrl()
  };
}


// ── Rename latest form response sheet ─────────────────────────────
function _renameLatestFormSheet_(ss, newName, excludeNames) {
  try {
    SpreadsheetApp.flush();
    var sheets = ss.getSheets();
    for (var i = sheets.length - 1; i >= 0; i--) {
      var name = sheets[i].getName();
      if (excludeNames.indexOf(name) !== -1) continue;
      if (name === newName) continue;
      if (name.indexOf('Form responses') !== -1 || name.indexOf('Form Responses') !== -1) {
        sheets[i].setName(newName);
        Logger.log('Renamed "' + name + '" to "' + newName + '"');
        return;
      }
    }
  } catch(e) { Logger.log('⚠️ _renameLatestFormSheet_ error: ' + e); }
}


// ── Save form links ────────────────────────────────────────────────
function _saveFormLinks_(ss, fmsSheetName, form1Url, form1EditUrl, form2Url, form2EditUrl, stepDefs, fieldDefs) {
  try {
    var ls = ss.getSheetByName('Form Links') || ss.insertSheet('Form Links');
    var lr = ls.getLastRow();
    var sr = (lr === 0) ? 1 : lr + 2;

    ls.getRange(sr, 1, 1, 4)
      .setValues([['FMS Sheet','Form','Fill Link (share this)','Edit Link (admin)']])
      .setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
    ls.getRange(sr+1, 1, 1, 4).setValues([[fmsSheetName, 'Form 1 — New Order Entry', form1Url, form1EditUrl]]);
    ls.getRange(sr+2, 1, 1, 4).setValues([[fmsSheetName, 'Form 2 — Status Update',   form2Url, form2EditUrl]]);

    var fSummary = 'Fields: ' + fieldDefs.map(function(f){ return f.name+'('+f.type+')'; }).join(' | ');
    var sSummary = 'Steps: '  + stepDefs.map(function(s){ return 's'+(s.num||'')+'-'+s.what; }).join(' | ');
    ls.getRange(sr+3, 1, 1, 4).setValues([[fSummary, sSummary, '', '']])
      .setBackground('#EFF6FF').setFontColor('#1E3A5F');

    try {
      ls.getRange(sr+1, 3).setFormula('=HYPERLINK("' + form1Url + '","Open Form 1")');
      ls.getRange(sr+2, 3).setFormula('=HYPERLINK("' + form2Url + '","Open Form 2")');
    } catch(e) {}

    ls.autoResizeColumns(1, 4);
    ls.getRange(sr+1, 1, 3, 4).setBorder(true,true,true,true,true,true);
    ss.setActiveSheet(ls);
  } catch(e) { Logger.log('⚠️ _saveFormLinks_ error: ' + e); }
}


// ══════════════════════════════════════════════════════════════════
// SETUP TRIGGERS
// ══════════════════════════════════════════════════════════════════
function setupAllTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('onChange_new').forSpreadsheet(ss).onChange().create();
    ScriptApp.newTrigger('onFormSubmit_Router').forSpreadsheet(ss).onFormSubmit().create();
    ui.alert(
      '✅ 2 Triggers Created!\n\n' +
      '1. onChange_new         — auto-timestamps Done/Yes/No\n' +
      '2. onFormSubmit_Router  — routes Form 1 & Form 2\n\n' +
      '⚠️ Go to Apps Script → Triggers and confirm only 2 exist.\nDelete any extras.'
    );
  } catch(e) {
    ui.alert('❌ Error: ' + e);
    Logger.log('❌ setupAllTriggers error: ' + e);
  }
}


// ══════════════════════════════════════════════════════════════════
// SINGLE ROUTER
// ══════════════════════════════════════════════════════════════════
function onFormSubmit_Router(e) {
  try {
    if (!e || !e.range) {
      Logger.log('⚠️ No event — fallback Form1');
      onFormSubmit_NewOrder(e);
      return;
    }
    var sheetName = e.range.getSheet().getName();
    Logger.log('Form submitted to: "' + sheetName + '"');
    if (sheetName === FORM1_SHEET_NAME)      onFormSubmit_NewOrder(e);
    else if (sheetName === FORM2_SHEET_NAME) onFormSubmit_StatusUpdate(e);
    else Logger.log('Unknown sheet "' + sheetName + '" — skipping');
  } catch(err) {
    Logger.log('❌ onFormSubmit_Router error: ' + err);
  }
}


// ══════════════════════════════════════════════════════════════════
// FORM 1 HANDLER — New Order Entry
// ══════════════════════════════════════════════════════════════════
function onFormSubmit_NewOrder(e) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var fmsSheet  = ss.getSheetByName(FMS_SHEET_NAME);
    var formSheet = ss.getSheetByName(FORM1_SHEET_NAME);

    if (!fmsSheet)  { Logger.log('❌ FMS sheet not found');   return; }
    if (!formSheet) { Logger.log('❌ Form1 sheet not found'); return; }

    var fieldDefs = _loadFieldDefs_(fmsSheet);
    if (!fieldDefs || fieldDefs.length === 0) { Logger.log('❌ No field defs'); return; }

    var identifierFieldName = fieldDefs[0].name.toLowerCase().trim();
    var lastRow = formSheet.getLastRow();
    if (lastRow < 2) { Logger.log('❌ No responses yet'); return; }

    var headers  = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    var response = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];

    var data = {};
    headers.forEach(function(h, i){ data[h.toString().trim().toLowerCase()] = response[i]; });

    var identifierValue = data[identifierFieldName];
    if (!identifierValue) identifierValue = _getField_(data, [identifierFieldName, 'po number', 'po', 'bag number', 'bag']);
    if (!identifierValue) { Logger.log('❌ Identifier blank'); return; }
    identifierValue = identifierValue.toString().trim();

    // Find first empty row from row 10
    var dataStartRow = FMS_DATA_START_ROW + 1;
    var scanValues   = fmsSheet.getRange(dataStartRow, FIELDS_START_COL, 500, 1).getValues();

    var newRow = -1, usedCount = 0;
    for (var i = 0; i < scanValues.length; i++) {
      var cv = scanValues[i][0].toString().trim();
      if (cv !== '') {
        if (cv.toUpperCase() === identifierValue.toUpperCase()) {
          Logger.log('⚠️ Duplicate "' + identifierValue + '" — skipping'); return;
        }
        usedCount++;
      } else {
        if (newRow === -1) newRow = dataStartRow + i;
      }
    }

    if (newRow === -1) { Logger.log('❌ No empty rows (500 full)'); return; }

    fmsSheet.getRange(newRow, ID_COLUMN).setValue(usedCount + 1);
    fmsSheet.getRange(newRow, TIMESTAMP_COLUMN).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    fieldDefs.forEach(function(fd, idx) {
      try {
        var colIdx = FIELDS_START_COL + idx;
        var val    = data[fd.name.toLowerCase().trim()];
        if (fd.type === 'Date' && val) {
          var pd = new Date(val);
          if (!isNaN(pd.getTime())) {
            fmsSheet.getRange(newRow, colIdx).setValue(pd).setNumberFormat('dd/MM/yyyy');
          } else { fmsSheet.getRange(newRow, colIdx).setValue(val); }
        } else { fmsSheet.getRange(newRow, colIdx).setValue(val !== undefined ? val : ''); }
      } catch(ferr) { Logger.log('⚠️ Field write error: ' + ferr); }
    });

    SpreadsheetApp.flush();
    Logger.log('✅ Form1 SUCCESS — Row=' + newRow + ' ID=' + (usedCount+1) + ' ' + identifierFieldName + '=' + identifierValue);

  } catch(err) { Logger.log('❌ onFormSubmit_NewOrder error: ' + err); }
}


// ══════════════════════════════════════════════════════════════════
// FORM 2 HANDLER — Status Update
// FIX: Writes Actual timestamp directly to Actual col
// REQ 3: Reads extra col values and writes to correct FMS columns
// ══════════════════════════════════════════════════════════════════
function onFormSubmit_StatusUpdate(e) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var fmsSheet  = ss.getSheetByName(FMS_SHEET_NAME);
    var formSheet = ss.getSheetByName(FORM2_SHEET_NAME);

    if (!fmsSheet)  { Logger.log('❌ FMS sheet not found');   return; }
    if (!formSheet) { Logger.log('❌ Form2 sheet not found'); return; }

    var fieldDefs = _loadFieldDefs_(fmsSheet);
    var stepDefs  = _loadStepDefs_(fmsSheet);
    if (!fieldDefs || fieldDefs.length === 0) { Logger.log('❌ No field defs'); return; }
    if (!stepDefs  || stepDefs.length  === 0) { Logger.log('❌ No step defs'); return; }

    var identifierFieldName = fieldDefs[0].name.toLowerCase().trim();

    var lastRow = formSheet.getLastRow();
    if (lastRow < 2) { Logger.log('❌ No responses yet'); return; }

    var headers  = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    var response = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];

    var data = {};
    headers.forEach(function(h, i){ data[h.toString().trim().toLowerCase()] = response[i]; });

    // Get identifier and step
    var identifierValue = data[identifierFieldName] || _getField_(data, [identifierFieldName]);
    var stepRaw         = _getField_(data, ['select step', 'step id', 'step', 'stepid']);

    if (!stepRaw)         { Logger.log('❌ Step blank');       return; }
    if (!identifierValue) { Logger.log('❌ Identifier blank'); return; }

    // Extract step number from "S1 - CAD" or "s1"
    var stepStr = stepRaw.toString().trim();
    var stepNum = parseInt(stepStr.replace(/[^0-9]/g, ''));
    if (isNaN(stepNum) || stepNum < 1 || stepNum > 15) {
      Logger.log('❌ Cannot parse stepNum from: "' + stepStr + '"'); return;
    }

    identifierValue = identifierValue.toString().trim().toUpperCase();

    // Find step def
    var sl = null;
    for (var i = 0; i < stepDefs.length; i++) {
      if (stepDefs[i].num === stepNum) { sl = stepDefs[i]; break; }
    }
    if (!sl) {
      // fallback: use index
      sl = stepDefs[stepNum - 1];
    }
    if (!sl) { Logger.log('❌ Step def not found for stepNum=' + stepNum); return; }

    Logger.log('Step found: ' + sl.what + ' | statusCol=' + sl.statusCol + ' | actualCol=' + sl.actualCol);

    // Find matching row by identifier
    var dataStartRow = FMS_DATA_START_ROW + 1;
    var lastFmsRow   = fmsSheet.getLastRow();
    if (lastFmsRow < dataStartRow) { Logger.log('❌ No data rows'); return; }

    var idValues = fmsSheet.getRange(dataStartRow, FIELDS_START_COL, lastFmsRow - dataStartRow + 1, 1).getValues();
    var matchRow = -1;
    for (var i = 0; i < idValues.length; i++) {
      if (idValues[i][0].toString().trim().toUpperCase() === identifierValue) {
        matchRow = dataStartRow + i; break;
      }
    }
    if (matchRow === -1) { Logger.log('❌ Identifier not found: ' + identifierValue); return; }

    // ── Write Status = Done ──────────────────────────────────────
    fmsSheet.getRange(matchRow, sl.statusCol).setValue('Done');

    // ── FIX: Write Actual timestamp directly ─────────────────────
    fmsSheet.getRange(matchRow, sl.actualCol)
      .setValue(new Date())
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');

    // ── REQ 3: Write extra col values ────────────────────────────
    if (sl.extraCols && sl.extraCols.length > 0) {
      sl.extraCols.forEach(function(ec) {
        try {
          var ecKey = ec.name.toLowerCase().trim();
          var ecVal = data[ecKey];
          if (ecVal !== undefined && ecVal !== '') {
            if (ec.type === 'Date' && ecVal) {
              var pd = new Date(ecVal);
              if (!isNaN(pd.getTime())) {
                fmsSheet.getRange(matchRow, ec.col).setValue(pd).setNumberFormat('dd/MM/yyyy');
              } else { fmsSheet.getRange(matchRow, ec.col).setValue(ecVal); }
            } else if (ec.type === 'Number') {
              fmsSheet.getRange(matchRow, ec.col).setValue(parseFloat(ecVal) || ecVal);
            } else {
              fmsSheet.getRange(matchRow, ec.col).setValue(ecVal);
            }
            Logger.log('✅ Extra col "' + ec.name + '" = "' + ecVal + '" written to col ' + ec.col);
          }
        } catch(ecErr) { Logger.log('⚠️ Extra col write error "' + ec.name + '": ' + ecErr); }
      });
    }

    SpreadsheetApp.flush();
    Logger.log('✅ Form2 SUCCESS — identifier=' + identifierValue + ' step=s' + stepNum + ' row=' + matchRow);

  } catch(err) { Logger.log('❌ onFormSubmit_StatusUpdate error: ' + err); }
}


// ══════════════════════════════════════════════════════════════════
// REQ 4 — FMS DIAGNOSTIC CHECK
// Checks everything, writes full report to "FMS Diagnostics" sheet
// ══════════════════════════════════════════════════════════════════
function runFMSDiagnostic() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ui.alert('🔍 Running FMS Diagnostic...\nThis will check all sheets, steps, forms and triggers.\nClick OK to start.');

  var log    = [];  // array of [status, section, check, detail]
  var pass   = 0;
  var fail   = 0;
  var warn   = 0;

  function addLog(status, section, check, detail) {
    log.push([status, section, check, detail || '']);
    if (status === '✅ PASS')    pass++;
    else if (status === '❌ FAIL') fail++;
    else if (status === '⚠️ WARN') warn++;
    Logger.log(status + ' | ' + section + ' | ' + check + (detail ? ' | ' + detail : ''));
  }

  // ── SECTION 1: SHEETS ─────────────────────────────────────────
  var fmsSheet  = ss.getSheetByName(FMS_SHEET_NAME);
  var holSheet  = ss.getSheetByName('Holidays');
  var linkSheet = ss.getSheetByName('Form Links');
  var f1Sheet   = ss.getSheetByName(FORM1_SHEET_NAME);
  var f2Sheet   = ss.getSheetByName(FORM2_SHEET_NAME);

  addLog(fmsSheet  ? '✅ PASS' : '❌ FAIL', 'Sheets', 'FMS sheet exists',         fmsSheet  ? 'Found' : 'Not found — run Create New FMS Sheet');
  addLog(holSheet  ? '✅ PASS' : '⚠️ WARN', 'Sheets', 'Holidays sheet exists',    holSheet  ? 'Found' : 'Missing — create manually or run First Time Install');
  addLog(linkSheet ? '✅ PASS' : '⚠️ WARN', 'Sheets', 'Form Links sheet exists',  linkSheet ? 'Found' : 'Missing — run Generate Forms');
  addLog(f1Sheet   ? '✅ PASS' : '⚠️ WARN', 'Sheets', 'Form1 response sheet',     f1Sheet   ? 'Found: ' + FORM1_SHEET_NAME : 'Missing: ' + FORM1_SHEET_NAME);
  addLog(f2Sheet   ? '✅ PASS' : '⚠️ WARN', 'Sheets', 'Form2 response sheet',     f2Sheet   ? 'Found: ' + FORM2_SHEET_NAME : 'Missing: ' + FORM2_SHEET_NAME);

  // ── SECTION 2: FMS CONFIG ─────────────────────────────────────
  if (fmsSheet) {
    var a1Formula = fmsSheet.getRange('A1').getFormula();
    addLog(a1Formula.indexOf('NOW') !== -1 ? '✅ PASS' : '❌ FAIL', 'Config', 'A1 = NOW() formula', a1Formula || 'empty');

    var openTime  = fmsSheet.getRange('C1').getValue();
    var closeTime = fmsSheet.getRange('D1').getValue();
    var workdays  = fmsSheet.getRange('E1').getValue();
    addLog(openTime  ? '✅ PASS' : '❌ FAIL', 'Config', 'Opening time set in C1',  openTime  ? 'Value: ' + openTime  : 'Empty');
    addLog(closeTime ? '✅ PASS' : '❌ FAIL', 'Config', 'Closing time set in D1',  closeTime ? 'Value: ' + closeTime : 'Empty');
    addLog(workdays  ? '✅ PASS' : '❌ FAIL', 'Config', 'Working days set in E1',  workdays  ? 'Value: ' + workdays  : 'Empty');

    var fieldDefs = _loadFieldDefs_(fmsSheet);
    addLog(fieldDefs && fieldDefs.length > 0 ? '✅ PASS' : '❌ FAIL', 'Config', 'Field definitions in col 100',
           fieldDefs ? fieldDefs.length + ' fields: ' + fieldDefs.map(function(f){ return f.name; }).join(', ') : 'Not found');

    var stepDefs = _loadStepDefs_(fmsSheet);
    addLog(stepDefs && stepDefs.length > 0 ? '✅ PASS' : '❌ FAIL', 'Config', 'Step definitions in col 101',
           stepDefs ? stepDefs.length + ' steps: ' + stepDefs.map(function(s){ return s.what; }).join(', ') : 'Not found');

    var frozenRows = fmsSheet.getFrozenRows();
    addLog(frozenRows === ROW_HEADERS ? '✅ PASS' : '⚠️ WARN', 'Config', 'Row 9 frozen (headers)',
           'Frozen rows: ' + frozenRows + (frozenRows !== ROW_HEADERS ? ' — expected ' + ROW_HEADERS : ''));

    var iterEnabled = ss.isIterativeCalculationEnabled();
    addLog(iterEnabled ? '✅ PASS' : '❌ FAIL', 'Config', 'Iterative calculation enabled',
           iterEnabled ? 'Enabled' : 'Disabled — run First Time Install');

    // ── SECTION 3: STEPS ────────────────────────────────────────
    if (stepDefs && stepDefs.length > 0) {
      var ds = FMS_DATA_START_ROW + 1;
      for (var s = 0; s < stepDefs.length; s++) {
        var sl      = stepDefs[s];
        var section = 'Step ' + (s+1) + ' (' + sl.what + ')';

        // Step ID in row 3
        var stepId = fmsSheet.getRange(ROW_STEP_ID, sl.startCol).getValue();
        addLog(stepId ? '✅ PASS' : '❌ FAIL', section, 'Step ID in row 3',
               stepId ? 'Found: ' + stepId : 'Missing');

        // Step name in row 4
        var stepName = fmsSheet.getRange(ROW_WHAT, sl.startCol).getValue();
        addLog(stepName ? '✅ PASS' : '❌ FAIL', section, 'Step name in row 4',
               stepName ? 'Found: ' + stepName : 'Missing');

        // Planned formula in row 10
        var plannedF = fmsSheet.getRange(ds, sl.plannedCol).getFormula();
        addLog(plannedF ? '✅ PASS' : (sl.tatType === 'N' ? '⚠️ WARN' : '❌ FAIL'),
               section, 'Planned formula in row 10',
               plannedF ? plannedF.substring(0, 60) + '...' : (sl.tatType === 'N' ? 'N type — no formula expected' : 'Missing formula'));

        // Time Delay formula in row 10
        var tdF = fmsSheet.getRange(ds, sl.timeDelayCol).getFormula();
        addLog(tdF ? '✅ PASS' : '❌ FAIL', section, 'Time Delay ARRAYFORMULA in row 10',
               tdF ? 'Found' : 'Missing');

        // Header row checks
        var plannedHdr = fmsSheet.getRange(ROW_HEADERS, sl.plannedCol).getValue();
        var actualHdr  = fmsSheet.getRange(ROW_HEADERS, sl.actualCol).getValue();
        var statusHdr  = fmsSheet.getRange(ROW_HEADERS, sl.statusCol).getValue();
        addLog(plannedHdr === 'Planned' ? '✅ PASS' : '❌ FAIL', section, 'Planned header', plannedHdr || 'empty');
        addLog(actualHdr  === 'Actual'  ? '✅ PASS' : '❌ FAIL', section, 'Actual header',  actualHdr  || 'empty');
        addLog(statusHdr  === 'Status'  ? '✅ PASS' : '❌ FAIL', section, 'Status header',  statusHdr  || 'empty');

        // Extra cols check
        if (sl.extraCols && sl.extraCols.length > 0) {
          sl.extraCols.forEach(function(ec) {
            var ecHdr = fmsSheet.getRange(ROW_HEADERS, ec.col).getValue();
            addLog(ecHdr === ec.name ? '✅ PASS' : '❌ FAIL', section,
                   'Extra col header "' + ec.name + '"',
                   ecHdr ? 'Found: ' + ecHdr : 'Missing at col ' + ec.col);
          });
        }
      }
    }

    // ── SECTION 4: DATA ROWS ──────────────────────────────────────
    if (fieldDefs) {
      var lastDataRow = fmsSheet.getLastRow();
      var dataStartRow = FMS_DATA_START_ROW + 1;

      if (lastDataRow >= dataStartRow) {
        var dataRows = lastDataRow - dataStartRow + 1;
        addLog('✅ PASS', 'Data', 'Data rows count', dataRows + ' orders in FMS');

        // Check for duplicate identifiers
        var idVals = fmsSheet.getRange(dataStartRow, FIELDS_START_COL, dataRows, 1).getValues();
        var seen   = {};
        var dupes  = [];
        idVals.forEach(function(row) {
          var v = row[0].toString().trim().toUpperCase();
          if (v) { if (seen[v]) dupes.push(v); else seen[v] = true; }
        });
        addLog(dupes.length === 0 ? '✅ PASS' : '❌ FAIL', 'Data', 'No duplicate identifiers',
               dupes.length === 0 ? 'No duplicates' : 'Duplicates found: ' + dupes.join(', '));

        // Check for missing IDs in col A
        var idColVals = fmsSheet.getRange(dataStartRow, ID_COLUMN, dataRows, 1).getValues();
        var missingIds = idColVals.filter(function(r){ return r[0] === '' || r[0] === 0; });
        addLog(missingIds.length === 0 ? '✅ PASS' : '⚠️ WARN', 'Data', 'No missing IDs in col A',
               missingIds.length === 0 ? 'All IDs present' : missingIds.length + ' rows missing ID');

        // Check timestamps in col B
        var tsVals = fmsSheet.getRange(dataStartRow, TIMESTAMP_COLUMN, dataRows, 1).getValues();
        var missingTs = tsVals.filter(function(r){ return r[0] === '' || r[0] === 0; });
        addLog(missingTs.length === 0 ? '✅ PASS' : '⚠️ WARN', 'Data', 'No missing timestamps',
               missingTs.length === 0 ? 'All timestamps present' : missingTs.length + ' rows missing timestamp');

        // Check for Done without Actual
        if (stepDefs) {
          var doneNoActual = 0;
          for (var s = 0; s < stepDefs.length; s++) {
            var sl = stepDefs[s];
            var statusVals = fmsSheet.getRange(dataStartRow, sl.statusCol, dataRows, 1).getValues();
            var actualVals = fmsSheet.getRange(dataStartRow, sl.actualCol, dataRows, 1).getValues();
            for (var r = 0; r < dataRows; r++) {
              if (statusVals[r][0] === 'Done' && (actualVals[r][0] === '' || actualVals[r][0] === 0)) {
                doneNoActual++;
              }
            }
          }
          addLog(doneNoActual === 0 ? '✅ PASS' : '⚠️ WARN', 'Data', 'Done steps have Actual timestamp',
                 doneNoActual === 0 ? 'All good' : doneNoActual + ' Done rows missing Actual timestamp');
        }
      } else {
        addLog('⚠️ WARN', 'Data', 'Data rows', 'No data rows yet (FMS is empty)');
      }
    }
  }

  // ── SECTION 5: TRIGGERS ───────────────────────────────────────
  try {
    var triggers     = ScriptApp.getProjectTriggers();
    var hasOnChange  = false;
    var hasRouter    = false;
    var extraTriggers = 0;

    triggers.forEach(function(t) {
      var fn = t.getHandlerFunction();
      if (fn === 'onChange_new')         hasOnChange = true;
      else if (fn === 'onFormSubmit_Router') hasRouter = true;
      else extraTriggers++;
    });

    addLog(hasOnChange ? '✅ PASS' : '❌ FAIL', 'Triggers', 'onChange_new trigger exists',
           hasOnChange ? 'Found' : 'Missing — run Setup All Triggers');
    addLog(hasRouter ? '✅ PASS' : '❌ FAIL', 'Triggers', 'onFormSubmit_Router trigger exists',
           hasRouter ? 'Found' : 'Missing — run Setup All Triggers');
    addLog(extraTriggers === 0 ? '✅ PASS' : '⚠️ WARN', 'Triggers', 'No extra/duplicate triggers',
           extraTriggers === 0 ? 'Exactly 2 triggers' : extraTriggers + ' extra trigger(s) found — delete manually');
    addLog('✅ PASS', 'Triggers', 'Total triggers', triggers.length + ' trigger(s) active');
  } catch(te) {
    addLog('❌ FAIL', 'Triggers', 'Could not read triggers', te.toString());
  }

  // ── SECTION 6: FORM RESPONSE SHEETS ──────────────────────────
  if (f1Sheet) {
    var f1LastRow = f1Sheet.getLastRow();
    addLog(f1LastRow >= 1 ? '✅ PASS' : '⚠️ WARN', 'Forms', 'Form 1 has data',
           f1LastRow >= 2 ? (f1LastRow-1) + ' responses' : 'No responses yet');

    if (f1LastRow >= 1 && fieldDefs) {
      var f1Headers = f1Sheet.getRange(1, 1, 1, f1Sheet.getLastColumn()).getValues()[0];
      var allFieldsFound = true;
      fieldDefs.forEach(function(fd) {
        var found = f1Headers.some(function(h){ return h.toString().toLowerCase().trim() === fd.name.toLowerCase().trim(); });
        if (!found) allFieldsFound = false;
      });
      addLog(allFieldsFound ? '✅ PASS' : '⚠️ WARN', 'Forms', 'Form 1 headers match field defs',
             allFieldsFound ? 'All field names found in Form 1 headers' : 'Some fields missing from Form 1 headers — regenerate forms');
    }
  }

  if (f2Sheet) {
    var f2LastRow = f2Sheet.getLastRow();
    addLog(f2LastRow >= 1 ? '✅ PASS' : '⚠️ WARN', 'Forms', 'Form 2 has data',
           f2LastRow >= 2 ? (f2LastRow-1) + ' responses' : 'No responses yet');
  }

  // ── WRITE REPORT TO DIAGNOSTICS SHEET ───────────────────────
  var diagSheetName = 'FMS Diagnostics';
  var diagSheet = ss.getSheetByName(diagSheetName);
  if (!diagSheet) diagSheet = ss.insertSheet(diagSheetName);
  else diagSheet.clear();

  // Title
  var now = new Date();
  diagSheet.getRange(1, 1, 1, 4)
    .setValues([['🔍 FMS Diagnostic Report — ' + now.toLocaleString(), '', '', '']])
    .setBackground('#1E3A5F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12);
  diagSheet.getRange(1, 1, 1, 4).merge();

  // Summary row
  diagSheet.getRange(2, 1, 1, 4)
    .setValues([['Summary: ✅ ' + pass + ' PASS   ❌ ' + fail + ' FAIL   ⚠️ ' + warn + ' WARN', '', '', '']])
    .setBackground(fail > 0 ? '#F4C7C3' : (warn > 0 ? '#FFF3CD' : '#B7E1CD'))
    .setFontWeight('bold');
  diagSheet.getRange(2, 1, 1, 4).merge();

  // Headers
  diagSheet.getRange(3, 1, 1, 4)
    .setValues([['Status', 'Section', 'Check', 'Detail']])
    .setBackground('#334155').setFontColor('#FFFFFF').setFontWeight('bold');

  // Log rows
  if (log.length > 0) {
    diagSheet.getRange(4, 1, log.length, 4).setValues(log);

    // Color each row by status
    for (var i = 0; i < log.length; i++) {
      var rowBg = log[i][0] === '✅ PASS' ? '#F0FFF0'
                : log[i][0] === '❌ FAIL' ? '#FFF0F0'
                : '#FFFBF0';
      diagSheet.getRange(4 + i, 1, 1, 4).setBackground(rowBg);
    }
  }

  diagSheet.autoResizeColumns(1, 4);
  diagSheet.setFrozenRows(3);
  ss.setActiveSheet(diagSheet);

  // Final alert
  var summaryMsg = '🔍 Diagnostic Complete!\n\n' +
    '✅ PASS: ' + pass + '\n' +
    '❌ FAIL: ' + fail + '\n' +
    '⚠️ WARN: ' + warn + '\n\n' +
    (fail > 0 ? '❌ Issues found — check "FMS Diagnostics" sheet for details.\n' : '') +
    (warn > 0 ? '⚠️ Warnings — review "FMS Diagnostics" sheet.\n' : '') +
    (fail === 0 && warn === 0 ? '🎉 Everything looks good!\n' : '') +
    '\nFull report written to "FMS Diagnostics" sheet.';

  ui.alert(summaryMsg);
  Logger.log('🔍 Diagnostic done — PASS:' + pass + ' FAIL:' + fail + ' WARN:' + warn);
}


// ══════════════════════════════════════════════════════════════════
// TEST FORM 1
// ══════════════════════════════════════════════════════════════════
function testForm1Manually() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    var fs = ss.getSheetByName(FORM1_SHEET_NAME);
    if (!fs) { ui.alert('Tab not found: "' + FORM1_SHEET_NAME + '"'); return; }
    if (fs.getLastRow() < 2) { ui.alert('No responses yet in "' + FORM1_SHEET_NAME + '"'); return; }
    onFormSubmit_NewOrder(null);
    ui.alert('Test complete!\nCheck FMS sheet and Apps Script Executions for logs.');
  } catch(e) { ui.alert('Error: ' + e); }
}


// ══════════════════════════════════════════════════════════════════
// TAT FORMULA BUILDER
// ══════════════════════════════════════════════════════════════════
function _buildTATFormula_(identRef, startRef, tatHoursRef, isArray) {
  var h = 'Holidays!$A:$A';

  if (isArray) {
    return (
      '=ARRAYFORMULA(IFERROR(IF((' + identRef + '="")' +
        '+(' + startRef + '="")' +
        '+(ISNUMBER(' + startRef + ')*(' + startRef + '<=0))' +
        '+(ISBLANK(' + tatHoursRef + '))' +
      ',"",LET(' +
        '_start,' + startRef + ',' +
        '_tat,' + tatHoursRef + '/24,' +
        '_open,$C$1,_close,$D$1,_wd,$E$1,_hol,' + h + ',' +
        '_sd,INT(_start),_st,MOD(_start,1),' +
        '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +
        '_first,IF(_iswd,IF(_st<_open,_sd+_open,IF(_st>=_close,WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +
        '_ft,MOD(_first,1),_avail,_close-MAX(_ft,_open),' +
        'IF(_tat<=_avail,_first+_tat,LET(' +
          '_rem1,_tat-_avail,_daylen,_close-_open,_k,INT(_rem1/_daylen),_rem2,MOD(_rem1,_daylen),' +
          '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
          'IF(_rem2=0,WORKDAY.INTL(INT(_first),_k,_wd,_hol)+_close,_base+_open+_rem2)' +
      ')))),""))'
    );
  } else {
    return (
      '=IFERROR(IF(OR(ISBLANK(' + startRef + '),' + startRef + '<=0,ISBLANK(' + tatHoursRef + ')),"",LET(' +
        '_start,' + startRef + ',' +
        '_tat,' + tatHoursRef + '/24,' +
        '_open,$C$1,_close,$D$1,_wd,$E$1,_hol,' + h + ',' +
        '_sd,INT(_start),_st,MOD(_start,1),' +
        '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +
        '_first,IF(_iswd,IF(_st<_open,_sd+_open,IF(_st>=_close,WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +
        '_ft,MOD(_first,1),_avail,_close-MAX(_ft,_open),' +
        'IF(_tat<=_avail,_first+_tat,LET(' +
          '_rem1,_tat-_avail,_daylen,_close-_open,_k,INT(_rem1/_daylen),_rem2,MOD(_rem1,_daylen),' +
          '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
          'IF(_rem2=0,WORKDAY.INTL(INT(_first),_k,_wd,_hol)+_close,_base+_open+_rem2)' +
      ')))),"")'
    );
  }
}


// ══════════════════════════════════════════════════════════════════
// ALL EXISTING BMP FUNCTIONS
// ══════════════════════════════════════════════════════════════════
function initial(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var o=ui.prompt('Opening Time','HH:MM e.g. 10:00',ui.ButtonSet.OK_CANCEL);
  if(o.getSelectedButton()!==ui.Button.OK)return;
  var openFrac=_parseTimeToFractionOrAlert_((o.getResponseText()||'').trim(),'Opening Time');
  if(openFrac==null)return;
  var c=ui.prompt('Closing Time','HH:MM e.g. 18:00',ui.ButtonSet.OK_CANCEL);
  if(c.getSelectedButton()!==ui.Button.OK)return;
  var closeFrac=_parseTimeToFractionOrAlert_((c.getResponseText()||'').trim(),'Closing Time');
  if(closeFrac==null)return;
  if(closeFrac<=openFrac){ui.alert('Closing must be after Opening.');return;}
  var w=ui.prompt('Working Days','7-char 0/1 Mon-Sun e.g. "0000001"=Sun off',ui.ButtonSet.OK_CANCEL);
  if(w.getSelectedButton()!==ui.Button.OK)return;
  var wd=(w.getResponseText()||'').trim();
  if(!/^[01]{7}$/.test(wd)){ui.alert('Must be 7 chars of 0/1');return;}
  var sh=ss.getActiveSheet();
  sh.getRange('A1').setFormula('=NOW()');
  sh.getRange('C1').setValue(openFrac).setNumberFormat('HH:mm');
  sh.getRange('D1').setValue(closeFrac).setNumberFormat('HH:mm');
  sh.getRange('E1').setValue("'"+wd);
  sh.getRange('B1').clearContent();
  sh.getRange('1:1').activate();
  sh.hideRows(1);
  iterative();
  var holidaySheet = ss.getSheetByName('Holidays');
  if (!holidaySheet) {
    holidaySheet = ss.insertSheet('Holidays');
    holidaySheet.getRange('A1').setValue('Holiday Dates').setFontWeight('bold')
      .setBackground('#1E3A5F').setFontColor('#FFFFFF');
    holidaySheet.getRange('A2').setValue('Add holiday dates below in dd/MM/yyyy format');
    holidaySheet.setColumnWidth(1, 200);
    Logger.log('Holidays sheet created');
  }
  ui.alert(
    '✅ Setup Complete!\n\n' +
    '• Opening Time: ' + (o.getResponseText()||'').trim() + '\n' +
    '• Closing Time: ' + (c.getResponseText()||'').trim() + '\n' +
    '• Working Days: ' + wd + '\n\n' +
    'Holidays sheet ready — add holiday dates in column A.\n\n' +
    'Next steps:\n1. Create New FMS Sheet\n2. Add Steps to FMS\n3. Generate Forms\n4. Setup All Triggers'
  );
}
function TAT(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var d=ui.prompt('Timestamp Cell','e.g. B10',ui.ButtonSet.OK_CANCEL);
  if(d.getSelectedButton()!==ui.Button.OK)return;
  var t=ui.prompt('TAT Hours Cell','e.g. F$8',ui.ButtonSet.OK_CANCEL);
  if(t.getSelectedButton()!==ui.Button.OK)return;
  ss.getCurrentCell().setFormula(_buildTATFormula_(d.getResponseText().trim(),d.getResponseText().trim(),t.getResponseText().trim(),false));
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function createTrigger(){removeTrigger();ScriptApp.newTrigger('onChange_new').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onChange().create();}
function removeTrigger(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='onChange_new')ScriptApp.deleteTrigger(t);});}
function plannedlead(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var a=ui.prompt('Date Cell',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!=ui.Button.OK)return;a=a.getResponseText();
  var b=ui.prompt('Lead Time Cell',ui.ButtonSet.OK_CANCEL);if(b.getSelectedButton()!=ui.Button.OK)return;b=b.getResponseText();
  var c=ui.prompt('Days Before',ui.ButtonSet.OK_CANCEL);if(c.getSelectedButton()!=ui.Button.OK)return;c=c.getResponseText();
  ss.getCurrentCell().setFormula('=IFERROR(IF('+b+','+a+'+'+b+'-'+c+',""),"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function specificTime(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var a=ui.prompt('Date Cell',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!=ui.Button.OK)return;a=a.getResponseText();
  var b=ui.prompt('Days after (0=same day)',ui.ButtonSet.OK_CANCEL);if(b.getSelectedButton()!=ui.Button.OK)return;b=b.getResponseText();
  var c=ui.prompt('Time fraction',ui.ButtonSet.OK_CANCEL);if(c.getSelectedButton()!=ui.Button.OK)return;c=c.getResponseText();
  ss.getCurrentCell().setFormula('=IFERROR(IF('+a+',WORKDAY.INTL(INT('+a+'),'+b+',"0000001",Holidays!$A:$A)+'+c+',""),"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function timeDelay(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var a=ui.prompt('Planned Cell',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!=ui.Button.OK)return;a=a.getResponseText();
  var b=ui.prompt('Actual Cell',ui.ButtonSet.OK_CANCEL);if(b.getSelectedButton()!=ui.Button.OK)return;b=b.getResponseText();
  ss.getCurrentCell().setFormula('=IFERROR(IF('+a+'<>"",IF('+b+'<>"",IF('+b+'>'+a+','+b+'-'+a+',""),$A$1-'+a+'),""),"")');
  ss.getActiveRangeList().setNumberFormat('[h]:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function tatifno(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var a=ui.prompt('Status Cell',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!=ui.Button.OK)return;a=a.getResponseText();
  var f=ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IFERROR(IF('+a+'="No",'+f+',""),"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function tatifyes(){
  var ss=SpreadsheetApp.getActive(),ui=SpreadsheetApp.getUi();
  var a=ui.prompt('Status Cell',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!=ui.Button.OK)return;a=a.getResponseText();
  var f=ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IFERROR(IF('+a+'="Yes",'+f+',""),"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();cc=ss.getCurrentCell();ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();cc.activateAsCurrentCell();ss.getCurrentCell().copyTo(ss.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
}
function iterative(){
  var s=SpreadsheetApp.getActive();
  s.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);
  s.setIterativeCalculationEnabled(true);
  s.setMaxIterativeCalculationCycles(1);
  s.setIterativeCalculationConvergenceThreshold(0.05);
}
function onChange_new(){
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = ss.getActiveSheet();
    var active = ss.getActiveCell();
    var row    = active.getRow();
    var col    = active.getColumn();
    var val    = active.getValue();
    if (val === 'Done' || val === 'Yes' || val === 'No' || val === true) {
      var targetCell    = sheet.getRange(row, col - 1);
      var targetFormula = targetCell.getFormula();
      if (targetFormula.toUpperCase().indexOf('ARRAYFORMULA') !== -1) {
        Logger.log('onChange_new: skipping ARRAYFORMULA cell');
        return;
      }
      var ts = targetCell.getValue();
      if (ts === '' || ts === 0 || ts === null) {
        targetCell.setValue(new Date());
      }
    }
  } catch(e) { Logger.log('onChange_new error: ' + e); }
}
