const PAINT_SHEET_NAME = "RateShiftData";
const STAR_SHEET_NAME = "StarData";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: "Rate-shift upload endpoint is ready." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const rows = Array.isArray(payload.rows)
      ? payload.rows.map((row) => ({
          ...row,
          condition: formatTaskCondition(row.condition),
          ...(row.trial_type ? { trial_type: formatTaskCondition(row.trial_type) } : {}),
          entry_condition: formatEntryCondition(
            row.entry_condition || payload.entry_condition || "standalone"
          )
        }))
      : [];

    if (!rows.length) {
      return jsonResponse({ ok: false, error: "No rows received." });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = payload.task === "star_task"
      ? STAR_SHEET_NAME
      : PAINT_SHEET_NAME;
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    const headers = ensureHeaders(sheet, rows, payload.task);
    const values = rows.map((row) => headers.map((header) => valueForCell(row[header])));

    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    return jsonResponse({ ok: true, appended: values.length });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function ensureHeaders(sheet, rows, task) {
  const commonHeaders = [
    "entry_condition",
    "participant_id",
    "task",
    "task_version",
    "session_id",
    "session_started_at",
    "session_finished_at",
    "trial_index",
    "condition",
    "event_type",
    "response_recorded_at",
    "event_time_ms",
    "delay_duration_ms",
    "delay_onset_ms",
    "key_event_index",
    "time_since_delay_onset_ms",
    "key_hold_duration_ms",
    "is_initial_press",
    "is_extra_press",
    "is_auto_repeat",
    "second_bin"
  ];

  const paintHeaders = commonHeaders.concat([
    "phase",
    "block",
    "segment_index",
    "challenge_slot",
    "response_index",
    "condition_key",
    "transition_from",
    "challenge_pattern",
    "rate_cells_per_sec",
    "effective_rate_cells_per_sec",
    "ideal_interval_ms",
    "elapsed_ms",
    "segment_elapsed_ms",
    "segment_start_ms",
    "segment_end_ms",
    "segment_response_count",
    "is_transition_window",
    "first_response_latency_ms",
    "within_segment_response_interval_ms",
    "response_interval_ms",
    "interval_error_ms",
    "acc",
    "target_hit_acc",
    "interval_accuracy",
    "level_status",
    "timing_status",
    "level_proximity",
    "level_proximity_percent",
    "level_error_cells",
    "timing_correct",
    "status",
    "filled_cells_before_response",
    "filled_cells_after_response",
    "cell_change_direction",
    "cell_change_cells",
    "input_source",
    "input_key",
    "input_code",
    "input_location",
    "input_repeat",
    "input_recorded_at",
    "speed_mode",
    "accuracy_window_ms",
    "user_agent"
  ]);

  const starHeaders = commonHeaders.concat([
    "trial",
    "trial_type",
    "attempt",
    "delay_ms",
    "delay_duration_ms",
    "delay_onset_ms",
    "analysis_included",
    "target_side",
    "star_count",
    "distractor_shape",
    "chosen_side",
    "correct",
    "won_amount",
    "response_rt_ms",
    "elapsed_ms",
    "key_event_index",
    "event_time_ms",
    "time_since_delay_onset_ms",
    "key_hold_duration_ms",
    "is_initial_press",
    "is_extra_press",
    "is_auto_repeat",
    "second_bin",
    "input_key",
    "input_code",
    "input_location",
    "input_repeat",
    "input_sequence",
    "pointer_event_index",
    "pointer_type",
    "pointer_id",
    "pointer_x",
    "pointer_y",
    "page_x",
    "page_y",
    "screen_x",
    "screen_y",
    "movement_x",
    "movement_y",
    "button",
    "buttons",
    "target_id",
    "target_class",
    "viewport_width",
    "viewport_height",
    "delay_elapsed_ms",
    "trial_elapsed_ms",
    "input_recorded_at",
    "user_agent"
  ]);

  const preferred = task === "star_task" ? starHeaders : paintHeaders;

  const discovered = [];
  rows.forEach((row) => {
    Object.keys(row).forEach((key) => {
      if (!preferred.includes(key) && !discovered.includes(key)) discovered.push(key);
    });
  });

  if (sheet.getLastRow() === 0) {
    const headers = preferred.concat(discovered);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return headers;
  }

  const width = Math.max(sheet.getLastColumn(), 1);
  let existing = sheet.getRange(1, 1, 1, width).getValues()[0].filter(String);
  const entryIndex = existing.indexOf("entry_condition");
  if (entryIndex === -1) {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("entry_condition");
    existing = ["entry_condition"].concat(existing);
  } else if (entryIndex > 0) {
    sheet.moveColumns(sheet.getRange(1, entryIndex + 1, sheet.getMaxRows(), 1), 1);
    existing = ["entry_condition"].concat(existing.filter((key) => key !== "entry_condition"));
  }
  const additions = preferred.concat(discovered).filter((key) => !existing.includes(key));

  if (additions.length) {
    sheet.getRange(1, existing.length + 1, 1, additions.length).setValues([additions]);
  }

  const allHeaders = existing.concat(additions);
  const orderedHeaders = preferred.concat(
    allHeaders.filter((header) => !preferred.includes(header))
  );
  reorderSheetColumns(sheet, allHeaders, orderedHeaders);
  return orderedHeaders;
}

function reorderSheetColumns(sheet, currentHeaders, orderedHeaders) {
  const changed = currentHeaders.some(
    (header, index) => header !== orderedHeaders[index]
  );
  if (!changed) return;

  const dataRowCount = Math.max(sheet.getLastRow() - 1, 0);
  const currentColumnCount = currentHeaders.length;
  const data = dataRowCount
    ? sheet.getRange(2, 1, dataRowCount, currentColumnCount).getValues()
    : [];
  const indexByHeader = Object.fromEntries(
    currentHeaders.map((header, index) => [header, index])
  );
  const reorderedData = data.map((row) => orderedHeaders.map((header) => {
    const index = indexByHeader[header];
    return index === undefined ? "" : row[index];
  }));

  sheet.getRange(1, 1, 1, orderedHeaders.length).setValues([orderedHeaders]);
  if (reorderedData.length) {
    sheet
      .getRange(2, 1, reorderedData.length, orderedHeaders.length)
      .setValues(reorderedData);
  }
}

function valueForCell(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object") return JSON.stringify(value);
  return value;
}

function formatEntryCondition(value) {
  const labels = {
    combined_paint_first: "통합 과제: 물감 먼저",
    combined_coin_first: "통합 과제: 별 먼저",
    paint_only: "물감 과제 단독",
    coin_only: "별 과제 단독",
    standalone: "단독 실행"
  };

  return labels[value] || value || "단독 실행";
}

function jsonResponse(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatTaskCondition(value) {
  const labels = {
    "practice": "연습 시행",
    "practice-no-delay": "연습 시행"
  };

  return labels[value] || value;
}
