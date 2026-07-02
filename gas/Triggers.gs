function installAstraTriggers_() {
  removeAstraTriggers_();
  const spreadsheet = getResponseSpreadsheet_();
  ScriptApp.newTrigger('onFormSubmitInstalled_')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
  ScriptApp.newTrigger('processPendingSubmissions')
    .timeBased()
    .everyMinutes(ASTRA_CONFIG.WORKER_INTERVAL_MINUTES)
    .create();
}

function removeAstraTriggers_() {
  const handlers = ['onFormSubmitInstalled_', 'processPendingSubmissions'];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function onFormSubmitInstalled_(event) {
  if (!event || !event.range) {
    throw new Error('フォーム送信イベントから回答行を取得できませんでした。');
  }
  const sheet = event.range.getSheet();
  const rowNumber = event.range.getRow();
  if (sheet.getName() !== getResponseSheetName_()) {
    appendLog_('WARN', '', rowNumber, 'フォーム送信',
      '対象外シート「' + sheet.getName() + '」への送信のため処理しません。' +
      '本命フォームのシートなら Script Properties の RESPONSE_SHEET_NAME を更新してください。');
    return;
  }
  markRowPending_(sheet, rowNumber);

  try {
    processSubmissionWithLock_(rowNumber, false);
  } catch (error) {
    appendLog_('ERROR', '', rowNumber, 'フォーム送信直後処理', error.stack || error.message);
    // 未処理またはエラー行は5分ごとのワーカーが最大3回まで再試行する。
  }
}
