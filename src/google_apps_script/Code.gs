var TASK_TRACKER_CONFIG_ = {
  // For a bound script this remains empty. A standalone deployment can set
  // SPREADSHEET_ID in Script Properties instead.
  spreadsheetId: '',
  incomingSheet: 'Входящие',
  tasksSheet: 'Задачи',
  subtasksSheet: 'Подзадачи',
  routinesSheet: 'Рутины',
  shoppingSheet: 'Список покупок',
  referencesSheet: 'Справочники',
  settingsSheet: 'Настройки',
  headerRow: 11,
  dataStartRow: 12,
  incomingReminderDays: 1,
  morningHour: 8,
  eveningPlanHour: 20,
  weeklyReflectionHour: 12,
  draftWaitMinutes: 2
};

var WORK_STATUSES_ = ['new', 'planned', 'in_progress', 'blocked', 'done', 'cancelled', 'skipped'];
var TASK_PRIORITIES_ = ['1', '2', '3'];
var INCOMING_STATUSES_ = WORK_STATUSES_;
var INCOMING_TYPES_ = ['задача', 'рутина', 'покупка'];
var ROUTINE_STATUSES_ = ['active', 'paused', 'archived'];
var ROUTINE_REPEAT_RULES_ = ['ежедневно', 'еженедельно', 'каждые N дней', 'каждые N недель', 'ежемесячно', 'ежегодно', 'вручную'];
var CLOSED_STATUSES_ = ['done', 'cancelled', 'skipped'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Таск-трекер')
    .addItem('Создать готовые задачи сейчас', 'runTaskTrackerAutomationNow')
    .addItem('Собрать сводку за сегодня', 'createDailySummaryNow')
    .addItem('Собрать итоги недели', 'createWeeklySummaryNow')
    .addItem('Позвать на недельную рефлексию', 'sendWeeklyReflectionNow')
    .addItem('Как это работает', 'showTaskTrackerMenuHelp')
    .addSeparator()
    .addItem('Восстановить автоматизацию', 'installTaskTrackerAutomation')
    .addToUi();
}

function showTaskTrackerMenuHelp() {
  SpreadsheetApp.getUi().alert(
    'Таск-трекер',
    '«Создать готовые задачи сейчас» сразу разбирает заполненные Входящие, не дожидаясь двух минут. Автоматическая обработка по таймеру по-прежнему ждет две минуты после последней правки, чтобы не создавать недописанные черновики.\n\n' +
    '«Восстановить автоматизацию» используйте только если что-то перестало обновляться само. Команда пересоздаст технические триггеры и сразу выполнит обновление. Данные в таблице не удаляются.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function installTaskTrackerAutomation() {
  ensureTaskTrackerSchema_();
  deleteAutomationTriggers_();

  var spreadsheet = getTrackerSpreadsheet_();
  ScriptApp.newTrigger('onIncomingTaskDraftEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
  ScriptApp.newTrigger('runTaskTrackerAutomation')
    .timeBased()
    // Проверяем раз в минуту, но создаем задачу только спустя две минуты без правок.
    .everyMinutes(1)
    .create();
  runTaskTrackerAutomation();
}

function onIncomingTaskDraftEdit(event) {
  if (!event || !event.range) {
    return;
  }
  var sheet = event.range.getSheet();
  if (event.range.getRow() < TASK_TRACKER_CONFIG_.dataStartRow) {
    return;
  }

  if (sheet.getName() === TASK_TRACKER_CONFIG_.incomingSheet) {
    var incomingHeaders = getHeaders_(sheet);
    var draftColumn = columnNumber_(incomingHeaders, 'Формулировка задачи');
    var typeColumn = columnNumber_(incomingHeaders, 'Предложенный тип');
    if (event.range.getColumn() === draftColumn || event.range.getColumn() === typeColumn) {
      markIncomingDraftChanged_(event.range.getRow());
      scheduleDraftPromotion_();
    }
    return;
  }

  if (sheet.getName() === TASK_TRACKER_CONFIG_.subtasksSheet) {
    var subtaskHeaders = getHeaders_(sheet);
    if (rangeIncludesColumn_(event.range, columnNumber_(subtaskHeaders, 'Задача'))) {
      linkSubtaskToSelectedTask_(event.range.getRow());
      refreshSubtaskChecks();
    }
    if (rangeIncludesColumn_(event.range, columnNumber_(subtaskHeaders, 'Дата'))) {
      syncSubtaskPlanningStatuses_(event.range.getRow(), event.range.getNumRows());
    }
    if (rangeIncludesColumn_(event.range, columnNumber_(subtaskHeaders, 'Статус'))) {
      updateClosedDates_(sheet, subtaskHeaders, event.range, 'Статус', 'Дата закрытия');
      touchSubtaskUpdated_(event.range);
      completeRoutineFromSubtask_(event.range.getRow(), true);
      syncTaskStatusesFromSubtasks_();
    }
    return;
  }

  if (sheet.getName() === TASK_TRACKER_CONFIG_.shoppingSheet) {
    var shoppingHeaders = getHeaders_(sheet);
    if (rangeIncludesColumn_(event.range, columnNumber_(shoppingHeaders, 'Статус'))) {
      updateClosedDates_(sheet, shoppingHeaders, event.range, 'Статус', 'Дата закрытия');
      touchRowsUpdated_(sheet, shoppingHeaders, event.range);
    }
    return;
  }

  if (sheet.getName() === TASK_TRACKER_CONFIG_.tasksSheet) {
    var taskHeaders = getHeaders_(sheet);
    if (rangeIncludesColumn_(event.range, columnNumber_(taskHeaders, 'Заголовок')) ||
        rangeIncludesColumn_(event.range, columnNumber_(taskHeaders, 'Проект / область'))) {
      // Название и проект задачи - источники истины для связанных подзадач и плана дня.
      syncSubtaskTaskTitles_();
    }
    if (rangeIncludesColumn_(event.range, columnNumber_(taskHeaders, 'Статус'))) {
      // Статус задачи не редактируется вручную: он всегда следует за подзадачами.
      syncTaskStatusesFromSubtasks_();
    }
    return;
  }

  if (sheet.getName() === TASK_TRACKER_CONFIG_.routinesSheet) {
    syncRoutineOccurrences_(true);
  }
}

function runTaskTrackerAutomation() {
  ensureTaskTrackerSchema_();
  reconcileIncomingTaskLinks_();
  promoteReadyIncoming_(false);
  ensureStarterSubtasks_();
  syncSubtaskTaskTitles_();
  syncSubtaskPlanningStatuses_();
  refreshSubtaskChecks();
  syncTaskStatusesFromSubtasks_();
  syncRoutineOccurrences_(false);
  createDailySummaryIfDue_();
  createWeeklySummaryIfDue_();
  sendWeeklyReflectionIfDue_();
  sendDailyPlanIfDue_();
  sendEveningPlanIfDue_();
  sendDailyIncomingReminderIfDue_();
  sendShoppingListIfDue_();
}

// Ручной запуск осознанно обходит задержку черновика: пользователь уже закончил ввод.
function runTaskTrackerAutomationNow() {
  ensureTaskTrackerSchema_();
  reconcileIncomingTaskLinks_();
  promoteReadyIncoming_(true);
  ensureStarterSubtasks_();
  syncSubtaskTaskTitles_();
  syncSubtaskPlanningStatuses_();
  refreshSubtaskChecks();
  syncTaskStatusesFromSubtasks_();
  syncRoutineOccurrences_(false);
}

// "Задача создана" допустима только пока все сохраненные ID существуют в листе задач.
function reconcileIncomingTaskLinks_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var incomingSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var routinesSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.routinesSheet);
  var incomingHeaders = getHeaders_(incomingSheet);
  var taskHeaders = getHeaders_(tasksSheet);
  if (!hasDataRows_(incomingSheet)) {
    return;
  }

  var createdIds = {};
  if (hasDataRows_(tasksSheet)) {
    tasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, columnNumber_(taskHeaders, 'ID'), dataRowCount_(tasksSheet), 1).getValues().forEach(function(row) {
      if (row[0]) {
        createdIds[String(row[0]).trim()] = true;
      }
    });
  }
  if (hasDataRows_(routinesSheet)) {
    var routineHeaders = getHeaders_(routinesSheet);
    routinesSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, columnNumber_(routineHeaders, 'ID'), dataRowCount_(routinesSheet), 1).getValues().forEach(function(row) {
      if (row[0]) {
        createdIds[String(row[0]).trim()] = true;
      }
    });
  }

  incomingSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(incomingSheet), incomingHeaders.length).getValues().forEach(function(row, index) {
    var item = rowToObject_(incomingHeaders, row);
    if (String(item['Статус разбора'] || '').toLowerCase() !== 'done') {
      return;
    }
    var linkedIds = String(item['Созданные задачи ID'] || '').split(',').map(function(value) {
      return value.trim();
    }).filter(function(value) {
      return Boolean(value);
    });
    var allMissing = linkedIds.length === 0 || linkedIds.every(function(id) {
      return !createdIds[id];
    });
    if (allMissing) {
      setObjectFields_(incomingSheet, incomingHeaders, index + TASK_TRACKER_CONFIG_.dataStartRow, {
        'Статус разбора': 'new',
        'Созданные задачи ID': ''
      });
    }
  });
}

function promoteReadyIncoming_(forceDraftPromotion) {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var headers = getHeaders_(sheet);
  if (!hasDataRows_(sheet)) {
    return 0;
  }
  var rows = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues();
  var promoted = 0;
  rows.forEach(function(row, index) {
    var item = rowToObject_(headers, row);
    var rowNumber = index + TASK_TRACKER_CONFIG_.dataStartRow;
    // Покупка создается без черновика задачи и остается, даже если позже появится задача.
    if (isShoppingIncoming_(item['Предложенный тип']) &&
        !String(item['Созданные покупки ID'] || '').trim()) {
      promoted += promoteIncomingToShopping_(rowNumber) ? 1 : 0;
      item = rowToObject_(headers, sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0]);
    }
    if (String(item['Созданные покупки ID'] || '').trim() &&
        String(item['Статус разбора'] || '').toLowerCase() === 'new') {
      setObjectFields_(sheet, headers, rowNumber, { 'Статус разбора': 'done' });
      item = rowToObject_(headers, sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0]);
    }
    if ((!isTerminalIncomingStatus_(item['Статус разбора']) ||
         String(item['Созданные покупки ID'] || '').trim()) &&
        String(item['Формулировка задачи'] || '').trim() &&
        !String(item['Созданные задачи ID'] || '').trim() &&
        (forceDraftPromotion || isDraftReadyForPromotion_(item))) {
      promoted += promoteIncomingRow_(rowNumber) ? 1 : 0;
    }
  });
  return promoted;
}

function promoteIncomingRow_(rowNumber) {
  var spreadsheet = getTrackerSpreadsheet_();
  var incomingSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var incomingHeaders = getHeaders_(incomingSheet);
  var incoming = rowToObject_(
    incomingHeaders,
    incomingSheet.getRange(rowNumber, 1, 1, incomingHeaders.length).getValues()[0]
  );
  var draft = String(incoming['Формулировка задачи'] || '').trim();
  if (!draft || isTerminalIncomingStatus_(incoming['Статус разбора']) ||
      String(incoming['Созданные задачи ID'] || '').trim()) {
    return false;
  }

  var titles = splitTaskDrafts_(draft);
  if (!titles.length) {
    return false;
  }

  if (isRoutineIncoming_(incoming['Предложенный тип'])) {
    return promoteIncomingToRoutines_(incomingSheet, incomingHeaders, incoming, rowNumber, titles);
  }

  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var taskHeaders = getHeaders_(tasksSheet);
  var now = nowIso_();
  var taskRows = [];
  var taskIds = [];

  var existingTasks = getExistingTasks_(tasksSheet, taskHeaders);

  titles.forEach(function(title, index) {
    var duplicate = findDuplicateTask_(existingTasks, title);
    if (duplicate) {
      taskIds.push(duplicate.ID);
      return;
    }
    var taskId = 'task_' + String(incoming.ID) + '_' + (index + 1);
    taskIds.push(taskId);
    taskRows.push(objectToRow_(taskHeaders, {
      ID: taskId,
      Заголовок: title,
      'Проект / область': incoming['Подсказка проекта'] || '',
      Статус: 'new',
      // Приоритет всегда выбирается вручную: 1 - высокий, 3 - низкий.
      Приоритет: '',
      Дедлайн: extractDeadline_(title) || extractDeadline_(incoming['Исходный текст']) || '',
      Источник: incoming.Источник || 'Входящие',
      'Ссылка на входящее': incoming.ID,
      Создано: now,
      Обновлено: now,
      Заметки: incoming.Заметки || ''
    }));
    existingTasks.push({ ID: taskId, Заголовок: title });
  });

  if (taskRows.length) {
    tasksSheet.getRange(nextDataRow_(tasksSheet), 1, taskRows.length, taskHeaders.length).setValues(taskRows);
  }
  setObjectFields_(incomingSheet, incomingHeaders, rowNumber, {
    'Статус разбора': 'done',
    'Созданные задачи ID': taskIds.join(', ')
  });
  clearDraftState_(incoming.ID);
  return true;
}

// Рутины из входящих создаются на паузе: расписание нельзя угадать по одному сообщению.
function promoteIncomingToRoutines_(incomingSheet, incomingHeaders, incoming, rowNumber, titles) {
  var spreadsheet = getTrackerSpreadsheet_();
  var routinesSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.routinesSheet);
  var routineHeaders = getHeaders_(routinesSheet);
  var existingRoutines = getExistingRoutines_(routinesSheet, routineHeaders);
  var routineRows = [];
  var routineIds = [];

  titles.forEach(function(title, index) {
    var duplicate = findDuplicateRoutine_(existingRoutines, title);
    if (duplicate) {
      routineIds.push(duplicate.ID);
      return;
    }
    var routineId = 'routine_' + String(incoming.ID) + '_' + (index + 1);
    routineIds.push(routineId);
    routineRows.push(objectToRow_(routineHeaders, {
      ID: routineId,
      Название: title,
      Категория: incoming['Подсказка проекта'] || '',
      Статус: 'paused',
      Заметки: appendIncomingNote_(incoming.Заметки, incoming.ID)
    }));
    existingRoutines.push({ ID: routineId, Название: title });
  });

  if (routineRows.length) {
    routinesSheet.getRange(nextDataRow_(routinesSheet), 1, routineRows.length, routineHeaders.length).setValues(routineRows);
  }
  setObjectFields_(incomingSheet, incomingHeaders, rowNumber, {
    'Статус разбора': 'done',
    'Созданные задачи ID': routineIds.join(', ')
  });
  clearDraftState_(incoming.ID);
  return true;
}

// Покупка независима от задачи: та же строка Входящих может позднее создать задачу.
function promoteIncomingToShopping_(rowNumber) {
  var spreadsheet = getTrackerSpreadsheet_();
  var incomingSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var incomingHeaders = getHeaders_(incomingSheet);
  var incoming = rowToObject_(incomingHeaders, incomingSheet.getRange(rowNumber, 1, 1, incomingHeaders.length).getValues()[0]);
  if (!isShoppingIncoming_(incoming['Предложенный тип']) ||
      String(incoming['Созданные покупки ID'] || '').trim()) {
    return false;
  }

  var title = shoppingTitle_(incoming['Исходный текст']);
  if (!title) {
    return false;
  }
  var shoppingSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.shoppingSheet);
  var shoppingHeaders = getHeaders_(shoppingSheet);
  var existing = getExistingPurchases_(shoppingSheet, shoppingHeaders);
  var duplicate = findDuplicatePurchase_(existing, title);
  var purchaseId = duplicate ? duplicate.ID : 'purchase_' + String(incoming.ID);
  if (!duplicate) {
    shoppingSheet.getRange(nextDataRow_(shoppingSheet), 1, 1, shoppingHeaders.length).setValues([objectToRow_(shoppingHeaders, {
      ID: purchaseId,
      Покупка: title,
      Статус: 'new',
      Источник: incoming.Источник || 'Входящие',
      'Ссылка на входящее': incoming.ID,
      Создано: nowIso_(),
      Обновлено: nowIso_(),
      Заметки: incoming.Заметки || ''
    })]);
  }
  setObjectFields_(incomingSheet, incomingHeaders, rowNumber, {
    'Статус разбора': 'done',
    'Созданные покупки ID': purchaseId
  });
  return true;
}

function getExistingPurchases_(sheet, headers) {
  if (!hasDataRows_(sheet)) {
    return [];
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().map(function(row) {
    return rowToObject_(headers, row);
  }).filter(function(item) {
    return item.ID && item.Покупка && String(item.Статус || '').toLowerCase() === 'new';
  });
}

function findDuplicatePurchase_(purchases, title) {
  var normalizedTitle = normalizeTaskText_(title);
  return purchases.filter(function(item) {
    return normalizeTaskText_(item.Покупка) === normalizedTitle;
  })[0] || null;
}

function shoppingTitle_(rawText) {
  var text = String(rawText || '').trim();
  var withoutLeadingVerb = text.replace(/^\s*купить\b[\s:,-]*/i, '').trim();
  return withoutLeadingVerb || text;
}

function getExistingRoutines_(sheet, headers) {
  if (!hasDataRows_(sheet)) {
    return [];
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().map(function(row) {
    return rowToObject_(headers, row);
  }).filter(function(routine) {
    return routine.ID && routine.Название;
  });
}

function findDuplicateRoutine_(routines, title) {
  return routines.filter(function(routine) {
    return taskTextsOverlap_(title, routine.Название);
  })[0] || null;
}

function appendIncomingNote_(notes, incomingId) {
  var prefix = String(notes || '').trim();
  var source = 'Черновик создан из входящего: ' + String(incomingId) + '.';
  var prompt = 'Заполнить повторение, следующую дату и оценку, затем включить.';
  return [prefix, source, prompt].filter(function(value) {
    return Boolean(value);
  }).join('\n');
}

function markIncomingDraftChanged_(rowNumber) {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var headers = getHeaders_(sheet);
  var item = rowToObject_(headers, sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0]);
  if (item.ID && String(item['Формулировка задачи'] || '').trim()) {
    PropertiesService.getScriptProperties().setProperty(draftStateKey_(item.ID), JSON.stringify({
      text: String(item['Формулировка задачи']).trim(),
      type: normalizeIncomingType_(item['Предложенный тип']),
      changedAt: Date.now()
    }));
  }
}

// Отдельный однократный запуск делает ожидание предсказуемым: не нужно ждать
// следующего общего минутного обхода после того, как пользователь дописал текст.
function scheduleDraftPromotion_() {
  var properties = PropertiesService.getScriptProperties();
  var previousTriggerId = properties.getProperty('PENDING_DRAFT_PROMOTION_TRIGGER_ID');
  if (previousTriggerId) {
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      if (trigger.getUniqueId() === previousTriggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
  var trigger = ScriptApp.newTrigger('promotePendingIncomingDrafts_')
    .timeBased()
    // Apps Script может запустить триггер на доли секунды раньше границы.
    // Небольшой запас не дает пропустить двухминутную проверку и ждать еще минуту.
    .after(TASK_TRACKER_CONFIG_.draftWaitMinutes * 60 * 1000 + 10 * 1000)
    .create();
  properties.setProperty('PENDING_DRAFT_PROMOTION_TRIGGER_ID', trigger.getUniqueId());
}

function promotePendingIncomingDrafts_() {
  PropertiesService.getScriptProperties().deleteProperty('PENDING_DRAFT_PROMOTION_TRIGGER_ID');
  runTaskTrackerAutomation();
}

function isDraftReadyForPromotion_(incoming) {
  var properties = PropertiesService.getScriptProperties();
  var key = draftStateKey_(incoming.ID);
  var text = String(incoming['Формулировка задачи'] || '').trim();
  var type = normalizeIncomingType_(incoming['Предложенный тип']);
  var stored = properties.getProperty(key);
  if (!stored) {
    properties.setProperty(key, JSON.stringify({ text: text, type: type, changedAt: Date.now() }));
    return false;
  }
  var state;
  try {
    state = JSON.parse(stored);
  } catch (error) {
    properties.deleteProperty(key);
    return false;
  }
  if (state.text !== text || state.type !== type) {
    properties.setProperty(key, JSON.stringify({ text: text, type: type, changedAt: Date.now() }));
    return false;
  }
  return Date.now() - Number(state.changedAt || 0) >= TASK_TRACKER_CONFIG_.draftWaitMinutes * 60 * 1000;
}

function draftStateKey_(incomingId) {
  return 'incoming_draft_' + String(incomingId);
}

function clearDraftState_(incomingId) {
  PropertiesService.getScriptProperties().deleteProperty(draftStateKey_(incomingId));
}

function getExistingTasks_(sheet, headers) {
  if (!hasDataRows_(sheet)) {
    return [];
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().map(function(row) {
    return rowToObject_(headers, row);
  }).filter(function(task) {
    return task.ID && task.Заголовок;
  });
}

// Каждая задача получает один стартовый шаг; дополнительные шаги пользователь добавляет вручную.
function ensureStarterSubtasks_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var taskHeaders = getHeaders_(tasksSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var existingTaskIds = {};

  if (hasDataRows_(subtasksSheet)) {
    subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(subtasksSheet), subtaskHeaders.length).getValues().forEach(function(row) {
      var subtask = rowToObject_(subtaskHeaders, row);
      if (isTaskSourceId_(subtask['Источник ID'])) {
        existingTaskIds[String(subtask['Источник ID'])] = true;
      }
    });
  }

  var now = nowIso_();
  var additions = getExistingTasks_(tasksSheet, taskHeaders).filter(function(task) {
    return !existingTaskIds[String(task.ID)];
  }).map(function(task) {
    return objectToRow_(subtaskHeaders, {
      ID: 'subtask_starter_' + String(task.ID).replace(/^task_/, ''),
      Задача: task.Заголовок,
      'Источник ID': task.ID,
      'Проект / область': task['Проект / область'] || '',
      Название: task.Заголовок,
      Статус: 'new',
      Обновлено: now,
      Заметки: 'Базовая подзадача создана автоматически.'
    });
  });
  if (additions.length) {
    subtasksSheet.getRange(nextDataRow_(subtasksSheet), 1, additions.length, subtaskHeaders.length).setValues(additions);
  }
}

function findDuplicateTask_(tasks, title) {
  return tasks.filter(function(task) {
    return taskTextsOverlap_(title, task.Заголовок);
  })[0] || null;
}

function taskTextsOverlap_(left, right) {
  var first = normalizeTaskText_(left);
  var second = normalizeTaskText_(right);
  var firstNumbers = (first.match(/\d+/g) || []).sort();
  var secondNumbers = (second.match(/\d+/g) || []).sort();
  // Одинаковая тема не означает один результат: «9 дней» и «40 дней» - разные дела.
  // Числа и даты считаем частью смысла формулировки, а не шумом при поиске дублей.
  if (firstNumbers.join('|') !== secondNumbers.join('|')) {
    return false;
  }
  if (first.length < 10 || second.length < 10) {
    return false;
  }
  if (first.indexOf(second) !== -1 || second.indexOf(first) !== -1) {
    return true;
  }
  var firstWords = uniqueLongWords_(first);
  var secondWords = uniqueLongWords_(second);
  if (firstWords.length < 2 || secondWords.length < 2) {
    return false;
  }
  var common = firstWords.filter(function(word) {
    return secondWords.indexOf(word) !== -1;
  }).length;
  return common / Math.min(firstWords.length, secondWords.length) >= 0.8;
}

function normalizeTaskText_(text) {
  return String(text || '').toLowerCase().replace(/[^a-zа-яё0-9]+/gi, ' ').trim().replace(/\s+/g, ' ');
}

function uniqueLongWords_(text) {
  return String(text).split(' ').filter(function(word, index, words) {
    return word.length >= 4 && words.indexOf(word) === index;
  });
}

function refreshSubtaskChecks() {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var headers = getHeaders_(sheet);
  if (!hasDataRows_(sheet)) {
    return;
  }
  var rows = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues();
  var checkColumn = columnNumber_(headers, 'Проверка планирования');
  var checks = rows.map(function(row) {
    var item = rowToObject_(headers, row);
    return [subtaskCheck_(item)];
  });
  sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, checkColumn, checks.length, 1).setValues(checks);
}

// Статус задачи - агрегат ее обычных подзадач. Рутины сюда не входят.
function syncTaskStatusesFromSubtasks_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  if (!hasDataRows_(tasksSheet)) {
    return 0;
  }

  var taskHeaders = getHeaders_(tasksSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var states = {};
  if (hasDataRows_(subtasksSheet)) {
    subtasksSheet.getRange(
      TASK_TRACKER_CONFIG_.dataStartRow,
      1,
      dataRowCount_(subtasksSheet),
      subtaskHeaders.length
    ).getValues().forEach(function(row) {
      var subtask = rowToObject_(subtaskHeaders, row);
      var taskId = String(subtask['Источник ID'] || '').trim();
      if (!isTaskSourceId_(taskId) || !subtask.ID) {
        return;
      }
      if (!states[taskId]) {
        states[taskId] = { total: 0, closed: 0, cancelled: 0, skipped: 0, inProgress: 0, blocked: 0 };
      }
      var state = states[taskId];
      state.total += 1;
      var status = String(subtask.Статус || '').trim().toLowerCase();
      if (status === 'done') {
        state.closed += 1;
      } else if (status === 'cancelled') {
        state.cancelled += 1;
      } else if (status === 'skipped') {
        state.skipped += 1;
      } else if (status === 'in_progress') {
        state.inProgress += 1;
      } else if (status === 'blocked') {
        state.blocked += 1;
      }
    });
  }

  var updated = 0;
  tasksSheet.getRange(
    TASK_TRACKER_CONFIG_.dataStartRow,
    1,
    dataRowCount_(tasksSheet),
    taskHeaders.length
  ).getValues().forEach(function(row, index) {
    var task = rowToObject_(taskHeaders, row);
    var nextStatus = taskStatusFromSubtasks_(states[String(task.ID || '').trim()]);
    if (String(task.Статус || '').trim().toLowerCase() === nextStatus) {
      return;
    }
    setObjectFields_(tasksSheet, taskHeaders, index + TASK_TRACKER_CONFIG_.dataStartRow, {
      Статус: nextStatus,
      Обновлено: nowIso_(),
      'Дата закрытия': nextStatus === 'done' ? nowIso_() : ''
    });
    updated += 1;
  });
  return updated;
}

function taskStatusFromSubtasks_(state) {
  if (!state || !state.total) {
    return 'new';
  }
  if (state.closed === state.total) {
    return 'done';
  }
  if (state.cancelled === state.total) {
    return 'cancelled';
  }
  if (state.skipped === state.total) {
    return 'skipped';
  }
  if (state.closed + state.cancelled + state.skipped === state.total) {
    return state.closed ? 'done' : (state.cancelled ? 'cancelled' : 'skipped');
  }
  // Закрытый шаг при оставшихся активных означает, что работа уже началась.
  if (state.inProgress || state.closed) {
    return 'in_progress';
  }
  if (state.blocked === state.total - state.closed - state.cancelled - state.skipped) {
    return 'blocked';
  }
  return 'new';
}

// Оставлено для уже созданных триггеров и старого пункта меню.
function closeTasksWithCompletedSubtasks_() {
  return syncTaskStatusesFromSubtasks_();
}

function sendDailyIncomingReminder() {
  sendDailyIncomingReminderIfDue_();
}

function sendDailyIncomingReminderIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var currentTime = Utilities.formatDate(now, timezone, 'HH:mm');
  if (currentTime < '08:00' || currentTime >= '08:30') {
    return;
  }
  var properties = PropertiesService.getScriptProperties();
  var dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  if (properties.getProperty('DAILY_INCOMING_REMINDER_DATE') === dateKey) {
    return;
  }
  var stale = getStaleIncoming_();
  if (!stale.length) {
    return;
  }
  var text = 'Доброе утро! ☀️\nВо Входящих ждут внимания: ' + stale.length + '. Давай разберем их спокойно, когда будет удобно.\n<a href="' + trackerUrl_() + '">Открыть Входящие</a>';
  requestYandexNotification_(text);
  properties.setProperty('DAILY_INCOMING_REMINDER_DATE', dateKey);
}

function sendShoppingListIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var currentTime = Utilities.formatDate(now, timezone, 'HH:mm');
  if (currentTime < '08:30' || currentTime >= '09:00') {
    return;
  }
  var properties = PropertiesService.getScriptProperties();
  var dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  if (properties.getProperty('DAILY_SHOPPING_LIST_SENT_DATE') === dateKey) {
    return;
  }
  var text = buildShoppingListText_(spreadsheet);
  if (!text) {
    return;
  }
  requestYandexNotification_(text);
  properties.setProperty('DAILY_SHOPPING_LIST_SENT_DATE', dateKey);
}

function buildShoppingListText_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.shoppingSheet);
  if (!sheet || !hasDataRows_(sheet)) {
    return '';
  }
  var headers = getHeaders_(sheet);
  var items = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().map(function(row, index) {
    var item = rowToObject_(headers, row);
    item.__statusUrl = shoppingStatusUrl_(TASK_TRACKER_CONFIG_.dataStartRow + index);
    return item;
  }).filter(function(item) {
    return item.ID && item.Покупка && String(item.Статус || '').toLowerCase() === 'new';
  });
  if (!items.length) {
    return '';
  }
  return '🛒 <b>Список покупок</b>\nСобрала, чтобы ничего не потерялось.\n\n' + items.map(function(item) {
    return '• <a href="' + item.__statusUrl + '">Куплено</a> ' + escapeTelegramHtml_(item.Покупка);
  }).join('\n');
}

function shoppingStatusUrl_(rowNumber) {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.shoppingSheet);
  return 'https://docs.google.com/spreadsheets/d/' + TASK_TRACKER_CONFIG_.spreadsheetId +
    '/edit#gid=' + sheet.getSheetId() + '&range=C' + rowNumber;
}

// Минутный триггер дает более точное время, чем суточный Apps Script-триггер.
function sendDailyPlanIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var currentTime = Utilities.formatDate(now, timezone, 'HH:mm');
  if (currentTime < '07:30' || currentTime >= '08:00') {
    return;
  }

  var dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('DAILY_PLAN_SENT_DATE') === dateKey) {
    return;
  }

  requestYandexNotifications_(buildDailyPlanTexts_(spreadsheet, now, timezone));
  properties.setProperty('DAILY_PLAN_SENT_DATE', dateKey);
}

function sendEveningPlanIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var currentTime = Utilities.formatDate(now, timezone, 'HH:mm');
  var startTime = ('0' + TASK_TRACKER_CONFIG_.eveningPlanHour).slice(-2) + ':00';
  var endTime = ('0' + TASK_TRACKER_CONFIG_.eveningPlanHour).slice(-2) + ':30';
  if (currentTime < startTime || currentTime >= endTime) {
    return;
  }

  var dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('EVENING_PLAN_SENT_DATE') === dateKey) {
    return;
  }

  requestYandexNotifications_(buildEveningPlanTexts_(spreadsheet, now, timezone));
  properties.setProperty('EVENING_PLAN_SENT_DATE', dateKey);
}

function buildDailyPlanTexts_(spreadsheet, now, timezone) {
  return buildPlanSummaryTexts_(
    spreadsheet,
    now,
    timezone,
    'Доброе утро! ☀️\nПлан на ' + Utilities.formatDate(now, timezone, 'd.MM') + ' уже собран.'
  );
}

function buildEveningPlanTexts_(spreadsheet, now, timezone) {
  return buildPlanSummaryTexts_(
    spreadsheet,
    now,
    timezone,
    'Добрый вечер! 🌙\nСегодня в работе были задачи ниже. Не забудь закрыть выполненное или перенести незавершенное.',
    false
  );
}

function buildPlanSummaryTexts_(spreadsheet, now, timezone, greeting, includeLoadSummary) {
  var sheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var headers = getHeaders_(sheet);
  var todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  var yesterdayKey = Utilities.formatDate(yesterday, timezone, 'yyyy-MM-dd');
  var routines = [];
  var today = [];
  var overdueYesterday = [];

  if (hasDataRows_(sheet)) {
    sheet.getRange(
      TASK_TRACKER_CONFIG_.dataStartRow,
      1,
      dataRowCount_(sheet),
      headers.length
    ).getValues().forEach(function(row, index) {
      var subtask = rowToObject_(headers, row);
      if (!subtask.ID || isClosed_(subtask.Статус)) {
        return;
      }
      subtask.__statusUrl = subtaskStatusUrl_(TASK_TRACKER_CONFIG_.dataStartRow + index);
      var dateKey = dateKeyInTimezone_(subtask.Дата, timezone);
      var isRoutine = isRoutineSubtask_(subtask);
      if (dateKey === todayKey) {
        (isRoutine ? routines : today).push(subtask);
      } else if (dateKey === yesterdayKey) {
        overdueYesterday.push(subtask);
      }
    });
  }

  var overdueBlock = overdueYesterday.length ?
    formatDailyPlanBlock_('⚠️ Просрочено вчера', overdueYesterday) :
    '✨ <b>Просрочено вчера</b>\nНичего нет. Отличный темп!';
  var summary = '';
  if (includeLoadSummary !== false) {
    var dayLimitHours = Number(spreadsheet.getSheetByName('План дня').getRange('B2').getValue()) || 0;
    summary = buildPlanLoadSummary_(routines, today, dayLimitHours * 60);
  } else {
    summary = buildCompletedTodaySummary_(spreadsheet, headers, todayKey, timezone);
  }
  var compactPlan = [
    greeting,
    summary,
    formatDailyPlanBlock_('🌿 Рутины', routines),
    formatDailyPlanBlock_('🎯 Сегодня', today),
    overdueBlock
  ].join('\n\n');
  // Если весь план помещается, он остается одним сообщением: так его проще читать.
  if (compactPlan.length <= 3600) {
    return [compactPlan];
  }

  var intro = greeting + '\n\n' + summary;
  var blocks = splitDailyPlanBlock_('🌿 Рутины', routines, Math.max(1000, 3600 - intro.length - 2))
    .concat(splitDailyPlanBlock_('🎯 Сегодня', today));
  blocks = blocks.concat(overdueYesterday.length ?
    splitDailyPlanBlock_('⚠️ Просрочено вчера', overdueYesterday) :
    [overdueBlock]);
  blocks[0] = intro + '\n\n' + blocks[0];
  return blocks;
}

function buildCompletedTodaySummary_(spreadsheet, subtaskHeaders, todayKey, timezone) {
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var completedSubtasks = 0;
  var completedRoutines = 0;
  var completedTasks = 0;

  if (hasDataRows_(subtasksSheet)) {
    subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(subtasksSheet), subtaskHeaders.length).getValues().forEach(function(row) {
      var subtask = rowToObject_(subtaskHeaders, row);
      if (String(subtask.Статус || '').toLowerCase() === 'done' &&
          dateKeyInTimezone_(subtask.Обновлено, timezone) === todayKey) {
        if (isRoutineSubtask_(subtask)) {
          completedRoutines += 1;
        } else {
          completedSubtasks += 1;
        }
      }
    });
  }
  if (hasDataRows_(tasksSheet)) {
    var taskHeaders = getHeaders_(tasksSheet);
    tasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(tasksSheet), taskHeaders.length).getValues().forEach(function(row) {
      var task = rowToObject_(taskHeaders, row);
      if (String(task.Статус || '').toLowerCase() === 'done' &&
          dateKeyInTimezone_(task.Обновлено, timezone) === todayKey) {
        completedTasks += 1;
      }
    });
  }

  var total = completedSubtasks + completedRoutines;
  var lines = ['✅ <b>Итоги дня</b>', 'Закрыто дел: ' + total + '.'];
  if (completedRoutines || completedSubtasks) {
    lines.push('🌿 Рутин: ' + completedRoutines + '; 🎯 подзадач: ' + completedSubtasks + '.');
  }
  if (completedTasks) {
    lines.push('Полностью завершено задач: ' + completedTasks + '.');
  }
  return lines.join('\n');
}

function touchSubtaskUpdated_(range) {
  touchRowsUpdated_(range.getSheet(), getHeaders_(range.getSheet()), range);
}

function touchRowsUpdated_(sheet, headers, range) {
  var firstRow = Math.max(range.getRow(), TASK_TRACKER_CONFIG_.dataStartRow);
  var count = Math.max(0, range.getLastRow() - firstRow + 1);
  if (!count) {
    return;
  }
  var updatedColumn = columnNumber_(headers, 'Обновлено');
  var now = nowIso_();
  var values = [];
  for (var index = 0; index < count; index += 1) {
    values.push([now]);
  }
  sheet.getRange(firstRow, updatedColumn, count, 1).setValues(values);
}

function updateClosedDates_(sheet, headers, range, statusHeader, closedDateHeader) {
  var firstRow = Math.max(range.getRow(), TASK_TRACKER_CONFIG_.dataStartRow);
  var count = Math.max(0, range.getLastRow() - firstRow + 1);
  if (!count) {
    return;
  }
  var statuses = sheet.getRange(firstRow, columnNumber_(headers, statusHeader), count, 1).getValues();
  var closedDateColumn = columnNumber_(headers, closedDateHeader);
  var existingDates = sheet.getRange(firstRow, closedDateColumn, count, 1).getValues();
  var now = nowIso_();
  var values = statuses.map(function(row, index) {
    return [String(row[0] || '').toLowerCase() === 'done' ? (existingDates[index][0] || now) : ''];
  });
  sheet.getRange(firstRow, closedDateColumn, count, 1).setValues(values);
}

function buildPlanLoadSummary_(routines, workSubtasks, limitMinutes) {
  var routineMinutes = 0;
  var workMinutes = 0;
  var withoutEstimate = 0;
  var projectMinutes = {};
  routines.forEach(function(subtask) {
    var minutes = estimateMinutes_(subtask['Оценка, мин']);
    routineMinutes += minutes;
    if (!minutes) {
      withoutEstimate += 1;
    }
  });
  workSubtasks.forEach(function(subtask) {
    var minutes = estimateMinutes_(subtask['Оценка, мин']);
    var project = String(subtask['Проект / область'] || '').trim() || 'Без проекта';
    workMinutes += minutes;
    projectMinutes[project] = (projectMinutes[project] || 0) + minutes;
    if (!minutes) {
      withoutEstimate += 1;
    }
  });

  var totalMinutes = routineMinutes + workMinutes;
  var lines = [
    '⏱ <b>Нагрузка на сегодня</b>',
    'Всего: ' + formatMinutes_(totalMinutes) + (limitMinutes ? ' из ' + formatMinutes_(limitMinutes) : '') + '.',
    '🌿 Рутины: ' + formatMinutes_(routineMinutes) + '.',
    '🎯 Работа: ' + formatMinutes_(workMinutes) + '.'
  ];
  var projects = Object.keys(projectMinutes).sort(function(left, right) {
    return projectMinutes[right] - projectMinutes[left];
  });
  if (projects.length) {
    lines.push('По проектам: ' + projects.map(function(project) {
      return escapeTelegramHtml_(project) + ' — ' + formatMinutes_(projectMinutes[project]);
    }).join('; ') + '.');
  }
  if (limitMinutes && totalMinutes > limitMinutes) {
    lines.push('⚠️ Перелимит: +' + formatMinutes_(totalMinutes - limitMinutes) + '.');
  }
  if (withoutEstimate) {
    lines.push('Без оценки: ' + withoutEstimate + '.');
  }
  return lines.join('\n');
}

function estimateMinutes_(value) {
  var normalized = String(value === undefined || value === null ? '' : value).replace(',', '.').trim();
  var minutes = Number(normalized);
  return isFinite(minutes) && minutes > 0 ? minutes : 0;
}

function formatMinutes_(minutes) {
  var total = Math.round(Number(minutes) || 0);
  var hours = Math.floor(total / 60);
  var remainder = total % 60;
  if (!hours) {
    return remainder + ' мин';
  }
  return hours + ' ч' + (remainder ? ' ' + remainder + ' мин' : '');
}

function subtaskStatusUrl_(rowNumber) {
  var spreadsheet = getTrackerSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  return 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() +
    '/edit#gid=' + sheet.getSheetId() + '&range=E' + rowNumber;
}

function splitDailyPlanBlock_(title, subtasks, maxLength) {
  if (!subtasks.length) {
    return ['<b>' + title + '</b>\nПока ничего нет.'];
  }
  var itemLines = dailyPlanItemLines_(subtasks);
  var limit = Number(maxLength) || 3600;
  var messages = [];
  var current = '<b>' + title + '</b>';
  itemLines.forEach(function(line) {
    // Telegram ограничивает одно сообщение 4096 символами. Делим только между
    // подзадачами, поэтому формулировка и ссылка никогда не обрезаются.
    if ((current + '\n' + line).length > limit && current !== '<b>' + title + '</b>') {
      messages.push(current);
      current = '<b>' + title + ' (продолжение)</b>';
    }
    current += '\n' + line;
  });
  messages.push(current);
  return messages;
}

function formatDailyPlanBlock_(title, subtasks) {
  if (!subtasks.length) {
    return '<b>' + title + '</b>\nПока ничего нет.';
  }
  return '<b>' + title + '</b>\n' + dailyPlanItemLines_(subtasks).join('\n');
}

function dailyPlanItemLines_(subtasks) {
  return subtasks.map(function(subtask) {
    var name = String(subtask.Название || subtask.Задача || '').trim();
    var closeUrl = subtask.__statusUrl;
    var closeLink = closeUrl ? '<a href="' + closeUrl + '">Закрыть</a> ' : '';
    return '• ' + closeLink + escapeTelegramHtml_(name);
  });
}

function dateKeyInTimezone_(value, timezone) {
  var date = routineDate_(value);
  return date ? Utilities.formatDate(date, timezone, 'yyyy-MM-dd') : '';
}

function escapeTelegramHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function configureYandexNotifier(functionUrl, relaySecret) {
  if (!functionUrl || !relaySecret) {
    throw new Error('Нужны URL Yandex Function и служебный ключ.');
  }
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('YANDEX_FUNCTION_URL', String(functionUrl));
  properties.setProperty('RELAY_SECRET', String(relaySecret));
}

function requestYandexNotification_(text) {
  requestYandexNotifications_([text]);
}

function requestYandexNotifications_(texts) {
  var messages = (texts || []).filter(function(text) {
    return Boolean(String(text || '').trim());
  });
  if (!messages.length) {
    return;
  }
  var properties = PropertiesService.getScriptProperties();
  var functionUrl = properties.getProperty('YANDEX_FUNCTION_URL');
  var relaySecret = properties.getProperty('RELAY_SECRET');
  if (!functionUrl || !relaySecret) {
    throw new Error('Напоминания не настроены: отсутствуют Script Properties Yandex Function.');
  }
  var response = UrlFetchApp.fetch(functionUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Relay-Secret': relaySecret },
    payload: JSON.stringify({ mode: 'send_notifications', texts: messages }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    throw new Error('Yandex Function не приняла уведомление: ' + response.getContentText());
  }
}

function requestYandexWeeklyReflection_(weekStart, cutoffKey) {
  var properties = PropertiesService.getScriptProperties();
  var functionUrl = properties.getProperty('YANDEX_FUNCTION_URL');
  var relaySecret = properties.getProperty('RELAY_SECRET');
  if (!functionUrl || !relaySecret) {
    throw new Error('Недельная рефлексия не настроена: отсутствуют Script Properties Yandex Function.');
  }
  var response = UrlFetchApp.fetch(functionUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Relay-Secret': relaySecret },
    payload: JSON.stringify({ mode: 'weekly_reflection_prompt', week_start: weekStart, cutoff: cutoffKey }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 300) {
    throw new Error('Yandex Function не смогла начать недельную рефлексию: ' + response.getContentText());
  }
}

function getStaleIncoming_() {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var headers = getHeaders_(sheet);
  if (!hasDataRows_(sheet)) {
    return [];
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().map(function(row) {
    return rowToObject_(headers, row);
  }).filter(function(item) {
    var status = String(item['Статус разбора'] || '').toLowerCase();
    // Покупка уже обработана отдельным сценарием, даже если строка остается «новой»
    // для сохранения пользовательского цветного выпадающего списка.
    return !item['Формулировка задачи'] &&
      !String(item['Созданные покупки ID'] || '').trim() &&
      status === 'new';
  });
}

function ensureTaskTrackerSchema_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var incomingSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  ensureIncomingColumns_(incomingSheet);
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  ensureSubtaskSourceIdColumn_(subtasksSheet);
  ensureSubtaskTaskTitleColumn_(subtasksSheet);
  ensureSubtaskProjectColumn_(subtasksSheet);
  ensureSubtaskCheckColumn_(subtasksSheet);
  var shoppingSheet = ensureShoppingSheet_(spreadsheet);
  ensurePlanDayShoppingSection_(spreadsheet, shoppingSheet, subtasksSheet);
  var routinesSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.routinesSheet);
  ensureRoutineColumns_(routinesSheet);
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  ensureColumn_(tasksSheet, 'Дата закрытия');
  ensureColumn_(subtasksSheet, 'Дата закрытия');
  ensureColumn_(shoppingSheet, 'Дата закрытия');
  ensureSummarySheets_(spreadsheet);
  migrateStatusModel_(spreadsheet, incomingSheet, tasksSheet, subtasksSheet, routinesSheet, shoppingSheet);
  var referencesSheet = ensureReferencesSheet_(spreadsheet, incomingSheet, tasksSheet, routinesSheet);
  applyReferenceValidation_(incomingSheet, 'Статус разбора', referencesSheet, 'Статусы работы');
  applyReferenceValidation_(incomingSheet, 'Предложенный тип', referencesSheet, 'Типы');
  applyReferenceValidation_(shoppingSheet, 'Статус', referencesSheet, 'Статусы работы');
  applyReferenceValidation_(tasksSheet, 'Статус', referencesSheet, 'Статусы работы');
  applyReferenceValidation_(tasksSheet, 'Приоритет', referencesSheet, 'Приоритеты');
  applyReferenceValidation_(tasksSheet, 'Проект / область', referencesSheet, 'Проекты / области');
  applyReferenceValidation_(subtasksSheet, 'Статус', referencesSheet, 'Статусы работы');
  applyReferenceValidation_(routinesSheet, 'Статус', referencesSheet, 'Статусы рутин');
  applyReferenceValidation_(routinesSheet, 'Повторение', referencesSheet, 'Повторение');
  applyReferenceValidation_(routinesSheet, 'Категория', referencesSheet, 'Категории рутин');
}

function ensureSummarySheets_(spreadsheet) {
  var summaries = spreadsheet.getSheetByName('Сводки');
  var weekly = spreadsheet.getSheetByName('Итоги недели');
  var dailyHeaders = ['ID', 'Дата', 'Тип', 'Период', 'Закрыто подзадач', 'Закрыто рутин', 'Закрыто задач', 'Куплено', 'Закрыто, мин', 'Запланировано, мин', 'Новых входящих', 'Разобрано входящих', 'Просрочено на конец дня', 'По проектам', 'Создано'];
  var weeklyHeaders = ['ID', 'Неделя', 'Главный результат', 'Энергия и фокус', 'Что сработало', 'Напряжение и откладывание', 'Урок недели', 'Коммуникация', 'Система работы', 'Рост', 'Следующий цикл', 'Закрыто подзадач', 'Закрыто рутин', 'Закрыто задач', 'Куплено', 'Закрыто, мин', 'Запланировано, мин', 'Просрочено на конец недели', 'По проектам', 'Создано'];
  migrateWeeklyReflectionLayout_(weekly, weeklyHeaders);
  [
    { sheet: summaries, headers: dailyHeaders },
    { sheet: weekly, headers: weeklyHeaders }
  ].forEach(function(item) {
    item.sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, item.headers.length).setValues([item.headers]);
    item.sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, item.headers.length).setFontWeight('bold').setBackground('#e6e6e6');
    item.sheet.setFrozenRows(TASK_TRACKER_CONFIG_.headerRow);
  });
  // Рефлексию недели пользователь пишет сам; отделяем ее цветом от автоматических полей.
  weekly.getRange(TASK_TRACKER_CONFIG_.headerRow, 3, 1, 9).setBackground('#fbc965');
}

function migrateWeeklyReflectionLayout_(sheet, newHeaders) {
  var oldHeaders = getHeaders_(sheet);
  var matches = oldHeaders.length === newHeaders.length && oldHeaders.every(function(header, index) {
    return header === newHeaders[index];
  });
  if (matches) {
    return;
  }
  var rows = hasDataRows_(sheet)
    ? sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), oldHeaders.length).getValues()
    : [];
  var migratedRows = rows.map(function(row) {
    var item = rowToObject_(oldHeaders, row);
    var values = {
      ID: item.ID,
      Неделя: item.Неделя,
      'Главный результат': item['Главный результат'] || item['Главный итог'],
      'Энергия и фокус': item['Энергия и фокус'],
      'Что сработало': item['Что сработало'] || item['Зеленое'],
      'Напряжение и откладывание': mergeReflectionValues_([
        ['Требует внимания', item['Желтое']],
        ['Откладывалось', item['Красное']]
      ]),
      'Урок недели': item['Урок недели'],
      Коммуникация: item.Коммуникация,
      'Система работы': item['Система работы'] || item['Что меняем'],
      Рост: item.Рост,
      'Следующий цикл': mergeReflectionValues_([
        ['Переношу', item['Что переносим']],
        ['Повторить / не повторять', item['Повторить / не повторять']],
        ['Главный фокус', item['Фокус следующей недели']],
        ['Поддержка', item.Поддержка]
      ]),
      'Закрыто подзадач': item['Закрыто подзадач'],
      'Закрыто рутин': item['Закрыто рутин'],
      'Закрыто задач': item['Закрыто задач'],
      Куплено: item.Куплено,
      'Закрыто, мин': item['Закрыто, мин'],
      'Запланировано, мин': item['Запланировано, мин'],
      'Просрочено на конец недели': item['Просрочено на конец недели'],
      'По проектам': item['По проектам'],
      Создано: item.Создано
    };
    return objectToRow_(newHeaders, values);
  });
  var occupiedColumns = Math.max(oldHeaders.length, newHeaders.length);
  var rowsToClear = Math.max(sheet.getLastRow() - TASK_TRACKER_CONFIG_.headerRow + 1, 1);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, rowsToClear, occupiedColumns).clearContent();
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, newHeaders.length).setValues([newHeaders]);
  if (migratedRows.length) {
    sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, migratedRows.length, newHeaders.length).setValues(migratedRows);
  }
  if (oldHeaders.length > newHeaders.length) {
    sheet.deleteColumns(newHeaders.length + 1, oldHeaders.length - newHeaders.length);
  }
}

function mergeReflectionValues_(items) {
  return items.filter(function(item) {
    return String(item[1] || '').trim();
  }).map(function(item) {
    return item[0] + ': ' + item[1];
  }).join('\n\n');
}

function createDailySummaryIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var time = Utilities.formatDate(now, timezone, 'HH:mm');
  if (time < '23:30') {
    return;
  }
  createDailySummary_(spreadsheet, Utilities.formatDate(now, timezone, 'yyyy-MM-dd'), timezone);
}

function createDailySummaryNow() {
  var spreadsheet = getTrackerSpreadsheet_();
  ensureTaskTrackerSchema_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  createDailySummary_(spreadsheet, Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd'), timezone);
}

function createDailySummary_(spreadsheet, dateKey, timezone) {
  var sheet = spreadsheet.getSheetByName('Сводки');
  var headers = getHeaders_(sheet);
  var metrics = buildSummaryMetrics_(spreadsheet, dateKey, timezone);
  var rowNumber = findSummaryRow_(sheet, headers, 'день', dateKey, timezone);
  var values = {
    ID: 'summary_day_' + dateKey,
    Дата: dateKey,
    Тип: 'день',
    Период: dateKey,
    'Закрыто подзадач': metrics.subtasks,
    'Закрыто рутин': metrics.routines,
    'Закрыто задач': metrics.tasks,
    Куплено: metrics.purchases,
    'Закрыто, мин': metrics.closedMinutes,
    'Запланировано, мин': metrics.plannedMinutes,
    'Новых входящих': metrics.incoming,
    'Разобрано входящих': metrics.processedIncoming,
    'Просрочено на конец дня': metrics.overdue,
    'По проектам': formatProjectMetrics_(metrics.projects),
    Создано: nowIso_()
  };
  if (rowNumber) {
    setObjectFields_(sheet, headers, rowNumber, values);
  } else {
    sheet.getRange(nextDataRow_(sheet), 1, 1, headers.length).setValues([objectToRow_(headers, values)]);
  }
}

function buildSummaryMetrics_(spreadsheet, dateKey, timezone) {
  var result = { subtasks: 0, routines: 0, tasks: 0, purchases: 0, closedMinutes: 0, plannedMinutes: 0, incoming: 0, processedIncoming: 0, overdue: 0, projects: {} };
  var subtasks = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var subtaskHeaders = getHeaders_(subtasks);
  if (hasDataRows_(subtasks)) {
    subtasks.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(subtasks), subtaskHeaders.length).getValues().forEach(function(row) {
      var item = rowToObject_(subtaskHeaders, row);
      var status = String(item.Статус || '').toLowerCase();
      var plannedDate = dateKeyInTimezone_(item.Дата, timezone);
      if (plannedDate === dateKey) {
        result.plannedMinutes += estimateMinutes_(item['Оценка, мин']);
      }
      if (!isClosed_(status) && plannedDate && plannedDate < dateKey) {
        result.overdue += 1;
      }
      if (status !== 'done' || dateKeyInTimezone_(item['Дата закрытия'], timezone) !== dateKey) {
        return;
      }
      var minutes = estimateMinutes_(item['Оценка, мин']);
      result.closedMinutes += minutes;
      if (isRoutineSubtask_(item)) {
        result.routines += 1;
      } else {
        result.subtasks += 1;
        var project = String(item['Проект / область'] || '').trim() || 'Без проекта';
        result.projects[project] = (result.projects[project] || 0) + minutes;
      }
    });
  }
  var tasks = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var taskHeaders = getHeaders_(tasks);
  if (hasDataRows_(tasks)) {
    tasks.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(tasks), taskHeaders.length).getValues().forEach(function(row) {
      var item = rowToObject_(taskHeaders, row);
      if (String(item.Статус || '').toLowerCase() === 'done' && dateKeyInTimezone_(item['Дата закрытия'], timezone) === dateKey) {
        result.tasks += 1;
      }
    });
  }
  var purchases = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.shoppingSheet);
  var purchaseHeaders = getHeaders_(purchases);
  if (hasDataRows_(purchases)) {
    purchases.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(purchases), purchaseHeaders.length).getValues().forEach(function(row) {
      var item = rowToObject_(purchaseHeaders, row);
      if (String(item.Статус || '').toLowerCase() === 'done' && dateKeyInTimezone_(item['Дата закрытия'], timezone) === dateKey) {
        result.purchases += 1;
      }
    });
  }
  var incoming = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  var incomingHeaders = getHeaders_(incoming);
  if (hasDataRows_(incoming)) {
    incoming.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(incoming), incomingHeaders.length).getValues().forEach(function(row) {
      var item = rowToObject_(incomingHeaders, row);
      if (dateKeyInTimezone_(item['Дата захвата'], timezone) === dateKey) {
        result.incoming += 1;
        if (String(item['Статус разбора'] || '').toLowerCase() === 'done') {
          result.processedIncoming += 1;
        }
      }
    });
  }
  return result;
}

function findSummaryRow_(sheet, headers, type, period, timezone) {
  if (!hasDataRows_(sheet)) {
    return 0;
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().reduce(function(found, row, index) {
    var item = rowToObject_(headers, row);
    var savedPeriod = timezone ? dateKeyInTimezone_(item.Период, timezone) : String(item.Период || '');
    return found || (String(item.Тип || '') === type && savedPeriod === period ? index + TASK_TRACKER_CONFIG_.dataStartRow : 0);
  }, 0);
}

function formatProjectMetrics_(projects) {
  return Object.keys(projects).sort(function(left, right) { return projects[right] - projects[left]; }).map(function(project) {
    return project + ' — ' + formatMinutes_(projects[project]);
  }).join('; ');
}

function createWeeklySummaryIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  if (now.getDay() !== 4 || Utilities.formatDate(now, timezone, 'HH:mm') < '12:00') {
    return;
  }
  createWeeklySummary_(spreadsheet, reflectionWeekStartKey_(now, timezone), timezone, reflectionCutoffKey_(now, timezone));
}

function createWeeklySummaryNow() {
  var spreadsheet = getTrackerSpreadsheet_();
  ensureTaskTrackerSchema_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  createWeeklySummary_(spreadsheet, reflectionWeekStartKey_(now, timezone), timezone, reflectionCutoffKey_(now, timezone));
}

function createWeeklySummary_(spreadsheet, weekStart, timezone, cutoffKey) {
  var sheet = spreadsheet.getSheetByName('Итоги недели');
  var headers = getHeaders_(sheet);
  var totals = { subtasks: 0, routines: 0, tasks: 0, purchases: 0, closedMinutes: 0, plannedMinutes: 0, overdue: 0, projects: {} };
  for (var offset = 0; offset < 7; offset += 1) {
    var date = new Date(weekStart + 'T12:00:00');
    date.setDate(date.getDate() + offset);
    var key = Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
    var metrics = buildSummaryMetrics_(spreadsheet, key, timezone);
    ['subtasks', 'routines', 'tasks', 'purchases', 'closedMinutes', 'plannedMinutes', 'overdue'].forEach(function(name) {
      totals[name] += metrics[name];
    });
    Object.keys(metrics.projects).forEach(function(project) {
      totals.projects[project] = (totals.projects[project] || 0) + metrics.projects[project];
    });
  }
  var period = weekStart + ' 12:00 — ' + cutoffKey + ' 12:00';
  var rowNumber = findWeeklySummaryRow_(sheet, headers, period);
  var values = {
    ID: 'week_' + cutoffKey,
    Неделя: period,
    'Закрыто подзадач': totals.subtasks,
    'Закрыто рутин': totals.routines,
    'Закрыто задач': totals.tasks,
    Куплено: totals.purchases,
    'Закрыто, мин': totals.closedMinutes,
    'Запланировано, мин': totals.plannedMinutes,
    'Просрочено на конец недели': totals.overdue,
    'По проектам': formatProjectMetrics_(totals.projects)
  };
  if (rowNumber) {
    setObjectFields_(sheet, headers, rowNumber, values);
  } else {
    values.Создано = nowIso_();
    sheet.getRange(nextDataRow_(sheet), 1, 1, headers.length).setValues([objectToRow_(headers, values)]);
  }
}

function findWeeklySummaryRow_(sheet, headers, period) {
  if (!hasDataRows_(sheet)) {
    return 0;
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().reduce(function(found, row, index) {
    var item = rowToObject_(headers, row);
    return found || (String(item.Неделя || '') === period ? index + TASK_TRACKER_CONFIG_.dataStartRow : 0);
  }, 0);
}

function weekStartKey_(date, timezone) {
  var start = new Date(date.getTime());
  var day = (start.getDay() + 6) % 7;
  start.setDate(start.getDate() - day);
  return Utilities.formatDate(start, timezone, 'yyyy-MM-dd');
}

// Рефлексия закрывает цикл в четверг в 12:00. Цифры опираются на семь
// завершенных дневных снимков с прошлого четверга по среду включительно.
function reflectionCutoffKey_(now, timezone) {
  var cutoff = new Date(now.getTime());
  var daysSinceThursday = (cutoff.getDay() + 3) % 7;
  var currentTime = Utilities.formatDate(cutoff, timezone, 'HH:mm');
  if (cutoff.getDay() === 4 && currentTime < '12:00') {
    daysSinceThursday = 7;
  }
  cutoff.setDate(cutoff.getDate() - daysSinceThursday);
  return Utilities.formatDate(cutoff, timezone, 'yyyy-MM-dd');
}

function reflectionWeekStartKey_(now, timezone) {
  var cutoff = reflectionCutoffKey_(now, timezone);
  var start = new Date(cutoff + 'T12:00:00');
  start.setDate(start.getDate() - 7);
  return Utilities.formatDate(start, timezone, 'yyyy-MM-dd');
}

function sendWeeklyReflectionIfDue_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  if (now.getDay() !== 4 || Utilities.formatDate(now, timezone, 'HH:mm') < '12:00') {
    return;
  }
  var cutoffKey = reflectionCutoffKey_(now, timezone);
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('WEEKLY_REFLECTION_PROMPT_CUTOFF') === cutoffKey) {
    return;
  }
  requestYandexWeeklyReflection_(reflectionWeekStartKey_(now, timezone), cutoffKey);
  properties.setProperty('WEEKLY_REFLECTION_PROMPT_CUTOFF', cutoffKey);
}

function sendWeeklyReflectionNow() {
  var spreadsheet = getTrackerSpreadsheet_();
  ensureTaskTrackerSchema_();
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var now = new Date();
  var cutoffKey = reflectionCutoffKey_(now, timezone);
  createWeeklySummary_(spreadsheet, reflectionWeekStartKey_(now, timezone), timezone, cutoffKey);
  requestYandexWeeklyReflection_(reflectionWeekStartKey_(now, timezone), cutoffKey);
}

function ensureReferencesSheet_(spreadsheet, incomingSheet, tasksSheet, routinesSheet) {
  var sheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.referencesSheet);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(TASK_TRACKER_CONFIG_.referencesSheet);
    sheet.setFrozenRows(TASK_TRACKER_CONFIG_.headerRow);
  }

  migrateReferencesLayout_(sheet);
  var headers = ['Статусы работы', 'Статусы рутин', 'Проекты / области', 'Типы', 'Приоритеты', 'Повторение', 'Категории рутин'];
  sheet.getRange(1, 1).setValue('Дополняйте списки ниже: все выпадающие поля таблицы берут значения отсюда.');
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#fbc965');
  sheet.setFrozenRows(TASK_TRACKER_CONFIG_.headerRow);

  var defaults = {
    'Статусы работы': WORK_STATUSES_,
    'Статусы рутин': ROUTINE_STATUSES_,
    'Типы': INCOMING_TYPES_,
    'Приоритеты': TASK_PRIORITIES_,
    'Повторение': ROUTINE_REPEAT_RULES_,
    'Категории рутин': collectDistinctValues_(routinesSheet, 'Категория')
  };

  headers.forEach(function(header, index) {
    if (header === 'Статусы работы' || header === 'Статусы рутин') {
      setReferenceValues_(sheet, index + 1, defaults[header]);
    } else if (header !== 'Проекты / области') {
      appendMissingReferenceValues_(sheet, index + 1, defaults[header]);
    }
    sheet.setColumnWidth(index + 1, 190);
  });
  return sheet;
}

function migrateReferencesLayout_(sheet) {
  var headers = getHeaders_(sheet);
  if (headers[0] !== 'Статусы') {
    return;
  }
  // Переносим пользовательские проекты и прочие списки правее, освобождая две колонки статусов.
  var rows = Math.max(sheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow, 1);
  var oldLists = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 2, rows, 5).getValues();
  sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 3, rows, 5).setValues(oldLists);
  sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, rows, 2).clearContent();
}

function setReferenceValues_(sheet, column, values) {
  var rows = Math.max(sheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow, 1);
  var existing = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, rows, 1)
    .getDisplayValues()
    .map(function(row) { return String(row[0] || '').trim(); })
    .filter(function(value) { return value; });
  if (!existing.length) {
    sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, values.length, 1)
      .setValues(values.map(function(value) { return [value]; }));
    return;
  }
  appendMissingReferenceValues_(sheet, column, values);
}

function collectDistinctValues_(sheet, header) {
  if (!sheet || !hasDataRows_(sheet)) {
    return [];
  }
  var column = columnNumberOrZero_(getHeaders_(sheet), header);
  if (!column) {
    return [];
  }
  return sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, dataRowCount_(sheet), 1)
    .getDisplayValues()
    .map(function(row) { return String(row[0] || '').trim(); })
    .filter(function(value) { return value; });
}

function migrateStatusModel_(spreadsheet, incomingSheet, tasksSheet, subtasksSheet, routinesSheet, shoppingSheet) {
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('STATUS_MODEL_V3_MIGRATED') === 'yes') {
    return;
  }
  migrateStatusColumn_(incomingSheet, 'Статус разбора', {
    'новое': 'new', 'задача создана': 'done', 'неактуально': 'cancelled', 'пропущено': 'skipped'
  });
  var workMap = {
    'новая': 'new', 'запланирована': 'planned', 'в работе': 'in_progress',
    'заблокирована': 'blocked', 'закрыта': 'done', 'неактуальна': 'cancelled'
  };
  migrateStatusColumn_(tasksSheet, 'Статус', workMap);
  migrateStatusColumn_(subtasksSheet, 'Статус', workMap);
  migrateStatusColumn_(shoppingSheet, 'Статус', {
    'нужно купить': 'new', 'куплено': 'done', 'неактуально': 'cancelled'
  });
  migrateStatusColumn_(routinesSheet, 'Статус', {
    'активна': 'active', 'активная': 'active', 'пауза': 'paused',
    'закрыта': 'archived', 'неактуальна': 'archived'
  });
  properties.setProperty('STATUS_MODEL_V3_MIGRATED', 'yes');
}

function migrateStatusColumn_(sheet, header, mapping) {
  if (!sheet || !hasDataRows_(sheet)) {
    return;
  }
  var column = columnNumberOrZero_(getHeaders_(sheet), header);
  if (!column) {
    return;
  }
  var rows = dataRowCount_(sheet);
  var range = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, rows, 1);
  var current = range.getValues();
  var changed = false;
  var updated = current.map(function(row) {
    var original = row[0];
    var key = String(original || '').trim().toLowerCase();
    if (mapping[key]) {
      changed = true;
      return [mapping[key]];
    }
    return [original];
  });
  if (changed) {
    range.setValues(updated);
  }
}

function appendMissingReferenceValues_(sheet, column, values) {
  var maxRows = Math.max(sheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow, 1);
  var existing = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, maxRows, 1)
    .getDisplayValues()
    .map(function(row) { return String(row[0] || '').trim(); })
    .filter(function(value) { return value; });
  var known = {};
  existing.forEach(function(value) { known[value.toLowerCase()] = true; });
  var missing = values.map(function(value) { return String(value || '').trim(); })
    .filter(function(value) {
      var key = value.toLowerCase();
      if (!value || known[key]) {
        return false;
      }
      known[key] = true;
      return true;
    });
  if (missing.length) {
    sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow + existing.length, column, missing.length, 1)
      .setValues(missing.map(function(value) { return [value]; }));
  }
}

function ensureIncomingColumns_(sheet) {
  var headers = getHeaders_(sheet);
  var oldCreated = columnNumberOrZero_(headers, 'Созданная задача ID');
  if (oldCreated) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, oldCreated).setValue('Созданные задачи ID');
    headers[oldCreated - 1] = 'Созданные задачи ID';
  }
  if (!columnNumberOrZero_(headers, 'Формулировка задачи')) {
    var statusColumn = columnNumber_(headers, 'Статус разбора');
    sheet.insertColumnBefore(statusColumn);
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, statusColumn).setValue('Формулировка задачи');
  }
  ensureColumn_(sheet, 'Созданные покупки ID');
  // Это служебное поле, его не нужно заполнять вручную.
  var updatedHeaders = getHeaders_(sheet);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, columnNumber_(updatedHeaders, 'Созданные покупки ID'))
    .setBackground('#e6e6e6');
}

function ensureShoppingSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.shoppingSheet);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(TASK_TRACKER_CONFIG_.shoppingSheet);
    sheet.setFrozenRows(TASK_TRACKER_CONFIG_.headerRow);
  }
  var headers = ['ID', 'Покупка', 'Статус', 'Источник', 'Ссылка на входящее', 'Создано', 'Обновлено', 'Заметки', 'Дата закрытия'];
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e6e6e6');
  // Оранжевый фон отмечает поля, которые пользователь заполняет вручную.
  [2, 3, 8].forEach(function(column) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, column).setBackground('#fbc965');
  });
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(8, 280);
  return sheet;
}

function ensurePlanDayShoppingSection_(spreadsheet, shoppingSheet, subtasksSheet) {
  var sheet = spreadsheet.getSheetByName('План дня');
  var separator = '\\';
  var planHeaders = ['Закрыть', 'Подзадача', 'Задача', 'Проект / область', 'Раздел', 'Тип', 'Статус', 'Дата', 'Оценка, мин', 'Дедлайн', 'Заметки', 'ID подзадачи'];
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, planHeaders.length).setValues([planHeaders]);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, planHeaders.length).setFontWeight('bold').setBackground('#e6e6e6');
  sheet.setFrozenRows(TASK_TRACKER_CONFIG_.headerRow);
  sheet.setFrozenColumns(1);
  [1, 2, 3, 4].forEach(function(column) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, column).setBackground('#d9ead3');
  });
  var firstRow = TASK_TRACKER_CONFIG_.dataStartRow;
  var rows = sheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow;
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var subtaskProjectColumn = columnLetter_(columnNumber_(subtaskHeaders, 'Проект / область'));
  var emptyRow = Array(10).fill('""').join(separator);
  var header = function(title) {
    return ['""', '""', '""', '"' + title + '"'].concat(Array(6).fill('""')).join(separator);
  };
  var subtaskData = [
    '\'Подзадачи\'!$D$12:$D',
    '\'Подзадачи\'!$B$12:$B',
    "'Подзадачи'!$" + subtaskProjectColumn + "$12:$" + subtaskProjectColumn,
    'IF(\'Подзадачи\'!$A$12:$A<>"";"";"")',
    'IF(REGEXMATCH(\'Подзадачи\'!$C$12:$C;"^routine_");"Рутина";"Подзадача")',
    '\'Подзадачи\'!$E$12:$E',
    '\'Подзадачи\'!$F$12:$F',
    '\'Подзадачи\'!$G$12:$G',
    'IFERROR(VLOOKUP(\'Подзадачи\'!$C$12:$C;\'Задачи\'!$A$12:$F;6;FALSE);"")',
    '\'Подзадачи\'!$M$12:$M'
  ].join(separator);
  var commonCriteria = '\'Подзадачи\'!$A$12:$A<>"";\'Подзадачи\'!$F$12:$F<>"";\'Подзадачи\'!$E$12:$E<>"done";\'Подзадачи\'!$E$12:$E<>"cancelled";\'Подзадачи\'!$E$12:$E<>"skipped"';
  var dateBlock = function(title, criterion) {
    return '{' + header(title) + '};IFERROR(SORT(FILTER({' + subtaskData + '};' + commonCriteria + ';' + criterion + ');7;TRUE);{' + emptyRow + '})';
  };
  var shoppingData = [
    '\'Список покупок\'!$B$12:$B',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    'IF(\'Список покупок\'!$A$12:$A<>"";"Покупка";"")',
    '\'Список покупок\'!$C$12:$C',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    'IF(\'Список покупок\'!$A$12:$A<>"";"";"")',
    '\'Список покупок\'!$H$12:$H'
  ].join(separator);
  var shoppingBlock = '{' + header('🛒 Покупки') + '};IFERROR(FILTER({' + shoppingData + '};\'Список покупок\'!$A$12:$A<>"";\'Список покупок\'!$C$12:$C="new");{' + emptyRow + '})';
  var formula = '=VSTACK(' + [
    dateBlock('Просрочено', '\'Подзадачи\'!$F$12:$F<TODAY()'),
    dateBlock('Сегодня', '\'Подзадачи\'!$F$12:$F=TODAY()'),
    dateBlock('Завтра', '\'Подзадачи\'!$F$12:$F=TODAY()+1'),
    dateBlock('Послезавтра', '\'Подзадачи\'!$F$12:$F=TODAY()+2'),
    shoppingBlock
  ].join(';') + ')';
  // This range is fully generated. Clearing it removes legacy formulas that block the new spill.
  sheet.getRange(firstRow, 1, rows, 12).clearContent();
  sheet.getRange(firstRow, 2).setFormula(formula);
  var shoppingSheetId = shoppingSheet.getSheetId();
  var idFormula = '=IF(F12="Покупка";IFERROR(INDEX(FILTER(\'Список покупок\'!$A$12:$A;\'Список покупок\'!$B$12:$B=B12;\'Список покупок\'!$C$12:$C=G12);1);"");IFERROR(INDEX(FILTER(\'Подзадачи\'!$A$12:$A;\'Подзадачи\'!$B$12:$B=C12;\'Подзадачи\'!$D$12:$D=B12;\'Подзадачи\'!$E$12:$E=G12;\'Подзадачи\'!$F$12:$F=H12);1);""))';
  var actionFormula = '=IF(L12="";"";IF(F12="Покупка";HYPERLINK("#gid=' + shoppingSheetId + '&range=C"&MATCH(L12;\'Список покупок\'!$A:$A;0);"Купить");HYPERLINK("#gid=1075089967&range=E"&MATCH(L12;\'Подзадачи\'!$A:$A;0);"Закрыть")))';
  sheet.getRange(firstRow, 12).setFormula(idFormula);
  sheet.getRange(firstRow, 1).setFormula(actionFormula);
  sheet.getRange(firstRow, 1, 1, 1).copyTo(sheet.getRange(firstRow, 1, rows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  sheet.getRange(firstRow, 12, 1, 1).copyTo(sheet.getRange(firstRow, 12, rows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
}

function ensureSubtaskCheckColumn_(sheet) {
  var headers = getHeaders_(sheet);
  if (!columnNumberOrZero_(headers, 'Проверка планирования')) {
    sheet.insertColumnAfter(headers.length);
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, headers.length + 1).setValue('Проверка планирования');
  }
}

function ensureSubtaskTaskTitleColumn_(sheet) {
  var headers = getHeaders_(sheet);
  if (!columnNumberOrZero_(headers, 'Задача')) {
    var sourceIdColumn = columnNumber_(headers, 'Источник ID');
    sheet.insertColumnBefore(sourceIdColumn);
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, sourceIdColumn).setValue('Задача');
  }
}

// Проект подзадачи всегда наследуется от основной задачи, а не заполняется вручную.
function ensureSubtaskProjectColumn_(sheet) {
  ensureColumn_(sheet, 'Проект / область');
  sheet.setColumnWidth(columnNumber_(getHeaders_(sheet), 'Проект / область'), 180);
}

function ensurePlanDayProjectColumn_(planSheet, subtasksSheet, firstRow, rows) {
  ensureColumn_(planSheet, 'Проект / область');
  var planHeaders = getHeaders_(planSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var planProjectColumn = columnNumber_(planHeaders, 'Проект / область');
  var subtaskProjectColumn = columnNumber_(subtaskHeaders, 'Проект / область');
  var subtaskLastColumn = columnLetter_(subtaskHeaders.length);
  var formula = '=IF($J12="";"";IFERROR(VLOOKUP($J12;\'Подзадачи\'!$A$12:$' + subtaskLastColumn + ';' + subtaskProjectColumn + ';FALSE);""))';
  planSheet.getRange(firstRow, planProjectColumn).setFormula(formula);
  planSheet.getRange(firstRow, planProjectColumn, 1, 1).copyTo(
    planSheet.getRange(firstRow, planProjectColumn, rows, 1),
    SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
    false
  );
  planSheet.setColumnWidth(planProjectColumn, 180);
}

// Подзадача имеет один источник: обычную задачу или правило рутины.
function ensureSubtaskSourceIdColumn_(sheet) {
  var headers = getHeaders_(sheet);
  var sourceIdColumn = columnNumberOrZero_(headers, 'Источник ID');
  var legacyTaskIdColumn = columnNumberOrZero_(headers, 'Задача ID');

  if (!sourceIdColumn && legacyTaskIdColumn) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, legacyTaskIdColumn).setValue('Источник ID');
    sourceIdColumn = legacyTaskIdColumn;
  } else if (!sourceIdColumn) {
    ensureColumn_(sheet, 'Источник ID');
    sourceIdColumn = columnNumber_(getHeaders_(sheet), 'Источник ID');
  }

  headers = getHeaders_(sheet);
  var routineIdColumn = columnNumberOrZero_(headers, 'Рутина ID');
  if (!routineIdColumn) {
    return;
  }

  // Старые строки рутин получают тот же единый источник, прежде чем удалить дубликат.
  var rowCount = dataRowCount_(sheet);
  if (rowCount) {
    var sourceValues = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, sourceIdColumn, rowCount, 1).getValues();
    var routineValues = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, routineIdColumn, rowCount, 1).getValues();
    var merged = sourceValues.map(function(row, index) {
      return [row[0] || routineValues[index][0] || ''];
    });
    sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, sourceIdColumn, rowCount, 1).setValues(merged);
  }
  sheet.deleteColumn(routineIdColumn);
}

function ensureRoutineColumns_(sheet) {
  var headers = getHeaders_(sheet);
  var oldFrequency = columnNumberOrZero_(headers, 'Частота');
  if (oldFrequency) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, oldFrequency).setValue('Повторение');
  }
  ensureColumn_(sheet, 'Интервал');
  ensureColumn_(sheet, 'Оценка, мин');
  ensureColumn_(sheet, 'Текущая подзадача ID');
  ensureColumn_(sheet, 'Пауза до');
  var updatedHeaders = getHeaders_(sheet);
  // Дата паузы — ручное правило, не служебное поле автоматики.
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, columnNumber_(updatedHeaders, 'Пауза до'))
    .setBackground('#fbc965');
}

function ensureColumn_(sheet, header) {
  var headers = getHeaders_(sheet);
  if (columnNumberOrZero_(headers, header)) {
    return;
  }
  var emptyIndex = headers.indexOf('');
  if (emptyIndex !== -1) {
    sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, emptyIndex + 1).setValue(header);
    return;
  }
  sheet.insertColumnAfter(headers.length);
  sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, headers.length + 1).setValue(header);
}

function linkSubtaskToSelectedTask_(rowNumber) {
  var spreadsheet = getTrackerSpreadsheet_();
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var taskHeaders = getHeaders_(tasksSheet);
  var subtask = rowToObject_(subtaskHeaders, subtasksSheet.getRange(rowNumber, 1, 1, subtaskHeaders.length).getValues()[0]);
  var selectedTitle = String(subtask.Задача || '').trim();
  if (!selectedTitle) {
    setObjectFields_(subtasksSheet, subtaskHeaders, rowNumber, { 'Источник ID': '' });
    return;
  }
  var matches = findTaskMatches_(getExistingTasks_(tasksSheet, taskHeaders), selectedTitle);
  if (matches.length !== 1) {
    setObjectFields_(subtasksSheet, subtaskHeaders, rowNumber, { 'Источник ID': '' });
    return;
  }
  var task = matches[0];
  var fields = {
    Задача: task.Заголовок,
    'Источник ID': task.ID,
    'Проект / область': task['Проект / область'] || '',
    Обновлено: nowIso_()
  };
  if (!subtask.ID) {
    fields.ID = 'subtask_manual_' + new Date().getTime();
  }
  setObjectFields_(subtasksSheet, subtaskHeaders, rowNumber, fields);
}

function findTaskMatches_(tasks, query) {
  var normalizedQuery = normalizeTaskText_(query);
  if (normalizedQuery.length < 3) {
    return [];
  }
  return tasks.filter(function(task) {
    var normalizedTitle = normalizeTaskText_(task.Заголовок);
    return normalizedTitle === normalizedQuery || normalizedTitle.indexOf(normalizedQuery) !== -1;
  });
}

function syncSubtaskTaskTitles_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var tasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.tasksSheet);
  if (!hasDataRows_(subtasksSheet)) {
    return;
  }
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var taskDetails = {};
  getExistingTasks_(tasksSheet, getHeaders_(tasksSheet)).forEach(function(task) {
    taskDetails[String(task.ID)] = {
      title: task.Заголовок,
      project: task['Проект / область'] || ''
    };
  });
  var titleColumn = columnNumber_(subtaskHeaders, 'Задача');
  var projectColumn = columnNumber_(subtaskHeaders, 'Проект / область');
  var values = subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(subtasksSheet), subtaskHeaders.length).getValues().map(function(row) {
    var subtask = rowToObject_(subtaskHeaders, row);
    var details = taskDetails[String(subtask['Источник ID'])];
    return [details ? details.title : (subtask.Задача || ''), details ? details.project : ''];
  });
  subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, titleColumn, values.length, 1).setValues(values.map(function(row) {
    return [row[0]];
  }));
  subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, projectColumn, values.length, 1).setValues(values.map(function(row) {
    return [row[1]];
  }));
}

// Дата означает, что шаг запланирован. Явные статусы работы и блокировки не меняем.
function syncSubtaskPlanningStatuses_(startRow, rowCount) {
  var sheet = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  if (!hasDataRows_(sheet)) {
    return;
  }
  var headers = getHeaders_(sheet);
  var firstRow = Math.max(Number(startRow) || TASK_TRACKER_CONFIG_.dataStartRow, TASK_TRACKER_CONFIG_.dataStartRow);
  var availableRows = Math.max(sheet.getLastRow() - firstRow + 1, 0);
  var count = Math.min(Number(rowCount) || availableRows, availableRows);
  if (!count) {
    return;
  }
  var statusColumn = columnNumber_(headers, 'Статус');
  var dateColumn = columnNumber_(headers, 'Дата');
  var statuses = sheet.getRange(firstRow, statusColumn, count, 1).getValues();
  var dates = sheet.getRange(firstRow, dateColumn, count, 1).getValues();
  var changed = false;
  var updated = statuses.map(function(row, index) {
    var status = String(row[0] || '').trim().toLowerCase();
    var hasDate = Boolean(dates[index][0]);
    if (hasDate && status === 'new') {
      changed = true;
      return ['planned'];
    }
    if (!hasDate && status === 'planned') {
      changed = true;
      return ['new'];
    }
    return [row[0]];
  });
  if (changed) {
    sheet.getRange(firstRow, statusColumn, count, 1).setValues(updated);
  }
}

// Рутины хранят правило, а в подзадачах живут конкретные выполнения по датам.
function syncRoutineOccurrences_(refreshCurrentOccurrence) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return;
  }
  try {
  var spreadsheet = getTrackerSpreadsheet_();
  var routinesSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.routinesSheet);
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  reconcileRoutineIds_(routinesSheet);
  if (!hasDataRows_(routinesSheet)) {
    return;
  }

  var routineHeaders = getHeaders_(routinesSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var subtasksById = {};
  var routineOccurrences = {};
  if (hasDataRows_(subtasksSheet)) {
    subtasksSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(subtasksSheet), subtaskHeaders.length).getValues().forEach(function(row, index) {
      var subtask = rowToObject_(subtaskHeaders, row);
      if (subtask.ID) {
        subtasksById[String(subtask.ID)] = { item: subtask, rowNumber: index + TASK_TRACKER_CONFIG_.dataStartRow };
      }
      if (isRoutineSourceId_(subtask['Источник ID']) && subtask.Дата) {
        routineOccurrences[String(subtask['Источник ID']) + '|' + routineDateKey_(subtask.Дата)] = subtask.ID;
      }
    });
  }

  var additions = [];
  routinesSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(routinesSheet), routineHeaders.length).getValues().forEach(function(row, index) {
    var routine = rowToObject_(routineHeaders, row);
    var rowNumber = index + TASK_TRACKER_CONFIG_.dataStartRow;
    if (!routine.ID || !routine.Название) {
      return;
    }
    var currentId = String(routine['Текущая подзадача ID'] || '');
    var current = subtasksById[currentId];
    // Старые экземпляры могли появиться до колонки «Источник ID». Восстанавливаем
    // связь по уже сохраненному ID текущей подзадачи, чтобы рутина не зависала.
    if (current && String(current.item['Источник ID'] || '') !== String(routine.ID)) {
      setObjectFields_(subtasksSheet, subtaskHeaders, current.rowNumber, { 'Источник ID': routine.ID });
      current.item['Источник ID'] = routine.ID;
    }
    if (!isRoutineActive_(routine.Статус)) {
      if (current && !isClosed_(current.item.Статус)) {
        setObjectFields_(subtasksSheet, subtaskHeaders, current.rowNumber, { Статус: 'cancelled', Обновлено: nowIso_() });
      }
      if (currentId) {
        setObjectFields_(routinesSheet, routineHeaders, rowNumber, { 'Текущая подзадача ID': '' });
      }
      return;
    }
    var today = startOfDay_(new Date());
    var pausedUntil = routineDate_(routine['Пауза до']);
    if (pausedUntil && today.getTime() < pausedUntil.getTime()) {
      // Во время паузы не оставляем открытый экземпляр в плане дня.
      if (current && !isClosed_(current.item.Статус)) {
        setObjectFields_(subtasksSheet, subtaskHeaders, current.rowNumber, {
          Статус: 'cancelled',
          Обновлено: nowIso_()
        });
      }
      if (currentId) {
        setObjectFields_(routinesSheet, routineHeaders, rowNumber, { 'Текущая подзадача ID': '' });
      }
      return;
    }
    if (pausedUntil && routine['Следующая дата']) {
      // После паузы не создаем накопившиеся экземпляры задним числом: продолжаем
      // расписание с текущего дня.
      var scheduledDate = routineDate_(routine['Следующая дата']);
      if (scheduledDate && scheduledDate.getTime() < today.getTime()) {
        var resumedDate = nextRoutineDateOnOrAfterToday_(scheduledDate, routine.Повторение, routine.Интервал);
        setObjectFields_(routinesSheet, routineHeaders, rowNumber, { 'Следующая дата': resumedDate || '' });
        routine['Следующая дата'] = resumedDate || '';
      }
    }
    if (!routine['Следующая дата']) {
      return;
    }

    if (current && isClosed_(current.item.Статус)) {
      completeRoutineFromSubtask_(current.rowNumber);
      var closedDate = routineDate_(current.item.Дата);
      var todayForClosed = startOfDay_(new Date());
      if (!isDailyRoutine_(routine) || !closedDate || closedDate.getTime() >= todayForClosed.getTime()) {
        return;
      }
      // Даже если старый ежедневный экземпляр закрыли поздно, сегодня остается
      // самостоятельным днем и должен получить свою подзадачу.
      routine['Следующая дата'] = nextRoutineDateOnOrAfterToday_(closedDate, routine.Повторение, routine.Интервал) || '';
      routine['Текущая подзадача ID'] = '';
      current = null;
    }

    if (current && !isClosed_(current.item.Статус)) {
      var currentDate = routineDate_(current.item.Дата);
      if (isDailyRoutine_(routine) && currentDate && currentDate.getTime() < today.getTime()) {
        // Вчерашний ежедневный шаг не переносим на сегодня: он пропущен,
        // а для нового дня создаем самостоятельный экземпляр.
        var nextDate = nextRoutineDateOnOrAfterToday_(currentDate, routine.Повторение, routine.Интервал);
        setObjectFields_(subtasksSheet, subtaskHeaders, current.rowNumber, {
          Статус: 'skipped',
          Обновлено: nowIso_()
        });
        setObjectFields_(routinesSheet, routineHeaders, rowNumber, {
          'Следующая дата': nextDate || '',
          'Текущая подзадача ID': ''
        });
        routine['Следующая дата'] = nextDate || '';
        current = null;
      }
    }

    if (current && !isClosed_(current.item.Статус)) {
      if (refreshCurrentOccurrence) {
        setObjectFields_(subtasksSheet, subtaskHeaders, current.rowNumber, {
          Задача: routine.Название,
          Название: routine.Название,
          Дата: routine['Следующая дата'],
          'Оценка, мин': routine['Оценка, мин'] || '',
          'Источник ID': routine.ID,
          Обновлено: nowIso_()
        });
      }
      return;
    }

    var dateKey = routineDateKey_(routine['Следующая дата']);
    if (!dateKey) {
      return;
    }
    var occurrenceKey = String(routine.ID) + '|' + dateKey;
    var knownId = routineOccurrences[occurrenceKey];
    if (knownId && subtasksById[knownId] && !isClosed_(subtasksById[knownId].item.Статус)) {
      setObjectFields_(routinesSheet, routineHeaders, rowNumber, { 'Текущая подзадача ID': knownId });
      return;
    }

    var subtaskId = 'subtask_routine_' + String(routine.ID).replace(/[^a-zA-Z0-9_]/g, '_') + '_' + dateKey.replace(/-/g, '');
    additions.push(objectToRow_(subtaskHeaders, {
      ID: subtaskId,
      Задача: routine.Название,
      Название: routine.Название,
      Статус: 'new',
      Дата: routine['Следующая дата'],
      'Оценка, мин': routine['Оценка, мин'] || '',
      Обновлено: nowIso_(),
      Заметки: 'Создано автоматически из рутины.',
      'Источник ID': routine.ID
    }));
    routineOccurrences[occurrenceKey] = subtaskId;
    setObjectFields_(routinesSheet, routineHeaders, rowNumber, { 'Текущая подзадача ID': subtaskId });
  });

  if (additions.length) {
    subtasksSheet.getRange(nextDataRow_(subtasksSheet), 1, additions.length, subtaskHeaders.length).setValues(additions);
  }
  } finally {
    lock.releaseLock();
  }
}

function reconcileRoutineIds_(sheet) {
  if (!hasDataRows_(sheet)) {
    return;
  }
  var headers = getHeaders_(sheet);
  sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(sheet), headers.length).getValues().forEach(function(row, index) {
    var routine = rowToObject_(headers, row);
    if (!routine.ID && routine.Название) {
      setObjectFields_(sheet, headers, index + TASK_TRACKER_CONFIG_.dataStartRow, {
        ID: 'routine_' + new Date().getTime() + '_' + index
      });
    }
  });
}

function completeRoutineFromSubtask_(rowNumber, createNextNow) {
  var spreadsheet = getTrackerSpreadsheet_();
  var subtasksSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.subtasksSheet);
  var subtaskHeaders = getHeaders_(subtasksSheet);
  var subtask = rowToObject_(subtaskHeaders, subtasksSheet.getRange(rowNumber, 1, 1, subtaskHeaders.length).getValues()[0]);
  if (!isRoutineSourceId_(subtask['Источник ID']) || !isClosed_(subtask.Статус)) {
    return;
  }

  var routinesSheet = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.routinesSheet);
  if (!hasDataRows_(routinesSheet)) {
    return;
  }
  var routineHeaders = getHeaders_(routinesSheet);
  var routines = routinesSheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(routinesSheet), routineHeaders.length).getValues();
  var completedRoutine = false;
  routines.forEach(function(row, index) {
    var routine = rowToObject_(routineHeaders, row);
    if (String(routine.ID) !== String(subtask['Источник ID']) ||
        String(routine['Текущая подзадача ID']) !== String(subtask.ID)) {
      return;
    }
    var completedDate = routineDate_(subtask.Дата || routine['Следующая дата']);
    var today = startOfDay_(new Date());
    var nextDate = isDailyRoutine_(routine) && completedDate && completedDate.getTime() < today.getTime()
      ? nextRoutineDateOnOrAfterToday_(completedDate, routine.Повторение, routine.Интервал)
      : nextRoutineDateAfterToday_(completedDate || routine['Следующая дата'], routine.Повторение, routine.Интервал);
    setObjectFields_(routinesSheet, routineHeaders, index + TASK_TRACKER_CONFIG_.dataStartRow, {
      'Последнее выполнение': new Date(),
      'Следующая дата': nextDate || '',
      'Текущая подзадача ID': ''
    });
    completedRoutine = true;
  });
  if (completedRoutine && createNextNow) {
    syncRoutineOccurrences_(false);
  }
}

function isRoutineActive_(status) {
  var normalized = String(status || '').toLowerCase();
  return normalized === 'active';
}

function isDailyRoutine_(routine) {
  return String(routine.Повторение || '').toLowerCase() === 'ежедневно';
}

function nextRoutineDateOnOrAfterToday_(value, rule, interval) {
  var next = routineDate_(value);
  if (!next || String(rule || '').toLowerCase() === 'вручную') {
    return null;
  }
  var today = startOfDay_(new Date());
  var attempts = 0;
  do {
    next = advanceRoutineDate_(next, rule, interval);
    attempts += 1;
  } while (next && next.getTime() < today.getTime() && attempts < 1000);
  return next;
}

function nextRoutineDateAfterToday_(value, rule, interval) {
  var next = routineDate_(value);
  if (!next || String(rule || '').toLowerCase() === 'вручную') {
    return null;
  }
  var today = startOfDay_(new Date());
  var attempts = 0;
  do {
    next = advanceRoutineDate_(next, rule, interval);
    attempts += 1;
  } while (next && next.getTime() <= today.getTime() && attempts < 1000);
  return next;
}

function advanceRoutineDate_(date, rule, interval) {
  var normalized = String(rule || '').toLowerCase();
  var next = new Date(date.getTime());
  var amount = Math.max(Number(interval) || 1, 1);
  if (normalized === 'ежедневно') {
    next.setDate(next.getDate() + 1);
  } else if (normalized === 'еженедельно') {
    next.setDate(next.getDate() + 7);
  } else if (normalized === 'каждые n дней') {
    next.setDate(next.getDate() + amount);
  } else if (normalized === 'каждые n недель') {
    next.setDate(next.getDate() + amount * 7);
  } else if (normalized === 'ежемесячно') {
    next = addMonthsKeepingDay_(next, 1);
  } else if (normalized === 'ежегодно') {
    next = addMonthsKeepingDay_(next, 12);
  } else {
    return null;
  }
  return startOfDay_(next);
}

function addMonthsKeepingDay_(date, months) {
  var targetMonth = date.getMonth() + months;
  var targetYear = date.getFullYear() + Math.floor(targetMonth / 12);
  targetMonth = ((targetMonth % 12) + 12) % 12;
  var lastDay = new Date(targetYear, targetMonth + 1, 0).getDate();
  return new Date(targetYear, targetMonth, Math.min(date.getDate(), lastDay));
}

function routineDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return startOfDay_(value);
  }
  if (!value) {
    return null;
  }
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : startOfDay_(parsed);
}

function routineDateKey_(value) {
  var date = routineDate_(value);
  return date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function applyReferenceValidation_(sheet, header, referencesSheet, referenceHeader) {
  var column = columnNumber_(getHeaders_(sheet), header);
  var rows = Math.max(sheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow, 1);
  var range = sheet.getRange(TASK_TRACKER_CONFIG_.dataStartRow, column, rows, 1);
  var referenceColumn = columnNumber_(getHeaders_(referencesSheet), referenceHeader);
  var referenceRange = referencesSheet.getRange(
    TASK_TRACKER_CONFIG_.dataStartRow,
    referenceColumn,
    Math.max(referencesSheet.getMaxRows() - TASK_TRACKER_CONFIG_.headerRow, 1),
    1
  );
  var currentRule = range.getCell(1, 1).getDataValidation();
  if (currentRule &&
      currentRule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    var currentRange = currentRule.getCriteriaValues()[0];
    if (currentRange.getSheet().getSheetId() === referencesSheet.getSheetId() &&
        currentRange.getA1Notation() === referenceRange.getA1Notation()) {
      return;
    }
  }
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(referenceRange, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function deleteAutomationTriggers_() {
  var handlers = ['onIncomingTaskDraftEdit', 'runTaskTrackerAutomation', 'sendDailyIncomingReminder'];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function getTrackerSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }
  var configuredId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || TASK_TRACKER_CONFIG_.spreadsheetId;
  if (!configuredId) {
    throw new Error('Не найдена таблица. Привяжите скрипт к Google Таблице или задайте SPREADSHEET_ID в Script Properties.');
  }
  return SpreadsheetApp.openById(configuredId);
}

function hasDataRows_(sheet) {
  return dataRowCount_(sheet) > 0;
}

function dataRowCount_(sheet) {
  return Math.max(sheet.getLastRow() - TASK_TRACKER_CONFIG_.headerRow, 0);
}

function nextDataRow_(sheet) {
  return Math.max(sheet.getLastRow() + 1, TASK_TRACKER_CONFIG_.dataStartRow);
}

function getHeaders_(sheet) {
  return sheet.getRange(TASK_TRACKER_CONFIG_.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
}

function columnNumber_(headers, header) {
  var column = columnNumberOrZero_(headers, header);
  if (!column) {
    throw new Error('Не найдена колонка «' + header + '».');
  }
  return column;
}

function columnNumberOrZero_(headers, header) {
  var index = headers.indexOf(header);
  return index === -1 ? 0 : index + 1;
}

function columnLetter_(columnNumber) {
  var result = '';
  var number = Number(columnNumber);
  while (number > 0) {
    var remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function rangeIncludesColumn_(range, columnNumber) {
  return range.getColumn() <= columnNumber && range.getLastColumn() >= columnNumber;
}

function rowToObject_(headers, row) {
  var result = {};
  headers.forEach(function(header, index) {
    result[header] = row[index] === undefined ? '' : row[index];
  });
  return result;
}

function objectToRow_(headers, values) {
  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });
}

function setObjectFields_(sheet, headers, rowNumber, values) {
  Object.keys(values).forEach(function(header) {
    sheet.getRange(rowNumber, columnNumber_(headers, header)).setValue(values[header]);
  });
}

function splitTaskDrafts_(draft) {
  return String(draft).split(/\r?\n/).map(function(line) {
    return line.replace(/^\s*(?:[-*•]|\d+[.)])\s*/, '').trim();
  }).filter(function(line) {
    return Boolean(line);
  });
}

function subtaskCheck_(subtask) {
  if (isClosed_(subtask.Статус)) {
    return 'Закрыта';
  }
  if (!subtask['Источник ID']) {
    return subtask.Задача ? 'Уточнить задачу: найдено несколько или нет совпадений' : 'Не указана задача или рутина';
  }
  if (!subtask.Название) {
    return 'Не указано действие';
  }
  if (!subtask['Оценка, мин']) {
    return 'Нужна оценка времени до 240 мин';
  }
  if (!isPlanableEstimate_(subtask['Оценка, мин'])) {
    return 'Больше 4 часов: разбить на шаги';
  }
  if (!subtask.Дата) {
    return 'Нужна дата планирования';
  }
  return 'Готова к плану дня';
}

function isRoutineSourceId_(sourceId) {
  return String(sourceId || '').indexOf('routine_') === 0;
}

// Старые рутинные строки могли быть созданы до единого «Источник ID».
function isRoutineSubtask_(subtask) {
  return isRoutineSourceId_(subtask['Источник ID']) ||
    String(subtask.ID || '').indexOf('subtask_routine_') === 0;
}

function isTaskSourceId_(sourceId) {
  return String(sourceId || '').indexOf('task_') === 0;
}

function normalizeIncomingType_(value) {
  return String(value || '').trim().toLowerCase();
}

function isRoutineIncoming_(value) {
  return normalizeIncomingType_(value) === 'рутина';
}

function isShoppingIncoming_(value) {
  return normalizeIncomingType_(value) === 'покупка';
}

function isPlanableEstimate_(value) {
  var minutes = Number(value);
  return minutes > 0 && minutes <= 240;
}

function isClosed_(status) {
  return CLOSED_STATUSES_.indexOf(String(status || '').toLowerCase()) !== -1;
}

function isTerminalIncomingStatus_(status) {
  return ['done', 'cancelled', 'skipped'].indexOf(String(status || '').toLowerCase()) !== -1;
}

function getReminderDays_() {
  var settings = getTrackerSpreadsheet_().getSheetByName(TASK_TRACKER_CONFIG_.settingsSheet);
  if (!settings || !hasDataRows_(settings)) {
    return TASK_TRACKER_CONFIG_.incomingReminderDays;
  }
  var values = settings.getRange(TASK_TRACKER_CONFIG_.dataStartRow, 1, dataRowCount_(settings), 2).getValues();
  for (var index = 0; index < values.length; index += 1) {
    if (String(values[index][0]) === 'входящие_напоминание_дней') {
      return Number(values[index][1]) || TASK_TRACKER_CONFIG_.incomingReminderDays;
    }
  }
  return TASK_TRACKER_CONFIG_.incomingReminderDays;
}

function extractDeadline_(text) {
  var match = String(text || '').match(/\b(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\b/);
  if (!match) {
    return '';
  }
  var year = match[3] ? Number(match[3]) : new Date().getFullYear();
  if (year < 100) {
    year += 2000;
  }
  return [year, ('0' + Number(match[2])).slice(-2), ('0' + Number(match[1])).slice(-2)].join('-');
}

function nowIso_() {
  return new Date().toISOString();
}

function trackerUrl_() {
  var spreadsheet = getTrackerSpreadsheet_();
  var incoming = spreadsheet.getSheetByName(TASK_TRACKER_CONFIG_.incomingSheet);
  return 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/edit#gid=' + incoming.getSheetId();
}
