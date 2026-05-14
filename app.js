const STORAGE_KEYS = {
  app: "hal.architecture.v1",
  remember: "hal.auth.remember.v1",
};

const REMINDER_STOP_WORDS = new Set([
  "a",
  "an",
  "the",
  "my",
  "me",
  "to",
  "for",
  "of",
  "on",
  "at",
  "in",
  "each",
  "every",
  "day",
  "week",
  "month",
  "reminder",
]);

const INTENT_LEARNING_STOP_WORDS = new Set([
  "a",
  "an",
  "and",
  "for",
  "i",
  "im",
  "i'm",
  "it",
  "me",
  "my",
  "please",
  "the",
  "to",
  "with",
]);

const state = loadState();
state.pendingCapture = null;
state.pendingVoiceSave = false;
state.correctionContext = null;
const serverStateSync = {
  hydrated: false,
  saveTimer: null,
  lastMeta: null,
};
const authState = {
  checked: false,
  authenticated: false,
  passwordConfigured: true,
  usingBootstrapPassword: false,
};
const authDebug = {
  env: (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1") ? "local" : "hosted",
  rememberTokenPresent: false,
  lastStatus: "not-run",
  lastRestore: "not-run",
  lastLogin: "not-run",
  lastBoot: "not-started",
  lastHydration: "not-run",
  lastError: "",
};
const appBoot = {
  started: false,
};
const nativeFetch = window.fetch.bind(window);
let editingReminderId = null;
let editingTaskId = null;
let editingIdeaId = null;
let editingMeetingNoteId = null;
let editingTaskNoteId = null;
let taskNoteEditorMode = "existing";
let halSpeechVoice = null;
let microsoftSession = {
  authenticated: false,
  user: null,
};

const elements = {
  authOverlay: document.querySelector("#authOverlay"),
  authForm: document.querySelector("#authForm"),
  authPassword: document.querySelector("#authPassword"),
  authRememberDevice: document.querySelector("#authRememberDevice"),
  authStatus: document.querySelector("#authStatus"),
  authHelpText: document.querySelector("#authHelpText"),
  authDebugOutput: document.querySelector("#authDebugOutput"),
  voiceToggle: document.querySelector("#voiceToggle"),
  themeToggle: document.querySelector("#themeToggle"),
  quickLinks: document.querySelector("#quickLinks"),
  addQuickLink: document.querySelector("#addQuickLink"),
  quoteText: document.querySelector("#quoteText"),
  editQuote: document.querySelector("#editQuote"),
  openSecuritySettings: document.querySelector("#openSecuritySettings"),
  openReminders: document.querySelector("#openReminders"),
  closeReminders: document.querySelector("#closeReminders"),
  remindersDialog: document.querySelector("#remindersDialog"),
  openBackups: document.querySelector("#openBackups"),
  closeBackups: document.querySelector("#closeBackups"),
  backupDialog: document.querySelector("#backupDialog"),
  downloadBackupButton: document.querySelector("#downloadBackup"),
  saveLocalBackupButton: document.querySelector("#saveLocalBackup"),
  restoreBackupButton: document.querySelector("#restoreBackup"),
  restoreBackupInput: document.querySelector("#restoreBackupInput"),
  backupStatus: document.querySelector("#backupStatus"),
  backupSnapshotList: document.querySelector("#backupSnapshotList"),
  openReminderSettings: document.querySelector("#openReminderSettings"),
  reminderSettingsDialog: document.querySelector("#reminderSettingsDialog"),
  closeReminderSettings: document.querySelector("#closeReminderSettings"),
  reminderForm: document.querySelector("#reminderForm"),
  reminderText: document.querySelector("#reminderText"),
  reminderDateTime: document.querySelector("#reminderDateTime"),
  reminderDelivery: document.querySelector("#reminderDelivery"),
  reminderFrequency: document.querySelector("#reminderFrequency"),
  reminderRecurringPanel: document.querySelector("#reminderRecurringPanel"),
  reminderRecurringMode: document.querySelector("#reminderRecurringMode"),
  reminderRecurringInterval: document.querySelector("#reminderRecurringInterval"),
  reminderIntervalRow: document.querySelector("#reminderIntervalRow"),
  reminderMonthlyRow: document.querySelector("#reminderMonthlyRow"),
  reminderMonthlyOrdinal: document.querySelector("#reminderMonthlyOrdinal"),
  reminderMonthlyWeekday: document.querySelector("#reminderMonthlyWeekday"),
  reminderSubmit: document.querySelector("#reminderSubmit"),
  cancelReminderEdit: document.querySelector("#cancelReminderEdit"),
  reminderList: document.querySelector("#reminderList"),
  teamsNotificationsEnabled: document.querySelector("#teamsNotificationsEnabled"),
  teamsSettingsStatus: document.querySelector("#teamsSettingsStatus"),
  saveTeamsSettings: document.querySelector("#saveTeamsSettings"),
  sendTeamsTest: document.querySelector("#sendTeamsTest"),
  smsNotificationsEnabled: document.querySelector("#smsNotificationsEnabled"),
  smsPhone: document.querySelector("#smsPhone"),
  smsSettingsStatus: document.querySelector("#smsSettingsStatus"),
  saveSmsSettings: document.querySelector("#saveSmsSettings"),
  sendSmsTest: document.querySelector("#sendSmsTest"),
  securitySettingsDialog: document.querySelector("#securitySettingsDialog"),
  closeSecuritySettings: document.querySelector("#closeSecuritySettings"),
  securitySettingsForm: document.querySelector("#securitySettingsForm"),
  securityCurrentPassword: document.querySelector("#securityCurrentPassword"),
  securityNewPassword: document.querySelector("#securityNewPassword"),
  securityConfirmPassword: document.querySelector("#securityConfirmPassword"),
  securityStatus: document.querySelector("#securityStatus"),
  logoutButton: document.querySelector("#logoutButton"),
  toggleRecording: document.querySelector("#toggleRecording"),
  micStatus: document.querySelector("#micStatus"),
  micSupportNote: document.querySelector("#micSupportNote"),
  captureForm: document.querySelector("#captureForm"),
  content: document.querySelector("#content"),
  myDayDate: document.querySelector("#myDayDate"),
  openMyDaySettings: document.querySelector("#openMyDaySettings"),
  myDaySettingsDialog: document.querySelector("#myDaySettingsDialog"),
  closeMyDaySettings: document.querySelector("#closeMyDaySettings"),
  myDaySettingsForm: document.querySelector("#myDaySettingsForm"),
  uploadCalendarCsv: document.querySelector("#uploadCalendarCsv"),
  uploadCalendarCsvInput: document.querySelector("#uploadCalendarCsvInput"),
  calendarProvider: document.querySelector("#calendarProvider"),
  googleCalendarPanel: document.querySelector("#googleCalendarPanel"),
  googleCalendarStatus: document.querySelector("#googleCalendarStatus"),
  googleCalendarMode: document.querySelector("#googleCalendarMode"),
  googleWorkCalendarName: document.querySelector("#googleWorkCalendarName"),
  googlePersonalCalendarName: document.querySelector("#googlePersonalCalendarName"),
  syncOutlookCalendar: document.querySelector("#syncOutlookCalendar"),
  myDayCalendarStatus: document.querySelector("#myDayCalendarStatus"),
  addCalendarEvent: document.querySelector("#addCalendarEvent"),
  openPastDueTasks: document.querySelector("#openPastDueTasks"),
  myDayEvents: document.querySelector("#myDayEvents"),
  myDayTasks: document.querySelector("#myDayTasks"),
  meetingSearch: document.querySelector("#meetingSearch"),
  meetingList: document.querySelector("#meetingList"),
  openNewMeetingNote: document.querySelector("#openNewMeetingNote"),
  openMeetingNotesDialog: document.querySelector("#openMeetingNotesDialog"),
  meetingNotesDialog: document.querySelector("#meetingNotesDialog"),
  closeMeetingNotesDialog: document.querySelector("#closeMeetingNotesDialog"),
  meetingNotesDialogList: document.querySelector("#meetingNotesDialogList"),
  meetingNoteEditorDialog: document.querySelector("#meetingNoteEditorDialog"),
  closeMeetingNoteEditorDialog: document.querySelector("#closeMeetingNoteEditorDialog"),
  meetingNoteEditorForm: document.querySelector("#meetingNoteEditorForm"),
  meetingNoteEditorTitle: document.querySelector("#meetingNoteEditorTitle"),
  meetingNoteEditorName: document.querySelector("#meetingNoteEditorName"),
  meetingNoteEditorDate: document.querySelector("#meetingNoteEditorDate"),
  meetingNoteEditorToolbar: document.querySelector("#meetingNoteEditorToolbar"),
  meetingNoteEditorFontButton: document.querySelector("#meetingNoteEditorFontButton"),
  meetingNoteEditorSizeButton: document.querySelector("#meetingNoteEditorSizeButton"),
  meetingNoteEditorContent: document.querySelector("#meetingNoteEditorContent"),
  deleteMeetingNoteButton: document.querySelector("#deleteMeetingNoteButton"),
  ideasSearch: document.querySelector("#ideasSearch"),
  showArchivedIdeas: document.querySelector("#showArchivedIdeas"),
  ideasList: document.querySelector("#ideasList"),
  openNewIdea: document.querySelector("#openNewIdea"),
  openIdeasDialog: document.querySelector("#openIdeasDialog"),
  ideasDialog: document.querySelector("#ideasDialog"),
  openNewIdeaFromDialog: document.querySelector("#openNewIdeaFromDialog"),
  closeIdeasDialog: document.querySelector("#closeIdeasDialog"),
  ideasDialogSearch: document.querySelector("#ideasDialogSearch"),
  ideasDialogList: document.querySelector("#ideasDialogList"),
  ideaEditorDialog: document.querySelector("#ideaEditorDialog"),
  closeIdeaEditorDialog: document.querySelector("#closeIdeaEditorDialog"),
  ideaEditorForm: document.querySelector("#ideaEditorForm"),
  ideaEditorTitle: document.querySelector("#ideaEditorTitle"),
  ideaEditorName: document.querySelector("#ideaEditorName"),
  ideaEditorContent: document.querySelector("#ideaEditorContent"),
  ideaEditorToolbar: document.querySelector("#ideaEditorToolbar"),
  ideaEditorFontButton: document.querySelector("#ideaEditorFontButton"),
  ideaEditorSizeButton: document.querySelector("#ideaEditorSizeButton"),
  archiveIdeaButton: document.querySelector("#archiveIdeaButton"),
  deleteIdeaButton: document.querySelector("#deleteIdeaButton"),
  exportIdeaWord: document.querySelector("#exportIdeaWord"),
  exportIdeaPowerPoint: document.querySelector("#exportIdeaPowerPoint"),
  exportIdeaExcel: document.querySelector("#exportIdeaExcel"),
  taskListsWidget: document.querySelector("#taskListsWidget"),
  tasksDialog: document.querySelector("#tasksDialog"),
  closeTasksDialog: document.querySelector("#closeTasksDialog"),
  tasksDialogTitle: document.querySelector("#tasksDialogTitle"),
  tasksDateFilter: document.querySelector("#tasksDateFilter"),
  clearTasksDateFilter: document.querySelector("#clearTasksDateFilter"),
  tasksSearch: document.querySelector("#tasksSearch"),
  toggleCompletedTasks: document.querySelector("#toggleCompletedTasks"),
  quickTaskForm: document.querySelector("#quickTaskForm"),
  quickTaskListSelect: document.querySelector("#quickTaskListSelect"),
  quickTaskInput: document.querySelector("#quickTaskInput"),
  quickTaskDueDate: document.querySelector("#quickTaskDueDate"),
  quickTaskHasTime: document.querySelector("#quickTaskHasTime"),
  quickTaskDueTime: document.querySelector("#quickTaskDueTime"),
  openQuickTaskNoteEditor: document.querySelector("#openQuickTaskNoteEditor"),
  openTaskListCreator: document.querySelector("#openTaskListCreator"),
  newTaskListName: document.querySelector("#newTaskListName"),
  createTaskList: document.querySelector("#createTaskList"),
  openTaskComposer: document.querySelector("#openTaskComposer"),
  taskForm: document.querySelector("#taskForm"),
  taskListSelect: document.querySelector("#taskListSelect"),
  taskInput: document.querySelector("#taskInput"),
  taskDueDate: document.querySelector("#taskDueDate"),
  taskHasTime: document.querySelector("#taskHasTime"),
  taskDueTime: document.querySelector("#taskDueTime"),
  taskSubmit: document.querySelector("#taskSubmit"),
  cancelTaskEdit: document.querySelector("#cancelTaskEdit"),
  tasksList: document.querySelector("#tasksList"),
  completedTasksList: document.querySelector("#completedTasksList"),
  pastDueTasksDialog: document.querySelector("#pastDueTasksDialog"),
  closePastDueTasksDialog: document.querySelector("#closePastDueTasksDialog"),
  pastDueTasksList: document.querySelector("#pastDueTasksList"),
  taskNoteEditorDialog: document.querySelector("#taskNoteEditorDialog"),
  closeTaskNoteEditorDialog: document.querySelector("#closeTaskNoteEditorDialog"),
  taskNoteEditorForm: document.querySelector("#taskNoteEditorForm"),
  taskNoteEditorTitle: document.querySelector("#taskNoteEditorTitle"),
  taskNoteEditorToolbar: document.querySelector("#taskNoteEditorToolbar"),
  taskNoteEditorFontButton: document.querySelector("#taskNoteEditorFontButton"),
  taskNoteEditorSizeButton: document.querySelector("#taskNoteEditorSizeButton"),
  taskNoteEditorContent: document.querySelector("#taskNoteEditorContent"),
  taskListDialog: document.querySelector("#taskListDialog"),
  closeTaskListDialog: document.querySelector("#closeTaskListDialog"),
  taskListForm: document.querySelector("#taskListForm"),
  localModeGrid: document.querySelector("#localModeGrid"),
  focusGrid: document.querySelector("#focusGrid"),
  teamsStatusGrid: document.querySelector("#teamsStatusGrid"),
  entryDetail: document.querySelector("#entryDetail"),
  promptList: document.querySelector("#promptList"),
  documentPreview: document.querySelector("#documentPreview"),
  copyDocument: document.querySelector("#copyDocument"),
  downloadDocument: document.querySelector("#downloadDocument"),
  downloadMarkdown: document.querySelector("#downloadMarkdown"),
  saveFollowUp: document.querySelector("#saveFollowUp"),
  followUpList: document.querySelector("#followUpList"),
  connectMicrosoft: document.querySelector("#connectMicrosoft"),
  createMeeting: document.querySelector("#createMeeting"),
  teamsPlaceholder: document.querySelector("#teamsPlaceholder"),
  teamsActionCopy: document.querySelector("#teamsActionCopy"),
  seedDemoData: document.querySelector("#seedDemoData"),
  downloadArchive: document.querySelector("#downloadArchive"),
  clearAllData: document.querySelector("#clearAllData"),
  searchMicButtons: document.querySelectorAll(".search-mic-button"),
};

window.fetch = function halFetch(input, init = {}) {
  const requestUrl = typeof input === "string"
    ? new URL(input, window.location.origin)
    : new URL(input.url, window.location.origin);

  const isHalApiRequest = requestUrl.origin === window.location.origin && requestUrl.pathname.startsWith("/api/");
  if (!isHalApiRequest) {
    return nativeFetch(input, init);
  }

  const headers = new Headers(init.headers || (typeof input !== "string" ? input.headers : undefined) || {});
  const rememberToken = localStorage.getItem(STORAGE_KEYS.remember);
  if (rememberToken) {
    headers.set("X-HAL-Remember", rememberToken);
  }

  return nativeFetch(input, {
    ...init,
    credentials: "include",
    headers,
  });
};

initialize();

function initialize() {
  updateAuthDebugState({
    rememberTokenPresent: Boolean(localStorage.getItem(STORAGE_KEYS.remember)),
    lastBoot: "initialize() started",
  });
  applyTheme(state.theme || "dark");
  syncIconOnlyButtonTitles();
  configureEditorToolbarTabFlow();
  bindEvents();
  initializeTextToSpeech();
  syncTeamsSettingsUI();
  syncSmsSettingsUI();
  syncMyDaySettingsUI();
  syncReminderRecurrenceUI();
  syncReminderEditorState();
  syncTaskTimeField();
  syncQuickTaskTimeField();
  syncTaskEditorState();
  renderAccessState();
  void initializeAccess();
  try {
    render();
    updateAuthDebugState({
      lastBoot: "initial render complete",
      lastError: "",
    });
  } catch (error) {
    updateAuthDebugState({
      lastBoot: "initial render failed",
      lastError: error?.message || "initial render failed",
    });
    if (elements.authStatus) {
      elements.authStatus.textContent = "HAL hit a startup issue, but the login flow is still available.";
    }
  }
}

function updateAuthDebugState(patch = {}) {
  Object.assign(authDebug, patch);
  renderAuthDebug();
}

function renderAuthDebug() {
  if (!elements.authDebugOutput) {
    return;
  }

  const lines = [
    `env: ${authDebug.env}`,
    `remember token present: ${authDebug.rememberTokenPresent ? "yes" : "no"}`,
    `auth checked: ${authState.checked ? "yes" : "no"}`,
    `authenticated: ${authState.authenticated ? "yes" : "no"}`,
    `status check: ${authDebug.lastStatus}`,
    `remember restore: ${authDebug.lastRestore}`,
    `login flow: ${authDebug.lastLogin}`,
    `boot: ${authDebug.lastBoot}`,
    `hydration: ${authDebug.lastHydration}`,
    `theme: ${state.theme || "dark"}`,
  ];

  if (authDebug.lastError) {
    lines.push(`last error: ${authDebug.lastError}`);
  }

  elements.authDebugOutput.textContent = lines.join("\n");
}

function syncIconOnlyButtonTitles() {
  document.querySelectorAll('button[aria-label]').forEach((button) => {
    if (!button.getAttribute("title")) {
      button.setAttribute("title", button.getAttribute("aria-label"));
    }
  });
}

function configureEditorToolbarTabFlow() {
  [
    elements.ideaEditorToolbar,
    elements.meetingNoteEditorToolbar,
  ].forEach((toolbar) => {
    if (!toolbar) {
      return;
    }
    toolbar.querySelectorAll("button").forEach((button) => {
      button.tabIndex = -1;
    });
  });
}

function bindPasswordRevealButtons() {
  document.querySelectorAll(".password-reveal-button").forEach((button) => {
    button.addEventListener("click", () => togglePasswordReveal(button));
  });
}

async function initializeAccess() {
  updateAuthDebugState({
    rememberTokenPresent: Boolean(localStorage.getItem(STORAGE_KEYS.remember)),
    lastBoot: "initializeAccess() running",
  });
  await refreshAccessStatus();
  if (!authState.authenticated) {
    await tryRestoreRememberedAccess();
  }
  renderAccessState();

  if (authState.authenticated) {
    updateAuthDebugState({
      lastBoot: "authenticated before app boot",
    });
    await startAuthenticatedApp();
  } else {
    updateAuthDebugState({
      lastBoot: "stopped at login screen",
    });
  }
}

async function refreshAccessStatus() {
  try {
    const response = await fetch("/api/auth/status", {
      credentials: "include",
    });
    const payload = await response.json();
    authState.checked = true;
    authState.authenticated = Boolean(payload.authenticated);
    authState.passwordConfigured = payload.passwordConfigured !== false;
    authState.usingBootstrapPassword = Boolean(payload.usingBootstrapPassword);
    updateAuthDebugState({
      lastStatus: payload.authenticated ? "authenticated" : "unauthenticated",
      lastError: "",
    });
  } catch {
    authState.checked = true;
    authState.authenticated = false;
    authState.passwordConfigured = false;
    authState.usingBootstrapPassword = false;
    updateAuthDebugState({
      lastStatus: "error",
      lastError: "status check failed",
    });
  }
}

function renderAccessState() {
  document.body.classList.toggle("app-locked", !authState.authenticated);
  elements.authOverlay?.classList.toggle("hidden", authState.authenticated);
  elements.openSecuritySettings?.classList.toggle("hidden", !authState.authenticated);
  renderAuthDebug();

  if (!elements.authStatus || !elements.authHelpText) {
    return;
  }

  if (authState.authenticated) {
    elements.authStatus.textContent = "HAL is unlocked.";
    elements.authHelpText.textContent = "You are signed in to this HAL session.";
    return;
  }

  if (!authState.passwordConfigured) {
    elements.authHelpText.textContent = "HAL needs an access password configured before it can unlock.";
    elements.authStatus.textContent = "Set HAL_ACCESS_PASSWORD or SESSION_SECRET in the environment, then redeploy or restart.";
    return;
  }

  elements.authHelpText.textContent = authState.usingBootstrapPassword
    ? "Enter your current HAL password to unlock. You can change it later from Security settings."
    : "Enter your password to open HAL.";
  elements.authStatus.textContent = "HAL is locked until you sign in.";
}

function togglePasswordReveal(button) {
  const targetId = button?.dataset?.target;
  if (!targetId) {
    return;
  }

  const input = document.getElementById(targetId);
  if (!input) {
    return;
  }

  const showing = input.type === "text";
  input.type = showing ? "password" : "text";
  button.classList.toggle("is-revealed", !showing);
  const nextLabel = showing ? "Show password" : "Hide password";
  button.setAttribute("aria-label", nextLabel);
  button.setAttribute("title", nextLabel);
}

async function startAuthenticatedApp() {
  if (appBoot.started) {
    updateAuthDebugState({
      lastBoot: "startAuthenticatedApp skipped (already started)",
    });
    return;
  }

  appBoot.started = true;
  updateAuthDebugState({
    lastBoot: "startAuthenticatedApp running",
  });
  setupSpeechRecognition();
  ensureDefaultDate();
  void Promise.all([refreshSmsStatus(), refreshTeamsStatus(), refreshGoogleCalendarStatus()]);
  await initializeServerStateMirror();
  await refreshMicrosoftSession();
  updateAuthDebugState({
    lastBoot: "authenticated app boot complete",
  });
}

function bindEvents() {
  bindPasswordRevealButtons();
  on(elements.authForm, "submit", submitAuthForm);
  on(elements.voiceToggle, "click", toggleVoiceResponses);
  on(elements.themeToggle, "click", toggleTheme);
  on(elements.addQuickLink, "click", addQuickLink);
  on(elements.editQuote, "click", editQuote);
  on(elements.openSecuritySettings, "click", openSecuritySettingsDialog);
  on(elements.openReminders, "click", () => elements.remindersDialog.showModal());
  on(elements.closeReminders, "click", () => elements.remindersDialog.close());
  on(elements.openBackups, "click", openBackupDialog);
  on(elements.closeBackups, "click", () => elements.backupDialog?.close());
  on(elements.downloadBackupButton, "click", downloadBackupArchive);
  on(elements.saveLocalBackupButton, "click", createLocalBackupSnapshot);
  on(elements.restoreBackupButton, "click", promptRestoreBackup);
  on(elements.restoreBackupInput, "change", handleRestoreBackupSelection);
  on(elements.openReminderSettings, "click", () => elements.reminderSettingsDialog?.showModal());
  on(elements.closeReminderSettings, "click", () => elements.reminderSettingsDialog?.close());
  on(elements.reminderForm, "submit", addReminder);
  on(elements.reminderDateTime, "click", showReminderPicker);
  on(elements.reminderDateTime, "focus", showReminderPicker);
  on(elements.reminderFrequency, "change", syncReminderRecurrenceUI);
  on(elements.reminderRecurringMode, "change", syncReminderRecurrenceUI);
  on(elements.cancelReminderEdit, "click", resetReminderEditor);
  on(elements.saveTeamsSettings, "click", saveTeamsSettings);
  on(elements.sendTeamsTest, "click", sendTeamsTest);
  on(elements.saveSmsSettings, "click", saveSmsSettings);
  on(elements.sendSmsTest, "click", sendSmsTest);
  on(elements.closeSecuritySettings, "click", () => elements.securitySettingsDialog?.close());
  on(elements.securitySettingsForm, "submit", saveSecuritySettings);
  on(elements.logoutButton, "click", logoutHalAccess);
  on(elements.toggleRecording, "click", toggleRecording);
  on(elements.captureForm, "submit", saveCapture);
  on(elements.content, "input", previewCapture);
  on(elements.myDayDate, "change", handleMyDayDateChange);
  on(elements.myDayDate, "click", showMyDayPicker);
  on(elements.myDayDate, "focus", showMyDayPicker);
  on(elements.openMyDaySettings, "click", () => elements.myDaySettingsDialog?.showModal());
  on(elements.closeMyDaySettings, "click", () => elements.myDaySettingsDialog?.close());
  on(elements.myDaySettingsForm, "submit", saveMyDaySettings);
  on(elements.uploadCalendarCsv, "click", promptCalendarCsvUpload);
  on(elements.uploadCalendarCsvInput, "change", handleCalendarCsvUpload);
  on(elements.calendarProvider, "change", syncMyDaySettingsUI);
  on(elements.syncOutlookCalendar, "click", handleOutlookCalendarAction);
  on(elements.addCalendarEvent, "click", addCalendarEvent);
  on(elements.openPastDueTasks, "click", openPastDueTasksDialog);
  on(elements.closePastDueTasksDialog, "click", () => elements.pastDueTasksDialog.close());
  on(elements.meetingSearch, "input", renderMeetingNotes);
  on(elements.openNewMeetingNote, "click", createBlankMeetingNote);
  on(elements.openMeetingNotesDialog, "click", openMeetingNotesDialog);
  on(elements.closeMeetingNotesDialog, "click", () => elements.meetingNotesDialog.close());
  on(elements.closeMeetingNoteEditorDialog, "click", () => elements.meetingNoteEditorDialog.close());
  on(elements.meetingNoteEditorForm, "submit", saveMeetingNoteEditor);
  on(elements.meetingNoteEditorToolbar, "mousedown", preserveMeetingNoteEditorSelection);
  on(elements.meetingNoteEditorToolbar, "click", handleMeetingNoteEditorToolbarClick);
  on(elements.meetingNoteEditorFontButton, "click", handleMeetingNoteEditorFontButtonClick);
  on(elements.meetingNoteEditorSizeButton, "click", handleMeetingNoteEditorSizeButtonClick);
  on(elements.deleteMeetingNoteButton, "click", deleteMeetingNoteFromEditor);
  on(elements.ideasSearch, "input", renderIdeas);
  on(elements.showArchivedIdeas, "change", renderIdeas);
  on(elements.openNewIdea, "click", openNewIdeaEditor);
  on(elements.openIdeasDialog, "click", openIdeasDialog);
  on(elements.openNewIdeaFromDialog, "click", openNewIdeaEditor);
  on(elements.closeIdeasDialog, "click", () => elements.ideasDialog.close());
  on(elements.ideasDialogSearch, "input", renderIdeas);
  on(elements.closeIdeaEditorDialog, "click", () => elements.ideaEditorDialog.close());
  on(elements.ideaEditorForm, "submit", saveIdeaEditor);
  on(elements.ideaEditorToolbar, "mousedown", preserveIdeaEditorSelection);
  on(elements.ideaEditorToolbar, "click", handleIdeaEditorToolbarClick);
  on(elements.ideaEditorFontButton, "click", handleIdeaEditorFontButtonClick);
  on(elements.ideaEditorSizeButton, "click", handleIdeaEditorSizeButtonClick);
  on(elements.archiveIdeaButton, "click", toggleIdeaArchiveFromEditor);
  on(elements.deleteIdeaButton, "click", deleteIdeaFromEditor);
  on(elements.exportIdeaWord, "click", () => exportCurrentIdeaFromEditor("word"));
  on(elements.exportIdeaPowerPoint, "click", () => exportCurrentIdeaFromEditor("powerpoint"));
  on(elements.exportIdeaExcel, "click", () => exportCurrentIdeaFromEditor("excel"));
  on(elements.closeTasksDialog, "click", () => elements.tasksDialog.close());
  on(elements.tasksDateFilter, "change", renderTasks);
  on(elements.tasksDateFilter, "click", showTaskDatePicker);
  on(elements.tasksDateFilter, "focus", showTaskDatePicker);
  on(elements.clearTasksDateFilter, "click", clearTasksDateFilter);
  on(elements.tasksSearch, "input", renderTasks);
  on(elements.toggleCompletedTasks, "click", toggleCompletedTasksVisibility);
  on(elements.quickTaskForm, "submit", addQuickTask);
  on(elements.quickTaskDueDate, "click", showQuickTaskDueDatePicker);
  on(elements.quickTaskDueDate, "focus", showQuickTaskDueDatePicker);
  on(elements.quickTaskDueTime, "click", showQuickTaskTimePicker);
  on(elements.quickTaskDueTime, "focus", showQuickTaskTimePicker);
  on(elements.quickTaskHasTime, "change", syncQuickTaskTimeField);
  on(elements.openQuickTaskNoteEditor, "click", openDraftTaskNoteEditor);
  on(elements.openTaskListCreator, "click", () => elements.taskListDialog?.showModal());
  on(elements.closeTaskListDialog, "click", () => elements.taskListDialog?.close());
  on(elements.taskListForm, "submit", createTaskList);
  on(elements.openTaskComposer, "click", openTaskComposer);
  on(elements.taskForm, "submit", addManualTask);
  on(elements.taskDueDate, "click", showTaskDueDatePicker);
  on(elements.taskDueDate, "focus", showTaskDueDatePicker);
  on(elements.taskDueTime, "click", showTaskTimePicker);
  on(elements.taskDueTime, "focus", showTaskTimePicker);
  on(elements.taskHasTime, "change", syncTaskTimeField);
  on(elements.cancelTaskEdit, "click", resetTaskEditor);
  on(elements.closeTaskNoteEditorDialog, "click", closeTaskNoteEditor);
  on(elements.taskNoteEditorForm, "submit", saveTaskNoteEditor);
  on(elements.taskNoteEditorToolbar, "mousedown", preserveTaskNoteEditorSelection);
  on(elements.taskNoteEditorToolbar, "click", handleTaskNoteEditorToolbarClick);
  on(elements.taskNoteEditorFontButton, "click", handleTaskNoteEditorFontButtonClick);
  on(elements.taskNoteEditorSizeButton, "click", handleTaskNoteEditorSizeButtonClick);
  on(elements.copyDocument, "click", copyDraft);
  on(elements.downloadDocument, "click", downloadDraft);
  on(elements.downloadMarkdown, "click", downloadMarkdown);
  on(elements.saveFollowUp, "click", saveSelectedToFollowUp);
  on(elements.connectMicrosoft, "click", () => window.alert("Microsoft integration will plug into this architecture later."));
  on(elements.createMeeting, "click", () => window.alert("Meeting creation will use the selected note or my day event once calendar integration is connected."));
  on(elements.teamsPlaceholder, "click", () => window.alert("Teams, Entra, reminders, and text delivery are intentionally deferred while we finalize HAL's architecture."));
  on(elements.seedDemoData, "click", seedDemoData);
  on(elements.downloadArchive, "click", downloadArchive);
  on(elements.clearAllData, "click", clearAllData);
  elements.searchMicButtons.forEach((button) => {
    button.addEventListener("click", () => startFieldDictation(button.dataset.target));
  });
}

function render() {
  syncVoiceToggle();
  syncThemeToggleLabel();
  renderQuote();
  renderQuickLinks();
  renderBackupSummary();
  renderLocalMode();
  renderFocus();
  renderTeamsStatus();
  renderReminders();
  renderMyDay();
  renderMeetingNotes();
  renderIdeas();
  renderTasks();
  renderSelectedWorkspace();
  renderFollowUps();
}

function toggleVoiceResponses() {
  state.voiceResponsesEnabled = !state.voiceResponsesEnabled;
  persist();
  syncVoiceToggle();
  if (state.voiceResponsesEnabled) {
    speakHal("Voice responses on");
  } else if (window.speechSynthesis) {
    window.speechSynthesis.cancel();
  }
}

function syncVoiceToggle() {
  if (!elements.voiceToggle) {
    return;
  }
  const enabled = Boolean(state.voiceResponsesEnabled);
  const label = enabled ? "Turn spoken responses off" : "Turn spoken responses on";
  elements.voiceToggle.setAttribute("aria-label", label);
  elements.voiceToggle.setAttribute("title", label);
  elements.voiceToggle.classList.toggle("is-muted", !enabled);
}

function toggleTheme() {
  state.theme = state.theme === "light" ? "dark" : "light";
  applyTheme(state.theme);
  persist();
  syncThemeToggleLabel();
}

function applyTheme(theme) {
  document.body.dataset.theme = theme;
}

function syncThemeToggleLabel() {
  if (!elements.themeToggle) {
    return;
  }
  const nextTheme = state.theme === "light" ? "dark" : "light";
  const label = nextTheme === "light" ? "Switch to light mode" : "Switch to dark mode";
  elements.themeToggle.setAttribute("aria-label", label);
  elements.themeToggle.setAttribute("title", label);
}

function renderQuote() {
  elements.quoteText.textContent = state.quote;
}

function renderQuickLinks() {
  elements.quickLinks.innerHTML = "";
  state.quickLinks.forEach((link) => {
    const button = document.createElement("button");
    button.className = "quick-link-button";
    button.type = "button";
    button.title = `${link.label}\n${link.url}`;
    button.appendChild(buildQuickLinkIcon(link));
    button.addEventListener("click", () => window.open(link.url, "_blank", "noopener"));
    button.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      const confirmed = window.confirm(`Remove quick link "${link.label}"?`);
      if (!confirmed) {
        return;
      }
      state.quickLinks = state.quickLinks.filter((item) => item.id !== link.id);
      persist();
      renderQuickLinks();
    });
    elements.quickLinks.appendChild(button);
  });
}

function buildQuickLinkIcon(link) {
  const icon = document.createElement("span");
  icon.className = "quick-link-icon";
  icon.setAttribute("aria-hidden", "true");

  const knownGlyph = getKnownQuickLinkGlyph(link);
  if (knownGlyph) {
    icon.innerHTML = knownGlyph;
    return icon;
  }

  const fallback = document.createElement("span");
  fallback.className = "quick-link-fallback hidden";
  fallback.textContent = String(link.label || "")
    .trim()
    .slice(0, 2)
    .toUpperCase() || "?";

  const image = document.createElement("img");
  image.className = "quick-link-image";
  image.alt = "";
  image.loading = "lazy";
  image.referrerPolicy = "no-referrer";
  image.src = getQuickLinkIconUrl(link.url);
  image.addEventListener("error", () => {
    image.remove();
    fallback.classList.remove("hidden");
  });
  image.addEventListener("load", () => {
    fallback.classList.add("hidden");
  });

  icon.appendChild(image);
  icon.appendChild(fallback);
  return icon;
}

function getQuickLinkIconUrl(rawUrl) {
  try {
    const url = new URL(rawUrl);
    return `https://www.google.com/s2/favicons?sz=64&domain_url=${encodeURIComponent(url.origin)}`;
  } catch {
    return "";
  }
}

function getKnownQuickLinkGlyph(link) {
  const haystack = `${link.label} ${link.url}`.toLowerCase();

  if (haystack.includes("excel")) {
    return '<span class="quick-link-glyph excel">X</span>';
  }
  if (haystack.includes("powerpoint")) {
    return '<span class="quick-link-glyph powerpoint">P</span>';
  }
  if (haystack.includes("outlook")) {
    return '<span class="quick-link-glyph outlook">O</span>';
  }
  if (haystack.includes("power bi") || haystack.includes("powerbi")) {
    return '<span class="quick-link-glyph powerbi">BI</span>';
  }
  if (haystack.includes("d365") || haystack.includes("dynamics 365") || haystack.includes("dynamics")) {
    return '<span class="quick-link-glyph d365">D</span>';
  }

  return "";
}

function renderLocalMode() {
  renderStats(elements.localModeGrid, [
    { value: "Architecture", label: "Current focus", detail: "Layout and data model first" },
    { value: String(state.taskLists.length), label: "Task lists", detail: "Default list plus custom lists" },
    { value: String(state.reminders.length), label: "Reminders", detail: "Stored locally for now" },
  ]);
}

function renderFocus() {
  const selectedDate = elements.myDayDate.value;
  renderStats(elements.focusGrid, [
    { value: String(getTasksForDate(selectedDate).length), label: "Tasks for day", detail: formatDisplayDate(selectedDate) },
    { value: String(state.ideas.filter((item) => !item.archived).length), label: "Active ideas", detail: "Archive hides them from normal search" },
    { value: String(state.meetingNotes.filter((note) => note.callDate === selectedDate).length), label: "Meeting notes", detail: "Attached to selected day" },
  ]);
}

function renderTeamsStatus() {
  renderStats(elements.teamsStatusGrid, [
    { value: microsoftSession.authenticated ? "Ready" : "Local", label: "Calendar sync", detail: microsoftSession.authenticated ? "Outlook sync available" : "Local events for now" },
    { value: microsoftSession.authenticated ? "Connected" : "Optional", label: "Microsoft sign-in", detail: microsoftSession.authenticated ? (microsoftSession.user?.name || "Signed in") : "Connect when you want Outlook data" },
    { value: "Active", label: "Reminders", detail: "Teams and SMS paths are scaffolded" },
  ]);
}

function renderStats(container, cards) {
  if (!container) {
    return;
  }
  container.innerHTML = "";
  cards.forEach((card) => {
    const item = document.createElement("div");
    item.className = "item-card";
    item.innerHTML = `<div class="card-title-row"><strong>${card.value}</strong></div><div class="card-meta"><span>${card.label}</span></div><p class="card-body">${card.detail}</p>`;
    container.appendChild(item);
  });
}

function renderReminders() {
  elements.reminderList.innerHTML = "";
  if (!state.reminders.length) {
    elements.reminderList.innerHTML = '<div class="item-card"><p class="card-body">No reminders yet.</p></div>';
    return;
  }

  state.reminders.forEach((reminder) => {
    const card = document.createElement("article");
    card.className = "item-card";
    card.innerHTML = `
      <div class="card-title-row">
        <h3 class="card-title">${escapeHtml(reminder.text)}</h3>
        <div class="card-actions">
          <button class="ghost-button" type="button" data-action="edit">Edit</button>
          <button class="ghost-button" type="button" data-action="delete">Delete</button>
        </div>
      </div>
      <div class="card-meta">
        <span class="status-pill">${getReminderScheduleLabel(reminder)}</span>
        <span>${formatDateTime(reminder.at)}</span>
        <span>${getReminderDeliveryLabel(reminder)}</span>
      </div>
    `;
    card.querySelector('[data-action="edit"]').addEventListener("click", () => {
      startReminderEdit(reminder.id);
    });
    card.querySelector('[data-action="delete"]').addEventListener("click", () => {
      state.reminders = state.reminders.filter((item) => item.id !== reminder.id);
      if (editingReminderId === reminder.id) {
        resetReminderEditor();
      }
      persist();
      renderReminders();
    });
    elements.reminderList.appendChild(card);
  });
}

function showReminderPicker() {
  if (typeof elements.reminderDateTime?.showPicker === "function") {
    elements.reminderDateTime.showPicker();
  }
}

function syncReminderRecurrenceUI() {
  const isRecurring = elements.reminderFrequency?.value === "recurring";
  elements.reminderRecurringPanel?.classList.toggle("hidden", !isRecurring);

  const usesMonthlyWeekday = elements.reminderRecurringMode?.value === "monthlyWeekday";
  elements.reminderIntervalRow?.classList.toggle("hidden", !isRecurring || usesMonthlyWeekday);
  elements.reminderMonthlyRow?.classList.toggle("hidden", !isRecurring || !usesMonthlyWeekday);
}

function startReminderEdit(reminderId) {
  const reminder = state.reminders.find((item) => item.id === reminderId);
  if (!reminder) {
    return;
  }

  editingReminderId = reminder.id;
  elements.reminderText.value = reminder.text;
  elements.reminderDateTime.value = reminder.at;
  elements.reminderDelivery.value = reminder.delivery || "inApp";
  elements.reminderFrequency.value = reminder.frequency || "once";

  if (reminder.recurrence?.mode === "monthlyWeekday") {
    elements.reminderRecurringMode.value = "monthlyWeekday";
    elements.reminderMonthlyOrdinal.value = reminder.recurrence.ordinal || "first";
    elements.reminderMonthlyWeekday.value = reminder.recurrence.weekday || "monday";
  } else if (reminder.recurrence?.mode === "weeklyWeekday") {
    elements.reminderRecurringMode.value = "interval";
    elements.reminderRecurringInterval.value = "weekly";
  } else {
    elements.reminderRecurringMode.value = "interval";
    elements.reminderRecurringInterval.value = reminder.recurrence?.interval || "daily";
  }

  syncReminderRecurrenceUI();
  syncReminderEditorState();
  elements.reminderText.focus();
}

function syncReminderEditorState() {
  const isEditing = Boolean(editingReminderId);
  if (elements.reminderSubmit) {
    elements.reminderSubmit.textContent = isEditing ? "Save reminder" : "Add reminder";
  }
  elements.cancelReminderEdit?.classList.toggle("hidden", !isEditing);
}

function resetReminderEditor() {
  editingReminderId = null;
  elements.reminderForm?.reset();
  if (elements.reminderDelivery) {
    elements.reminderDelivery.value = "inApp";
  }
  syncReminderRecurrenceUI();
  syncReminderEditorState();
}

function renderMyDay() {
  const selectedDate = elements.myDayDate.value;
  const events = state.calendarEvents
    .filter((event) => event.date === selectedDate)
    .sort((left, right) => compareCalendarEvents(left, right));
  const tasks = getTasksForDate(selectedDate)
    .filter((task) => !task.done)
    .sort((left, right) => Number(Boolean(right.highlighted)) - Number(Boolean(left.highlighted)));
  const pastDueTasks = getPastDueTasks();

  elements.myDayEvents.innerHTML = "";
  elements.myDayTasks.innerHTML = "";

  if (!events.length) {
    const provider = state.calendarSettings?.provider || "local";
    const emptyMessage = provider === "csv"
      ? "No imported calendar events for this date."
      : "Calendar events will show here. You can add manual events or sync Outlook for this day.";
    elements.myDayEvents.innerHTML = `<div class="item-card"><p class="card-body">${emptyMessage}</p></div>`;
  } else {
    events.forEach((event) => {
      const card = document.createElement("article");
      card.className = "item-card";
      card.innerHTML = `
        <div class="card-title-row">
          <h3 class="card-title">${escapeHtml(event.title)}</h3>
          ${event.source === "outlook"
            ? (event.webLink
                ? '<a class="ghost-button compact-ghost-button" target="_blank" rel="noreferrer">Open</a>'
                : '<span class="status-pill">Outlook</span>')
            : event.source === "csv"
              ? ''
            : '<button class="ghost-button" type="button">Delete</button>'}
        </div>
        <div class="card-meta"><span>${escapeHtml(event.time || "All day")}</span>${event.source === "csv" ? "" : `<span>${event.source === "outlook" ? "Outlook" : "Manual"}</span>`}</div>
      `;
      if (event.source === "outlook") {
        const link = card.querySelector("a");
        if (link && event.webLink) {
          link.href = event.webLink;
        }
      } else if (event.source === "csv") {
        // Imported CSV events are read-only inside HAL.
      } else {
        card.querySelector("button").addEventListener("click", () => {
          state.calendarEvents = state.calendarEvents.filter((item) => item.id !== event.id);
          persist();
          render();
        });
      }
      elements.myDayEvents.appendChild(card);
    });
  }

  if (!tasks.length) {
    elements.myDayTasks.innerHTML = '<div class="item-card"><p class="card-body">No tasks assigned to this date.</p></div>';
  } else {
    tasks.forEach((task) => {
      const row = document.createElement("article");
      row.className = `item-card${task.highlighted ? " highlighted-task" : ""}`;
      row.innerHTML = `
        <div class="card-title-row">
          <div class="task-row">
            <input type="checkbox" ${task.done ? "checked" : ""}>
            <span>${escapeHtml(task.title)}${task.dueTime ? ` <small>${escapeHtml(formatTimeValue(task.dueTime))}</small>` : ""}</span>
          </div>
          <div class="card-actions">
            <button class="icon-button${task.highlighted ? " is-active" : ""}" type="button" data-action="highlight-task" aria-label="${task.highlighted ? "Unhighlight task" : "Highlight task"}" title="${task.highlighted ? "Unhighlight task" : "Highlight task"}">
              <span class="star-icon" aria-hidden="true"></span>
            </button>
            <button class="icon-button" type="button" data-action="reschedule-task" aria-label="Reschedule task" title="Reschedule task">
              <span class="calendar-icon" aria-hidden="true"></span>
            </button>
            <input class="inline-date-picker hidden" type="date" data-action="reschedule-input">
          </div>
        </div>
      `;
      const checkbox = row.querySelector('input[type="checkbox"]');
      const highlightButton = row.querySelector('[data-action="highlight-task"]');
      const rescheduleButton = row.querySelector('[data-action="reschedule-task"]');
      const rescheduleInput = row.querySelector('[data-action="reschedule-input"]');
      checkbox.addEventListener("mousedown", (event) => event.stopPropagation());
      checkbox.addEventListener("click", (event) => event.stopPropagation());
      checkbox.addEventListener("change", () => {
        setTaskCompletion(task, checkbox.checked);
      });
      highlightButton?.addEventListener("click", (event) => {
        event.stopPropagation();
        toggleTaskHighlight(task);
      });
      if (rescheduleButton && rescheduleInput) {
        rescheduleInput.value = task.dueDate || "";
        rescheduleButton.addEventListener("click", (event) => {
          event.stopPropagation();
          rescheduleInput.classList.toggle("hidden");
          if (!rescheduleInput.classList.contains("hidden")) {
            window.requestAnimationFrame(() => {
              rescheduleInput.focus();
              if (typeof rescheduleInput.showPicker === "function") {
                rescheduleInput.showPicker();
              }
            });
          }
        });
        rescheduleInput.addEventListener("mousedown", (event) => event.stopPropagation());
        rescheduleInput.addEventListener("click", (event) => event.stopPropagation());
        rescheduleInput.addEventListener("blur", () => {
          window.setTimeout(() => {
            rescheduleInput.classList.add("hidden");
          }, 180);
        });
        rescheduleInput.addEventListener("change", () => {
          if (!rescheduleInput.value) {
            return;
          }
          task.dueDate = rescheduleInput.value;
          persist();
          rescheduleInput.classList.add("hidden");
          render();
        });
      }
      elements.myDayTasks.appendChild(row);
    });
  }

  if (elements.openPastDueTasks) {
    elements.openPastDueTasks.textContent = pastDueTasks.length ? `Past due tasks (${pastDueTasks.length})` : "Past due tasks";
  }
  syncMyDayCalendarUI();
}

function renderMeetingNotes() {
  const query = elements.meetingSearch.value.trim().toLowerCase();
  const selectedDate = elements.myDayDate.value;
  const notes = state.meetingNotes.filter((note) => {
    const haystack = `${note.title} ${note.content} ${note.person}`.toLowerCase();
    return haystack.includes(query);
  });

  elements.meetingList.innerHTML = "";
  if (!notes.length) {
    elements.meetingList.innerHTML = '<div class="item-card compact-preview-card"><p class="compact-preview-line">No meeting notes yet.</p></div>';
    if (elements.meetingNotesDialogList) {
      elements.meetingNotesDialogList.innerHTML = '<div class="item-card"><p class="card-body">No meeting notes yet. Say something like "I need notes for my call with Sarah tomorrow".</p></div>';
    }
    return;
  }

  const previewNote = notes[0];
  const previewText = stripHtml(previewNote.content);
  const previewLine = `${previewNote.title}: ${truncate(previewText, 72)}`;
  const previewCard = document.createElement("article");
  previewCard.className = "item-card compact-preview-card";
  previewCard.innerHTML = `<p class="compact-preview-line" title="${escapeHtml(`${previewNote.title}: ${previewText}`)}">${escapeHtml(previewLine)}</p>`;
  previewCard.addEventListener("click", () => openMeetingNotesDialog());
  elements.meetingList.appendChild(previewCard);

  if (elements.meetingNotesDialogList) {
    elements.meetingNotesDialogList.innerHTML = "";
    notes.forEach((note) => {
      const card = document.createElement("article");
      card.className = "item-card";
      if (note.callDate && note.callDate < selectedDate) {
        card.classList.add("archived");
      }
      card.innerHTML = `
        <div class="card-title-row">
          <h3 class="card-title">${escapeHtml(note.title)}</h3>
          <div class="card-actions">
            <button class="ghost-button" type="button">Delete</button>
          </div>
        </div>
        <div class="card-meta">
          <span class="status-pill">${note.callDate ? formatDisplayDate(note.callDate) : "No date"}</span>
        </div>
        <div class="card-body formatted-note-body">${formatStructuredNoteForDisplay(note.content)}</div>
      `;
      const deleteButton = card.querySelector("button");
      deleteButton.addEventListener("click", () => {
        state.meetingNotes = state.meetingNotes.filter((item) => item.id !== note.id);
        persist();
        render();
      });
      card.addEventListener("click", () => editMeetingNote(note.id));
      elements.meetingNotesDialogList.appendChild(card);
    });
  }
}

function openMeetingNotesDialog() {
  renderMeetingNotes();
  if (!elements.meetingNotesDialog?.open) {
    elements.meetingNotesDialog.showModal();
  }
}

function createBlankMeetingNote() {
  editingMeetingNoteId = null;
  elements.meetingNoteEditorTitle.textContent = "New Meeting Note";
  elements.meetingNoteEditorName.value = "";
  elements.meetingNoteEditorDate.value = elements.myDayDate?.value || "";
  elements.meetingNoteEditorContent.innerHTML = "Talking Points<br>- ";
  elements.deleteMeetingNoteButton.classList.add("hidden");
  if (!elements.meetingNoteEditorDialog?.open) {
    elements.meetingNoteEditorDialog.showModal();
  }
  elements.meetingNoteEditorName.focus();
}

function renderIdeas() {
  const widgetQuery = elements.ideasSearch?.value.trim().toLowerCase() || "";
  const dialogQuery = elements.ideasDialogSearch?.value.trim().toLowerCase() || "";
  const query = dialogQuery || widgetQuery;
  const showArchived = elements.showArchivedIdeas.checked;
  const ideas = state.ideas.filter((idea) => {
    if (!showArchived && idea.archived) {
      return false;
    }
    const haystack = `${idea.title} ${idea.content}`.toLowerCase();
    return haystack.includes(query);
  });
  const visibleIdeas = ideas.slice(0, 10);

  elements.ideasList.innerHTML = "";
  if (!ideas.length) {
    elements.ideasList.innerHTML = '<div class="item-card"><p class="card-body">No ideas match this view yet.</p></div>';
    if (elements.ideasDialogList) {
      elements.ideasDialogList.innerHTML = '<div class="item-card"><p class="card-body">No ideas match this view yet.</p></div>';
    }
    return;
  }

  visibleIdeas.forEach((idea) => {
    elements.ideasList.appendChild(buildIdeaCard(idea, true));
  });

  if (ideas.length > 10) {
    const overflow = document.createElement("article");
    overflow.className = "item-card";
    overflow.innerHTML = `<p class="card-body">Showing the latest 10 of ${ideas.length} ideas. Use search or open the full ideas window for the rest.</p>`;
    elements.ideasList.appendChild(overflow);
  }

  if (elements.ideasDialogList) {
    elements.ideasDialogList.innerHTML = "";
    ideas.forEach((idea) => {
      elements.ideasDialogList.appendChild(buildIdeaCard(idea, false));
    });
  }
}

function buildIdeaCard(idea, compact = false) {
  const card = document.createElement("article");
  card.className = `item-card idea-card${compact ? " compact-idea-card" : ""}${idea.archived ? " archived" : ""}`;
  card.innerHTML = `
    <div class="card-title-row">
      <h3 class="card-title" title="${escapeHtml(idea.title)}">${escapeHtml(idea.title)}</h3>
      <div class="card-actions">
        <button class="icon-button" type="button" aria-label="${idea.archived ? "Restore idea" : "Archive idea"}" title="${idea.archived ? "Restore idea" : "Archive idea"}">
          <span class="folder-icon" aria-hidden="true"></span>
        </button>
        <button class="icon-button" type="button" aria-label="Delete idea" title="Delete idea">
          <span class="x-icon" aria-hidden="true"></span>
        </button>
      </div>
    </div>
    ${compact ? "" : `<div class="card-body formatted-note-body formatted-idea-preview">${formatFormattedIdeaPreview(idea.content)}</div>`}
  `;

  const [archiveButton, deleteButton] = card.querySelectorAll("button");
  archiveButton.addEventListener("click", (event) => {
    event.stopPropagation();
    idea.archived = !idea.archived;
    persist();
    render();
  });
  deleteButton.addEventListener("click", (event) => {
    event.stopPropagation();
    const confirmed = window.confirm(`Delete "${idea.title}"?`);
    if (!confirmed) {
      return;
    }
    state.ideas = state.ideas.filter((item) => item.id !== idea.id);
    if (editingIdeaId === idea.id) {
      elements.ideaEditorDialog?.close();
      editingIdeaId = null;
    }
    persist();
    render();
  });
  card.addEventListener("click", () => openIdeaEditor(idea.id));
  return card;
}

function renderTasks() {
  renderQuickTaskComposer();
  renderTaskListsWidget();
  renderTasksDialog();
}

function renderQuickTaskComposer() {
  if (!elements.quickTaskListSelect) {
    return;
  }
  const selectedId = elements.quickTaskListSelect.value || state.taskLists[0]?.id || "";
  elements.quickTaskListSelect.innerHTML = state.taskLists
    .map((list) => `<option value="${escapeHtml(list.id)}">${escapeHtml(list.name)}</option>`)
    .join("");
  elements.quickTaskListSelect.value = selectedId || state.taskLists[0]?.id || "";
  syncQuickTaskTimeField();
  syncQuickTaskNoteButton();
}

function renderTaskListsWidget() {
  if (!elements.taskListsWidget) {
    return;
  }

  elements.taskListsWidget.innerHTML = "";
  state.taskLists.forEach((list) => {
    const openCount = state.tasks.filter((task) => task.listId === list.id && !task.done).length;
    const row = document.createElement("article");
    row.className = "item-card task-list-row";
    row.innerHTML = `
      <button class="icon-button" type="button" data-action="edit-list" aria-label="Edit task list" title="Edit task list">
        <span class="pencil-icon" aria-hidden="true"></span>
      </button>
      <button class="ghost-button task-list-button" type="button" data-action="open-list">
        <span>${escapeHtml(list.name)}</span>
        <span class="task-list-meta">${openCount} open</span>
      </button>
      <button class="icon-button" type="button" data-action="delete-list" aria-label="Delete task list" title="Delete task list">
        <span class="x-icon" aria-hidden="true"></span>
      </button>
    `;

    row.querySelector('[data-action="open-list"]').addEventListener("click", () => openTaskList(list.id));
    row.querySelector('[data-action="edit-list"]').addEventListener("click", () => editTaskList(list.id));
    row.querySelector('[data-action="delete-list"]').addEventListener("click", () => deleteTaskList(list.id));
    elements.taskListsWidget.appendChild(row);
  });
}

function renderTasksDialog() {
  const listId = state.selectedTaskListId || state.taskLists[0]?.id;
  const list = state.taskLists.find((entry) => entry.id === listId);
  if (!list) {
    return;
  }

  if (elements.tasksDialogTitle) {
    elements.tasksDialogTitle.textContent = list.name;
  }

  renderTaskListSelect();

  const query = elements.tasksSearch?.value.trim().toLowerCase() || "";
  const filterDate = elements.tasksDateFilter?.value || "";
  const showCompleted = Boolean(state.showCompletedTasks);

  if (elements.toggleCompletedTasks) {
    elements.toggleCompletedTasks.textContent = showCompleted ? "Hide completed" : "Show Completed";
  }
  syncTasksDateFilterButton();

  const activeTasks = state.tasks.filter((task) => {
    if (task.listId !== listId || task.done) {
      return false;
    }
    if (filterDate && task.dueDate !== filterDate) {
      return false;
    }
    return task.title.toLowerCase().includes(query);
  });

  const completedTasks = state.tasks.filter((task) => {
    if (task.listId !== listId || !task.done) {
      return false;
    }
    if (filterDate && task.dueDate !== filterDate) {
      return false;
    }
    return task.title.toLowerCase().includes(query);
  });

  elements.tasksList.innerHTML = "";
  elements.completedTasksList.innerHTML = "";
  elements.completedTasksList.classList.toggle("hidden", !showCompleted);

  if (!activeTasks.length) {
    elements.tasksList.innerHTML = '<div class="item-card"><p class="card-body">No tasks in this list yet.</p></div>';
  } else {
    activeTasks.forEach((task) => elements.tasksList.appendChild(buildTaskCard(task)));
  }

  if (showCompleted) {
    if (!completedTasks.length) {
      elements.completedTasksList.innerHTML = '<div class="item-card"><p class="card-body">No completed tasks in this list.</p></div>';
    } else {
      completedTasks.forEach((task) => elements.completedTasksList.appendChild(buildTaskCard(task, true)));
    }
  }

  syncTaskEditorState();
}

function buildTaskCard(task, archived = false, options = {}) {
  const card = document.createElement("article");
  card.className = `item-card${archived ? " archived" : ""}${task.highlighted ? " highlighted-task" : ""}`;
  const quickReschedule = Boolean(options.quickReschedule);
  const hasNotes = Boolean(stripHtml(task.notes || ""));
  card.innerHTML = `
    <div class="card-title-row">
      <div class="task-row">
        <input type="checkbox" ${task.done ? "checked" : ""}>
        <span>${escapeHtml(task.title)}</span>
      </div>
      <div class="card-actions">
        <button class="icon-button${task.highlighted ? " is-active" : ""}" type="button" aria-label="${task.highlighted ? "Unhighlight task" : "Highlight task"}" title="${task.highlighted ? "Unhighlight task" : "Highlight task"}">
          <span class="star-icon" aria-hidden="true"></span>
        </button>
        <button class="ghost-button compact-ghost-button task-note-button${hasNotes ? " has-notes" : ""}" type="button" aria-label="${hasNotes ? "Edit task note" : "Add task note"}" title="${hasNotes ? "Edit task note" : "Add task note"}">
          Note
        </button>
        ${quickReschedule ? `
          <button class="icon-button" type="button" data-action="reschedule-task" aria-label="Reschedule task" title="Reschedule task">
            <span class="calendar-icon" aria-hidden="true"></span>
          </button>
          <input class="inline-date-picker hidden" type="date" data-action="reschedule-input">
        ` : ""}
        <button class="icon-button" type="button" aria-label="Delete task" title="Delete task">
          <span class="x-icon" aria-hidden="true"></span>
        </button>
      </div>
    </div>
    <div class="card-meta task-meta-line">
      ${task.dueDate
        ? `<span class="status-pill">${formatTaskDue(task)}</span>`
        : `
          <span class="inline-date-anchor">
            <button class="ghost-button compact-icon-button undated-task-chip" type="button" data-action="assign-task-date" aria-label="Assign task date" title="Assign task date">
              <span class="calendar-icon" aria-hidden="true"></span>
            </button>
            <input class="inline-date-picker hidden" type="date" data-action="assign-task-date-input">
          </span>
        `}
      <span>${getTaskListName(task.listId)}</span>
    </div>
  `;
  const checkbox = card.querySelector('input');
  checkbox.addEventListener("mousedown", (event) => event.stopPropagation());
  checkbox.addEventListener("click", (event) => event.stopPropagation());
  checkbox.addEventListener("change", (event) => {
    event.stopPropagation();
    setTaskCompletion(task, checkbox.checked);
  });
  const highlightButton = card.querySelector('[aria-label="Highlight task"], [aria-label="Unhighlight task"]');
  const noteButton = card.querySelector(".task-note-button");
  const rescheduleButton = card.querySelector('[data-action="reschedule-task"]');
  const rescheduleInput = card.querySelector('[data-action="reschedule-input"]');
  const assignDateButton = card.querySelector('[data-action="assign-task-date"]');
  const assignDateInput = card.querySelector('[data-action="assign-task-date-input"]');
  const deleteButton = card.querySelector('[aria-label="Delete task"]');
  highlightButton.addEventListener("click", (event) => {
    event.stopPropagation();
    toggleTaskHighlight(task);
  });
  noteButton.addEventListener("click", (event) => {
    event.stopPropagation();
    openTaskNoteEditor(task.id);
  });
  if (rescheduleButton && rescheduleInput) {
    rescheduleInput.value = task.dueDate || "";
    rescheduleButton.addEventListener("click", (event) => {
      event.stopPropagation();
      if (typeof rescheduleInput.showPicker === "function") {
        rescheduleInput.showPicker();
      } else {
        rescheduleInput.focus();
        rescheduleInput.click();
      }
    });
    rescheduleInput.addEventListener("click", (event) => event.stopPropagation());
    rescheduleInput.addEventListener("change", () => {
      if (!rescheduleInput.value) {
        return;
      }
      task.dueDate = rescheduleInput.value;
      persist();
      if (typeof options.onQuickReschedule === "function") {
        options.onQuickReschedule(task);
        return;
      }
      render();
    });
  }
  if (assignDateButton && assignDateInput) {
    assignDateInput.value = task.dueDate || "";
    assignDateButton.addEventListener("click", (event) => {
      event.stopPropagation();
      assignDateInput.classList.remove("hidden");
      window.requestAnimationFrame(() => {
        assignDateInput.focus();
        if (typeof assignDateInput.showPicker === "function") {
          assignDateInput.showPicker();
        } else {
          assignDateInput.click();
        }
      });
    });
    assignDateInput.addEventListener("click", (event) => event.stopPropagation());
    assignDateInput.addEventListener("blur", () => {
      window.setTimeout(() => assignDateInput.classList.add("hidden"), 120);
    });
    assignDateInput.addEventListener("change", () => {
      if (!assignDateInput.value) {
        return;
      }
      task.dueDate = assignDateInput.value;
      persist();
      assignDateInput.classList.add("hidden");
      render();
    });
  }
  deleteButton.addEventListener("click", (event) => {
    event.stopPropagation();
    state.tasks = state.tasks.filter((item) => item.id !== task.id);
    if (editingTaskId === task.id) {
      resetTaskEditor();
    }
    persist();
    render();
  });
  card.addEventListener("click", () => {
    if (typeof options.onEdit === "function") {
      options.onEdit(task);
      return;
    }
    editTask(task.id);
  });
  return card;
}

function openTaskList(listId) {
  state.selectedTaskListId = listId;
  if (!elements.tasksDialog?.open) {
    elements.tasksDialog?.showModal();
  }
  renderTasksDialog();
}

function openPastDueTasksDialog() {
  renderPastDueTasksDialog();
  if (!elements.pastDueTasksDialog?.open) {
    elements.pastDueTasksDialog?.showModal();
  }
}

function renderPastDueTasksDialog() {
  if (!elements.pastDueTasksList) {
    return;
  }

  const pastDueTasks = getPastDueTasks();
  elements.pastDueTasksList.innerHTML = "";

  if (!pastDueTasks.length) {
    elements.pastDueTasksList.innerHTML = '<div class="item-card"><p class="card-body">No past due tasks right now.</p></div>';
    return;
  }

  pastDueTasks.forEach((task) => {
    const card = buildTaskCard(task, false, {
      quickReschedule: true,
      onEdit: () => {
        elements.pastDueTasksDialog?.close();
        editTask(task.id);
      },
      onQuickReschedule: () => {
        renderPastDueTasksDialog();
        renderMyDay();
      },
    });
    elements.pastDueTasksList.appendChild(card);
  });
}

function renderSelectedWorkspace() {
  if (!elements.entryDetail || !elements.promptList || !elements.documentPreview) {
    return;
  }
  const selectedItem = getSelectedItem();
  const prompts = buildPrompts(selectedItem);
  elements.entryDetail.innerHTML = "";
  elements.promptList.innerHTML = "";

  if (!selectedItem) {
    elements.entryDetail.innerHTML = '<div class="item-card"><p class="card-body">Select a task, idea, or meeting note to inspect it here.</p></div>';
    elements.promptList.innerHTML = '<div class="item-card"><p class="card-body">HAL prompt suggestions will appear here.</p></div>';
    elements.documentPreview.textContent = "Select an item to preview an organized draft.";
    return;
  }

  elements.entryDetail.innerHTML = `
    <article class="item-card">
      <div class="card-title-row">
        <h3 class="card-title">${escapeHtml(selectedItem.title)}</h3>
        <span class="status-pill">${escapeHtml(selectedItem.sectionLabel)}</span>
      </div>
      <div class="card-meta">${selectedItem.meta.map((item) => `<span>${escapeHtml(item)}</span>`).join("")}</div>
      <p class="card-body">${escapeHtml(selectedItem.content || selectedItem.title)}</p>
    </article>
  `;

  prompts.forEach((prompt) => {
    const card = document.createElement("article");
    card.className = "item-card";
    card.innerHTML = `<h3 class="card-title">${escapeHtml(prompt.title)}</h3><p class="card-body">${escapeHtml(prompt.body)}</p>`;
    elements.promptList.appendChild(card);
  });

  elements.documentPreview.textContent = buildSelectedDraft(selectedItem);
}

function renderFollowUps() {
  if (!elements.followUpList) {
    return;
  }
  elements.followUpList.innerHTML = "";
  if (!state.followUps.length) {
    elements.followUpList.innerHTML = '<div class="item-card"><p class="card-body">Save an item here when it should become an external action later.</p></div>';
    return;
  }

  state.followUps.forEach((item) => {
    const card = document.createElement("article");
    card.className = "item-card";
    card.innerHTML = `
      <div class="card-title-row">
        <h3 class="card-title">${escapeHtml(item.title)}</h3>
        <button class="ghost-button" type="button">Delete</button>
      </div>
      <p class="card-body">${escapeHtml(truncate(item.summary, 120))}</p>
    `;
    card.querySelector("button").addEventListener("click", () => {
      state.followUps = state.followUps.filter((entry) => entry.id !== item.id);
      persist();
      renderFollowUps();
    });
    elements.followUpList.appendChild(card);
  });
}

async function saveCapture(event) {
  event?.preventDefault?.();
  const text = elements.content.value.trim();
  if (!text) {
    return;
  }

  if (state.pendingCapture) {
    completePendingCapture(text);
    return;
  }

  const cleanupIntent = detectSectionCleanupIntent(text);
  if (cleanupIntent) {
    applySectionCleanup(cleanupIntent);
    elements.captureForm.reset();
    resetCapturePlaceholder();
    persist();
    render();
    return;
  }

  const correctionIntent = detectCategoryCorrectionIntent(text);
  if (correctionIntent) {
    applyCategoryCorrection(correctionIntent, text);
    elements.captureForm.reset();
    resetCapturePlaceholder();
    persist();
    render();
    return;
  }

  setMicStatus("HAL is organizing...", "HAL is organizing what you said into the right place.");
  const interpretation = await interpretCaptureWithLearning(text, "");
  console.log("HAL saveCapture interpretation", interpretation);
  rememberPrompt(text, interpretation);
  state.correctionContext = null;

  if (interpretation.type === "follow-up") {
    promptFollowUp(interpretation.question, interpretation.supportText, interpretation.pendingCapture, interpretation.placeholder);
    return;
  }

  if (interpretation.type === "task") {
    const targetListId = interpretation.listId || state.taskLists[0].id;
    const duplicateTask = findDuplicateTask(interpretation.title, targetListId, interpretation.dueDate || "");
    if (duplicateTask) {
      const duplicateMessage = [
        "HAL found a similar existing task.",
        "",
        `Existing: ${duplicateTask.title}`,
        `List: ${getTaskListName(duplicateTask.listId)}`,
        `Date: ${duplicateTask.dueDate ? formatDisplayDate(duplicateTask.dueDate) : "No date"}`,
        "",
        "Create another copy anyway?",
      ].join("\n");

      const shouldCreateDuplicate = window.confirm(duplicateMessage);
      if (!shouldCreateDuplicate) {
        state.selected = { type: "task", id: duplicateTask.id };
        setMicStatus("Duplicate avoided.", "");
        render();
        return;
      }
    }

    const createdTask = {
      id: crypto.randomUUID(),
      title: interpretation.title,
      listId: targetListId,
      dueDate: interpretation.dueDate || "",
      dueTime: interpretation.dueTime || "",
      done: false,
      sourcePrompt: text,
      createdAt: new Date().toISOString(),
    };
    state.tasks.unshift(createdTask);
    state.lastCreated = { type: "task", id: createdTask.id };
    console.log("HAL created task", createdTask);
    state.selectedTaskListId = targetListId;
    state.selected = { type: "task", id: createdTask.id };
    if (elements.tasksSearch) {
      elements.tasksSearch.value = "";
    }
    if (elements.tasksDateFilter) {
      elements.tasksDateFilter.value = interpretation.dueDate || "";
    }
    if (elements.myDayDate && interpretation.dueDate) {
      elements.myDayDate.value = interpretation.dueDate;
    }
  } else if (interpretation.type === "meeting") {
    const createdMeeting = {
      id: crypto.randomUUID(),
      title: interpretation.title,
      person: interpretation.person,
      callDate: interpretation.dueDate || "",
      content: interpretation.formatted,
      createdAt: new Date().toISOString(),
    };
    state.meetingNotes.unshift(createdMeeting);
    state.lastCreated = { type: "meeting", id: createdMeeting.id };
    state.selected = { type: "meeting", id: createdMeeting.id };
  } else if (interpretation.type === "meeting-update") {
    if (!interpretation.target) {
      setMicStatus("HAL could not find the meeting note to update.", "");
      render();
      return;
    }
    const note = state.meetingNotes.find((entry) => entry.id === interpretation.target.id);
    if (!note) {
      setMicStatus("HAL could not find the meeting note to update.", "");
      render();
      return;
    }
    note.content = appendMeetingNoteContent(note.content, interpretation.formatted || interpretation.rawAppendText || "");
    note.sourcePrompt = note.sourcePrompt || text;
    learnMeetingTargetCorrection(text, note);
    state.lastCreated = { type: "meeting", id: note.id };
    state.selected = { type: "meeting", id: note.id };
  } else if (interpretation.type === "reminder") {
    const createdReminder = {
      id: crypto.randomUUID(),
      text: interpretation.title,
      at: interpretation.at,
      delivery: interpretation.delivery || "inApp",
      frequency: interpretation.frequency,
      recurrence: interpretation.recurrence,
    };
    state.reminders.unshift(createdReminder);
    state.lastCreated = { type: "reminder", id: createdReminder.id };
  } else if (interpretation.type === "reminder-delete") {
    state.reminders = state.reminders.filter((item) => item.id !== interpretation.target.id);
    if (editingReminderId === interpretation.target.id) {
      resetReminderEditor();
    }
    state.lastCreated = { type: null, id: null };
  } else if (interpretation.type === "reminder-delete-missing") {
    setMicStatus("No matching reminder found.", "");
    render();
    return;
  } else if (interpretation.type === "reminder-edit") {
    startReminderEdit(interpretation.target.id);
    if (elements.remindersDialog && !elements.remindersDialog.open) {
      elements.remindersDialog.showModal();
    }
  } else if (interpretation.type === "reminder-edit-missing") {
    setMicStatus("No matching reminder found.", "");
    render();
    return;
  } else {
    const createdIdea = {
      id: crypto.randomUUID(),
      title: interpretation.title,
      content: interpretation.formatted,
      archived: false,
      sourcePrompt: text,
      createdAt: new Date().toISOString(),
    };
    state.ideas.unshift(createdIdea);
    state.lastCreated = { type: "idea", id: createdIdea.id };
    state.selected = { type: "idea", id: createdIdea.id };
  }

  elements.captureForm.reset();
  resetCapturePlaceholder();
  persist();
  console.log("HAL state after persist", {
    selectedTaskListId: state.selectedTaskListId,
    tasksDateFilter: elements.tasksDateFilter?.value,
    myDayDate: elements.myDayDate?.value,
    tasks: state.tasks,
  });
  render();

  if (interpretation.type === "task") {
    setMicStatus("Task created.", "");
    speakHal("Task created.");
  } else if (interpretation.type === "meeting") {
    setMicStatus("Meeting note created.", "");
    speakHal("Meeting note created.");
  } else if (interpretation.type === "meeting-update") {
    setMicStatus("Meeting note updated.", "");
    speakHal("Meeting note updated.");
  } else if (interpretation.type === "reminder") {
    setMicStatus("Reminder created.", "");
    speakHal("Reminder created.");
  } else if (interpretation.type === "reminder-delete") {
    setMicStatus("Reminder deleted.", "");
    speakHal("Reminder deleted.");
  } else if (interpretation.type === "reminder-delete-missing") {
    setMicStatus("No matching reminder found.", "");
    speakHal("No matching reminder found.");
  } else if (interpretation.type === "reminder-edit") {
    setMicStatus("Reminder loaded for editing.", "");
    speakHal("Reminder loaded for editing.");
  } else if (interpretation.type === "reminder-edit-missing") {
    setMicStatus("No matching reminder found.", "");
    speakHal("No matching reminder found.");
  } else {
    setMicStatus("Idea created.", "");
    speakHal("Idea created.");
  }
}

async function submitCaptureFromVoice() {
  const text = elements.content.value.trim();
  if (!text) {
    setMicStatus("Nothing to save yet.", "");
    speakHal("Nothing to save yet.");
    return;
  }
  await saveCapture();
}

function completePendingCapture(text) {
  const pending = state.pendingCapture;
  if (!pending) {
    return;
  }

  if (pending.kind === "reminder-time") {
    const timeDetails = parseReminderTimeDetails(text.toLowerCase());
    if (!timeDetails.explicit) {
      setMicStatus("What time would you like to be reminded?", "Try something like 3:15 PM, 8 AM, or this afternoon.");
      if (elements.content) {
        elements.content.value = "";
        elements.content.placeholder = 'Example: "3:15 PM"';
        elements.content.focus();
      }
      return;
    }

    const createdReminder = {
      id: crypto.randomUUID(),
      text: pending.draft.title,
      at: `${pending.draft.date}T${timeDetails.time}`,
      delivery: pending.draft.delivery || "inApp",
      frequency: pending.draft.frequency,
      recurrence: pending.draft.recurrence,
    };
    state.reminders.unshift(createdReminder);
    state.lastCreated = { type: "reminder", id: createdReminder.id };
    state.pendingCapture = null;
    elements.captureForm.reset();
    resetCapturePlaceholder();
    persist();
    render();
    setMicStatus("Reminder created.", "");
    speakHal("Reminder created.");
  }
}

function promptFollowUp(question, supportText, pendingCapture, placeholder) {
  state.pendingCapture = pendingCapture;
  setMicStatus(question, supportText);
  speakHal(question);
  if (elements.content) {
    elements.content.value = "";
    elements.content.placeholder = placeholder || elements.content.placeholder;
    elements.content.focus();
  }
}

function resetCapturePlaceholder() {
  if (elements.content) {
    elements.content.placeholder = 'Example: "Create a task for me to follow up with Sarah tomorrow"';
  }
}

function previewCapture() {
  const text = elements.content.value.trim();
  if (!text) {
    return;
  }
}

function addManualTask(event) {
  event.preventDefault();
  const title = elements.taskInput.value.trim();
  if (!title) {
    return;
  }

  const payload = {
    title,
    listId: elements.taskListSelect?.value || state.selectedTaskListId || state.taskLists[0].id,
    dueDate: elements.taskDueDate.value,
    dueTime: elements.taskHasTime.checked ? elements.taskDueTime.value : "",
  };

  if (editingTaskId) {
    const existingTask = state.tasks.find((task) => task.id === editingTaskId);
    if (!existingTask) {
      return;
    }
    rememberCorrection({
      entityType: "task",
      entityId: existingTask.id,
      action: "edit",
      from: {
        title: existingTask.title,
        dueDate: existingTask.dueDate || "",
        dueTime: existingTask.dueTime || "",
        listId: existingTask.listId,
      },
      to: payload,
    });
    learnTitleCorrection("task", existingTask.title, payload.title);
    state.tasks = state.tasks.map((task) => (
      task.id === editingTaskId ? { ...task, ...payload } : task
    ));
    state.lastCreated = { type: "task", id: editingTaskId };
    setMicStatus("Task updated.", "");
  } else {
    state.tasks.unshift({
      id: crypto.randomUUID(),
      ...payload,
      notes: "",
      done: false,
      createdAt: new Date().toISOString(),
    });
    state.lastCreated = { type: "task", id: state.tasks[0].id };
    setMicStatus("Task created.", "");
  }

  if (elements.tasksSearch) {
    elements.tasksSearch.value = "";
  }
  state.selectedTaskListId = payload.listId;
  resetTaskEditor();
  persist();
  render();
  renderTasksDialog();
}

function createTaskList(event) {
  event?.preventDefault();
  const name = elements.newTaskListName.value.trim();
  if (!name) {
    elements.newTaskListName.focus();
    return;
  }

  const list = { id: crypto.randomUUID(), name };
  state.taskLists.push(list);
  state.selectedTaskListId = list.id;
  elements.newTaskListName.value = "";
  persist();
  render();
  elements.taskListDialog?.close();
  openTaskList(list.id);
  setMicStatus("Task list created.", "");
}

function addQuickTask(event) {
  event.preventDefault();
  const title = elements.quickTaskInput?.value.trim();
  if (!title) {
    elements.quickTaskInput?.focus();
    return;
  }
  const payload = {
    id: crypto.randomUUID(),
    title,
    listId: elements.quickTaskListSelect?.value || state.taskLists[0]?.id,
    dueDate: elements.quickTaskDueDate?.value || "",
    dueTime: elements.quickTaskHasTime?.checked ? (elements.quickTaskDueTime?.value || "") : "",
    notes: state.quickTaskDraftNote || "",
    done: false,
    createdAt: new Date().toISOString(),
  };
  state.tasks.unshift(payload);
  state.lastCreated = { type: "task", id: payload.id };
  state.selectedTaskListId = payload.listId;
  resetQuickTaskComposer();
  persist();
  render();
  setMicStatus("Task created.", "");
}

function renderTaskListSelect() {
  if (!elements.taskListSelect) {
    return;
  }

  const selectedId = editingTaskId
    ? (state.tasks.find((task) => task.id === editingTaskId)?.listId || state.selectedTaskListId || state.taskLists[0]?.id)
    : (state.selectedTaskListId || elements.taskListSelect.value || state.taskLists[0]?.id);

  elements.taskListSelect.innerHTML = state.taskLists
    .map((list) => `<option value="${escapeHtml(list.id)}">${escapeHtml(list.name)}</option>`)
    .join("");

  elements.taskListSelect.value = selectedId || state.taskLists[0]?.id || "";
}

function openTaskComposer() {
  editingTaskId = null;
  state.selectedTaskListId = state.taskLists[0]?.id || state.selectedTaskListId;
  if (!elements.tasksDialog?.open) {
    elements.tasksDialog.showModal();
  }
  resetTaskEditor();
  renderTasksDialog();
  elements.taskInput?.focus();
}

function editTaskList(id) {
  const list = state.taskLists.find((entry) => entry.id === id);
  if (!list) {
    return;
  }
  const nextName = window.prompt("Edit task list name", list.name);
  if (!nextName) {
    return;
  }
  list.name = nextName.trim() || list.name;
  persist();
  render();
}

function deleteTaskList(id) {
  if (state.taskLists.length === 1) {
    window.alert("HAL needs at least one task list.");
    return;
  }
  const list = state.taskLists.find((entry) => entry.id === id);
  if (!list) {
    return;
  }
  const confirmed = window.confirm(`Delete the "${list.name}" list? Its tasks will move to Regular Tasks.`);
  if (!confirmed) {
    return;
  }
  const fallbackListId = state.taskLists.find((entry) => entry.id !== id)?.id;
  state.tasks = state.tasks.map((task) => task.listId === id ? { ...task, listId: fallbackListId } : task);
  state.taskLists = state.taskLists.filter((entry) => entry.id !== id);
  if (state.selectedTaskListId === id) {
    state.selectedTaskListId = fallbackListId;
  }
  persist();
  render();
}

function editTask(id) {
  const task = state.tasks.find((entry) => entry.id === id);
  if (!task) {
    return;
  }
  editingTaskId = task.id;
  state.selectedTaskListId = task.listId;
  if (!elements.tasksDialog?.open) {
    elements.tasksDialog?.showModal();
  }
  elements.taskInput.value = task.title;
  if (elements.taskListSelect) {
    elements.taskListSelect.value = task.listId;
  }
  elements.taskDueDate.value = task.dueDate || "";
  elements.taskHasTime.checked = Boolean(task.dueTime);
  elements.taskDueTime.value = task.dueTime || "";
  syncTaskTimeField();
  syncTaskEditorState();
  renderTasksDialog();
  elements.taskInput.focus();
}

function resetTaskEditor() {
  editingTaskId = null;
  elements.taskForm?.reset();
  if (elements.taskListSelect) {
    elements.taskListSelect.value = state.selectedTaskListId || state.taskLists[0]?.id || "";
  }
  syncTaskTimeField();
  syncTaskEditorState();
}

function openTaskNoteEditor(id) {
  const task = state.tasks.find((entry) => entry.id === id);
  if (!task || !elements.taskNoteEditorContent) {
    return;
  }
  editingTaskNoteId = id;
  taskNoteEditorMode = "existing";
  if (elements.taskNoteEditorTitle) {
    elements.taskNoteEditorTitle.textContent = task.title;
  }
  elements.taskNoteEditorContent.innerHTML = task.notes || "";
  if (!elements.taskNoteEditorDialog?.open) {
    elements.taskNoteEditorDialog.showModal();
  }
  focusTaskNoteEditor();
}

function closeTaskNoteEditor() {
  editingTaskNoteId = null;
  taskNoteEditorMode = "existing";
  elements.taskNoteEditorDialog?.close();
}

function saveTaskNoteEditor(event) {
  event.preventDefault();
  if (!elements.taskNoteEditorContent) {
    return;
  }
  if (taskNoteEditorMode === "draft") {
    state.quickTaskDraftNote = sanitizeIdeaEditorHtml(elements.taskNoteEditorContent.innerHTML);
    syncQuickTaskNoteButton();
    closeTaskNoteEditor();
    setMicStatus("Task note saved.", "");
    return;
  }
  if (!editingTaskNoteId) {
    return;
  }
  const task = state.tasks.find((entry) => entry.id === editingTaskNoteId);
  if (!task) {
    return;
  }
  task.notes = sanitizeIdeaEditorHtml(elements.taskNoteEditorContent.innerHTML);
  persist();
  render();
  renderTasksDialog();
  renderPastDueTasksDialog();
  closeTaskNoteEditor();
  setMicStatus("Task note saved.", "");
}

function openDraftTaskNoteEditor() {
  taskNoteEditorMode = "draft";
  editingTaskNoteId = null;
  if (elements.taskNoteEditorTitle) {
    elements.taskNoteEditorTitle.textContent = elements.quickTaskInput?.value.trim() || "New Task Notes";
  }
  if (elements.taskNoteEditorContent) {
    elements.taskNoteEditorContent.innerHTML = state.quickTaskDraftNote || "";
  }
  if (!elements.taskNoteEditorDialog?.open) {
    elements.taskNoteEditorDialog.showModal();
  }
  focusTaskNoteEditor();
}

function syncTaskEditorState() {
  if (elements.taskSubmit) {
    elements.taskSubmit.textContent = editingTaskId ? "Save task" : "Add task";
  }
  elements.cancelTaskEdit?.classList.toggle("hidden", !editingTaskId);
}

function resetQuickTaskComposer() {
  elements.quickTaskForm?.reset();
  state.quickTaskDraftNote = "";
  if (elements.quickTaskListSelect) {
    elements.quickTaskListSelect.value = state.taskLists[0]?.id || "";
  }
  syncQuickTaskTimeField();
  syncQuickTaskNoteButton();
}

function syncTaskTimeField() {
  const hasTime = Boolean(elements.taskHasTime?.checked);
  elements.taskDueTime?.classList.toggle("hidden", !hasTime);
  if (!hasTime && elements.taskDueTime) {
    elements.taskDueTime.value = "";
  }
}

function syncQuickTaskTimeField() {
  const hasTime = Boolean(elements.quickTaskHasTime?.checked);
  elements.quickTaskDueTime?.classList.toggle("hidden", !hasTime);
  if (!hasTime && elements.quickTaskDueTime) {
    elements.quickTaskDueTime.value = "";
  }
}

function syncQuickTaskNoteButton() {
  if (!elements.openQuickTaskNoteEditor) {
    return;
  }
  const hasNote = Boolean(stripHtml(state.quickTaskDraftNote || ""));
  elements.openQuickTaskNoteEditor.classList.toggle("has-notes", hasNote);
  elements.openQuickTaskNoteEditor.textContent = hasNote ? "Note*" : "Note";
}

function showTaskDatePicker() {
  if (typeof elements.tasksDateFilter?.showPicker === "function") {
    elements.tasksDateFilter.showPicker();
  }
}

function showTaskDueDatePicker() {
  if (typeof elements.taskDueDate?.showPicker === "function") {
    elements.taskDueDate.showPicker();
  }
}

function showTaskTimePicker() {
  if (typeof elements.taskDueTime?.showPicker === "function") {
    elements.taskDueTime.showPicker();
  }
}

function showQuickTaskDueDatePicker() {
  if (typeof elements.quickTaskDueDate?.showPicker === "function") {
    elements.quickTaskDueDate.showPicker();
  }
}

function showQuickTaskTimePicker() {
  if (typeof elements.quickTaskDueTime?.showPicker === "function") {
    elements.quickTaskDueTime.showPicker();
  }
}

function clearTasksDateFilter() {
  if (elements.tasksDateFilter) {
    elements.tasksDateFilter.value = "";
  }
  renderTasksDialog();
}

function syncTasksDateFilterButton() {
  if (!elements.clearTasksDateFilter) {
    return;
  }
  elements.clearTasksDateFilter.classList.toggle("hidden", !elements.tasksDateFilter?.value);
}

function handleTaskNoteEditorToolbarClick(event) {
  const button = event.target.closest("[data-editor-command]");
  if (!button) {
    return;
  }
  event.preventDefault();
  focusTaskNoteEditor();
  if (button.dataset.editorCommand === "removeFormat") {
    document.execCommand("removeFormat", false);
    document.execCommand("unlink", false);
    return;
  }
  document.execCommand(button.dataset.editorCommand, false);
}

function handleTaskNoteEditorFontButtonClick(event) {
  event.preventDefault();
  const nextFont = window.prompt("Choose a font: IBM Plex Sans, Georgia, Arial, or Courier New", "IBM Plex Sans");
  if (!nextFont) {
    return;
  }
  focusTaskNoteEditor();
  document.execCommand("fontName", false, nextFont);
}

function handleTaskNoteEditorSizeButtonClick(event) {
  event.preventDefault();
  const nextSize = window.prompt("Choose a size: 2 small, 3 normal, 4 large, or 5 XL", "3");
  if (!nextSize) {
    return;
  }
  focusTaskNoteEditor();
  document.execCommand("fontSize", false, nextSize);
}

function preserveTaskNoteEditorSelection() {
  focusTaskNoteEditor();
}

function focusTaskNoteEditor() {
  elements.taskNoteEditorContent?.focus();
}

function toggleCompletedTasksVisibility() {
  state.showCompletedTasks = !state.showCompletedTasks;
  persist();
  renderTasksDialog();
}

function addReminder(event) {
  event.preventDefault();
  const text = elements.reminderText.value.trim();
  if (!text || !elements.reminderDateTime.value) {
    return;
  }

  const payload = {
    text,
    at: elements.reminderDateTime.value,
    delivery: elements.reminderDelivery.value || "inApp",
    frequency: elements.reminderFrequency.value,
    recurrence: buildReminderRecurrence(),
  };

  if (editingReminderId) {
    state.reminders = state.reminders.map((reminder) =>
      reminder.id === editingReminderId ? { ...reminder, ...payload } : reminder
    );
  } else {
    state.reminders.unshift({
      id: crypto.randomUUID(),
      ...payload,
    });
  }

  resetReminderEditor();
  persist();
  renderReminders();
}

function syncSmsSettingsUI() {
  if (elements.smsNotificationsEnabled) {
    elements.smsNotificationsEnabled.checked = Boolean(state.notificationSettings?.smsEnabled);
  }
  if (elements.smsPhone) {
    elements.smsPhone.value = state.notificationSettings?.phoneNumber || "";
  }
}

function syncTeamsSettingsUI() {
  if (elements.teamsNotificationsEnabled) {
    elements.teamsNotificationsEnabled.checked = Boolean(state.notificationSettings?.teamsEnabled);
  }
}

function syncMyDaySettingsUI() {
  state.calendarSettings = {
    ...state.calendarSettings,
    provider: "csv",
  };
}

function saveTeamsSettings() {
  state.notificationSettings = {
    ...state.notificationSettings,
    teamsEnabled: Boolean(elements.teamsNotificationsEnabled?.checked),
  };
  persist();
  syncTeamsSettingsUI();
  setMicStatus("Teams settings saved.", "");
  speakHal("Teams settings saved.");
}

function saveMyDaySettings(event) {
  event.preventDefault();
  state.calendarSettings = {
    ...state.calendarSettings,
    provider: "csv",
  };
  persist();
  syncMyDaySettingsUI();
  syncMyDayCalendarUI();
  renderMyDay();
  renderTeamsStatus();
  elements.myDaySettingsDialog?.close();
  void syncCsvCalendar(true);
}

function saveSmsSettings() {
  state.notificationSettings = {
    ...state.notificationSettings,
    smsEnabled: Boolean(elements.smsNotificationsEnabled?.checked),
    phoneNumber: normalizePhoneNumber(elements.smsPhone?.value || ""),
  };
  persist();
  syncSmsSettingsUI();
  setMicStatus("Text settings saved.", "");
  speakHal("Text settings saved.");
}

async function refreshSmsStatus() {
  if (!elements.smsSettingsStatus) {
    return;
  }

  try {
    const response = await fetch("/api/notifications/status");
    const payload = await response.json();
    const providerLabel = payload.providerConfigured ? "SMS ready" : "SMS not connected";
    elements.smsSettingsStatus.textContent = payload.smsEnabled ? providerLabel : "Text delivery off";
    elements.smsSettingsStatus.title = payload.providerConfigured
      ? `Provider: ${payload.provider || "configured"}`
      : "Add SMS provider credentials on the backend to enable text delivery.";
  } catch (_error) {
    elements.smsSettingsStatus.textContent = "SMS status unavailable";
    elements.smsSettingsStatus.title = "HAL could not reach the local notification service.";
  }
}

async function refreshTeamsStatus() {
  if (!elements.teamsSettingsStatus) {
    return;
  }

  try {
    const response = await fetch("/api/notifications/teams/status");
    const payload = await response.json();
    if (!payload.authenticated) {
      elements.teamsSettingsStatus.textContent = "Sign in required";
      elements.teamsSettingsStatus.title = "Sign in with Microsoft before HAL can send Teams notifications.";
      return;
    }
    if (!payload.providerConfigured) {
      elements.teamsSettingsStatus.textContent = "Teams not ready";
      elements.teamsSettingsStatus.title = payload.reason || "HAL still needs Teams notification configuration.";
      return;
    }
    elements.teamsSettingsStatus.textContent = payload.teamsEnabled ? "Teams ready" : "Teams delivery off";
    elements.teamsSettingsStatus.title = payload.detail || "HAL can send Teams activity notifications.";
  } catch (_error) {
    elements.teamsSettingsStatus.textContent = "Teams status unavailable";
    elements.teamsSettingsStatus.title = "HAL could not reach the local Teams notification service.";
  }
}

async function refreshGoogleCalendarStatus() {
  if (!elements.googleCalendarStatus) {
    return;
  }

  try {
    const response = await fetch("/api/google-calendar/status");
    const payload = await response.json();
    elements.googleCalendarStatus.textContent = payload.configured ? "Google ready" : "Google not connected";
    elements.googleCalendarStatus.title = payload.configured
      ? "HAL has Google Calendar credentials configured."
      : "Google Calendar sync is scaffolded, but OAuth credentials have not been added yet.";
  } catch (_error) {
    elements.googleCalendarStatus.textContent = "Google status unavailable";
    elements.googleCalendarStatus.title = "HAL could not reach the local Google calendar status route.";
  }
}

async function sendSmsTest() {
  const phoneNumber = normalizePhoneNumber(elements.smsPhone?.value || state.notificationSettings?.phoneNumber || "");
  if (!phoneNumber) {
    setMicStatus("Add your phone number first.", "");
    speakHal("Add your phone number first.");
    return;
  }

  state.notificationSettings = {
    ...state.notificationSettings,
    smsEnabled: Boolean(elements.smsNotificationsEnabled?.checked),
    phoneNumber,
  };
  persist();
  syncSmsSettingsUI();

  try {
    const response = await fetch("/api/notifications/test", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        phoneNumber,
        message: "HAL test message. Your text notification architecture is connected.",
      }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not send the test text.");
    }
    setMicStatus(payload.message || "Test text queued.", "");
    speakHal(payload.message || "Test text queued.");
    await refreshSmsStatus();
  } catch (error) {
    setMicStatus(error.message || "HAL could not send the test text.", "");
    speakHal("HAL could not send the test text.");
  }
}

async function sendTeamsTest() {
  state.notificationSettings = {
    ...state.notificationSettings,
    teamsEnabled: Boolean(elements.teamsNotificationsEnabled?.checked),
  };
  persist();
  syncTeamsSettingsUI();

  try {
    const response = await fetch("/api/notifications/teams/test", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: "HAL test reminder. Your Teams notification path is connected.",
      }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not send the Teams notification.");
    }
    setMicStatus(payload.message || "Teams notification sent.", "");
    speakHal(payload.message || "Teams notification sent.");
    await refreshTeamsStatus();
  } catch (error) {
    setMicStatus(error.message || "HAL could not send the Teams notification.", "");
    speakHal("HAL could not send the Teams notification.");
  }
}

function buildReminderRecurrence() {
  if (elements.reminderFrequency.value !== "recurring") {
    return null;
  }

  if (elements.reminderRecurringMode.value === "monthlyWeekday") {
    return {
      mode: "monthlyWeekday",
      ordinal: elements.reminderMonthlyOrdinal.value,
      weekday: elements.reminderMonthlyWeekday.value,
    };
  }

  return {
    mode: "interval",
    interval: elements.reminderRecurringInterval.value,
  };
}

function addQuickLink() {
  const label = window.prompt("Link label");
  if (!label) {
    return;
  }
  const url = window.prompt("Link URL");
  if (!url) {
    return;
  }

  state.quickLinks.push({ id: crypto.randomUUID(), label, url });
  persist();
  renderQuickLinks();
}

function editQuote() {
  const nextQuote = window.prompt("Edit your top reminder or quote", state.quote);
  if (!nextQuote) {
    return;
  }
  state.quote = nextQuote;
  persist();
  renderQuote();
}

function addCalendarEvent() {
  const title = window.prompt("Event title");
  if (!title) {
    return;
  }
  const time = window.prompt("Event time (optional)", "10:00 AM");
  state.calendarEvents.push({
    id: crypto.randomUUID(),
    title,
    date: elements.myDayDate.value,
    time: time || "",
    source: "manual",
  });
  persist();
  renderMyDay();
}

function handleMyDayDateChange() {
  renderMyDay();
  void syncCsvCalendar(true);
}

function showMyDayPicker() {
  if (!elements.myDayDate) {
    return;
  }

  if (typeof elements.myDayDate.showPicker === "function") {
    elements.myDayDate.showPicker();
    return;
  }

  elements.myDayDate.focus();
  elements.myDayDate.click();
}

function handleOutlookCalendarAction() {
  void syncCsvCalendar(false);
}

function promptCalendarCsvUpload() {
  elements.uploadCalendarCsvInput?.click();
}

async function handleCalendarCsvUpload(event) {
  const file = event.target?.files?.[0];
  if (!file) {
    return;
  }

  try {
    const csv = await file.text();
    const previewEvents = parseClientCsvCalendarEvents(csv);
    state.calendarSettings = {
      ...state.calendarSettings,
      uploadedCsvContent: csv,
      uploadedCsvName: file.name,
      uploadedCsvImportedAt: new Date().toISOString(),
    };
    state.calendarEvents = state.calendarEvents.filter((item) => item.source !== "csv");
    persist();
    void syncCsvCalendar(false);
    setMicStatus(`Imported ${previewEvents.length} calendar events.`, "");
  } catch (error) {
    setMicStatus("HAL could not import that calendar CSV.", error.message || "");
  } finally {
    event.target.value = "";
  }
}

async function refreshMicrosoftSession() {
  try {
    const response = await fetch("/api/session", {
      credentials: "include",
    });
    const payload = await response.json();
    microsoftSession = {
      authenticated: Boolean(payload.authenticated),
      user: payload.user || null,
    };
  } catch {
    microsoftSession = {
      authenticated: false,
      user: null,
    };
  }

  renderTeamsStatus();
  syncMyDayCalendarUI();
  void syncCsvCalendar(true);
}

async function submitAuthForm(event) {
  event.preventDefault();
  if (!elements.authPassword || !elements.authStatus) {
    return;
  }

  const password = elements.authPassword.value;
  const rememberDevice = Boolean(elements.authRememberDevice?.checked);
  if (!password) {
    elements.authStatus.textContent = "Enter your password first.";
    return;
  }

  elements.authStatus.textContent = "Unlocking HAL...";
  updateAuthDebugState({
    lastLogin: `submitting (${rememberDevice ? "remembered" : "session-only"})`,
    lastError: "",
  });

  try {
    const response = await fetch("/api/auth/login", {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password, rememberDevice }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not unlock.");
    }

    if (rememberDevice && payload.rememberToken) {
      localStorage.setItem(STORAGE_KEYS.remember, payload.rememberToken);
    } else {
      localStorage.removeItem(STORAGE_KEYS.remember);
    }
    updateAuthDebugState({
      rememberTokenPresent: Boolean(localStorage.getItem(STORAGE_KEYS.remember)),
    });

    elements.authPassword.value = "";
    if (elements.authRememberDevice) {
      elements.authRememberDevice.checked = false;
    }
    authState.checked = true;
    authState.authenticated = true;
    authState.passwordConfigured = payload.status?.passwordConfigured !== false;
    authState.usingBootstrapPassword = Boolean(payload.status?.usingBootstrapPassword);
    updateAuthDebugState({
      lastLogin: "success",
    });
    renderAccessState();
    await startAuthenticatedApp();
    setMicStatus("HAL unlocked.", "");
  } catch (error) {
    authState.authenticated = false;
    appBoot.started = false;
    updateAuthDebugState({
      lastLogin: "failed",
      lastError: error.message || "HAL could not unlock.",
    });
    renderAccessState();
    elements.authStatus.textContent = error.message || "HAL could not unlock.";
  }
}

function openSecuritySettingsDialog() {
  elements.securityStatus.textContent = authState.usingBootstrapPassword
    ? "HAL is still using the bootstrap password. Updating it here will store your own app password."
    : "Use a strong password you can keep somewhere safe.";
  elements.securitySettingsDialog?.showModal();
}

async function saveSecuritySettings(event) {
  event.preventDefault();

  const currentPassword = elements.securityCurrentPassword?.value || "";
  const newPassword = elements.securityNewPassword?.value || "";
  const confirmPassword = elements.securityConfirmPassword?.value || "";

  if (newPassword !== confirmPassword) {
    elements.securityStatus.textContent = "The new passwords do not match.";
    return;
  }

  if (newPassword.length < 8) {
    elements.securityStatus.textContent = "Use a password with at least 8 characters.";
    return;
  }

  elements.securityStatus.textContent = "Updating password...";

  try {
    const response = await fetch("/api/auth/password", {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        currentPassword,
        newPassword,
      }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not update the password.");
    }

    elements.securityCurrentPassword.value = "";
    elements.securityNewPassword.value = "";
    elements.securityConfirmPassword.value = "";
    authState.usingBootstrapPassword = false;
    elements.securityStatus.textContent = "Password updated.";
    setMicStatus("Security settings updated.", "");
  } catch (error) {
    elements.securityStatus.textContent = error.message || "HAL could not update the password.";
  }
}

async function logoutHalAccess() {
  try {
    await fetch("/api/auth/logout", {
      method: "POST",
      credentials: "include",
    });
  } catch {
    // Even if the request fails, treat the local session as closed.
  }

  localStorage.removeItem(STORAGE_KEYS.remember);
  authState.authenticated = false;
  appBoot.started = false;
  updateAuthDebugState({
    rememberTokenPresent: false,
    lastLogin: "logged out",
    lastBoot: "reset after logout",
  });
  renderAccessState();
  elements.securitySettingsDialog?.close();
  setMicStatus("HAL locked.", "");
}

async function tryRestoreRememberedAccess() {
  const token = localStorage.getItem(STORAGE_KEYS.remember);
  if (!token) {
    updateAuthDebugState({
      rememberTokenPresent: false,
      lastRestore: "no token found",
    });
    return false;
  }
  updateAuthDebugState({
    rememberTokenPresent: true,
    lastRestore: "attempting restore",
  });

  try {
    const response = await fetch("/api/auth/restore", {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not restore the remembered device session.");
    }

    authState.checked = true;
    authState.authenticated = true;
    authState.passwordConfigured = payload.status?.passwordConfigured !== false;
    authState.usingBootstrapPassword = Boolean(payload.status?.usingBootstrapPassword);
    updateAuthDebugState({
      lastRestore: "restore succeeded",
      lastError: "",
    });
    return true;
  } catch {
    localStorage.removeItem(STORAGE_KEYS.remember);
    authState.authenticated = false;
    updateAuthDebugState({
      rememberTokenPresent: false,
      lastRestore: "restore failed",
      lastError: "remember restore failed",
    });
    return false;
  }
}

function syncMyDayCalendarUI() {
  const provider = "csv";
  if (elements.syncOutlookCalendar) {
    elements.syncOutlookCalendar.textContent = "Refresh calendar";
  }

  if (elements.myDayCalendarStatus) {
    const selectedDate = elements.myDayDate?.value || "";
    const csvEventCount = state.calendarEvents.filter((event) => event.source === "csv" && event.date === selectedDate).length;
    const usingUpload = Boolean(state.calendarSettings?.uploadedCsvContent);
    elements.myDayCalendarStatus.textContent = state.lastCalendarSyncDate === selectedDate && !csvEventCount
      ? `No schedule items were found for ${formatDisplayDate(selectedDate)}.`
      : usingUpload
        ? `Using uploaded calendar file${state.calendarSettings?.uploadedCsvName ? `: ${state.calendarSettings.uploadedCsvName}` : ""}.`
        : "";
  }
}

async function syncCsvCalendar(silent = false) {
  const selectedDate = elements.myDayDate?.value;
  if (!selectedDate) {
    syncMyDayCalendarUI();
    return;
  }

  if (!isLocalHalHost()) {
    await refreshHostedCalendarSettingsFromServer();
  }

  const uploadedCsvContent = state.calendarSettings?.uploadedCsvContent || "";
  if (elements.myDayCalendarStatus) {
    elements.myDayCalendarStatus.textContent = uploadedCsvContent
      ? `Refreshing uploaded calendar for ${formatDisplayDate(selectedDate)}...`
      : `Loading calendar CSV for ${formatDisplayDate(selectedDate)}...`;
  }

  try {
    let events = [];
    if (uploadedCsvContent) {
      events = parseClientCsvCalendarEvents(uploadedCsvContent)
        .filter((event) => event.date === selectedDate);
    } else {
      const response = await fetch(`/api/calendar/csv/day?date=${encodeURIComponent(selectedDate)}`);
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "HAL could not read the calendar CSV.");
      }
      events = payload.events || [];
    }

    state.calendarEvents = state.calendarEvents.filter((event) => event.source !== "csv" || event.date !== selectedDate);
    state.calendarEvents.push(...events);
    state.lastCalendarSyncDate = selectedDate;
    persist();
    renderMyDay();
    if (!silent) {
      if (events.length) {
        setMicStatus("Calendar refreshed.", "");
        speakHal("Calendar refreshed.");
      } else {
        setMicStatus("No calendar events found for that date.", "");
        speakHal("No calendar events found for that date.");
      }
    }
  } catch (error) {
    if (elements.myDayCalendarStatus) {
      elements.myDayCalendarStatus.textContent = error.message || "HAL could not read the calendar.";
    }
    if (!silent) {
      setMicStatus(error.message || "HAL could not read the calendar.", "");
      speakHal("HAL could not read the calendar.");
    }
  }
}

async function refreshHostedCalendarSettingsFromServer() {
  try {
    const response = await fetch("/api/state", {
      credentials: "include",
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not refresh the hosted calendar state.");
    }

    const remoteSettings = payload.state?.calendarSettings || {};
    const nextUploadedContent = remoteSettings.uploadedCsvContent || "";
    const nextUploadedName = remoteSettings.uploadedCsvName || "";
    const nextImportedAt = remoteSettings.uploadedCsvImportedAt || "";

    const changed =
      nextUploadedContent !== (state.calendarSettings?.uploadedCsvContent || "") ||
      nextUploadedName !== (state.calendarSettings?.uploadedCsvName || "") ||
      nextImportedAt !== (state.calendarSettings?.uploadedCsvImportedAt || "");

    if (changed) {
      state.calendarSettings = {
        ...state.calendarSettings,
        uploadedCsvContent: nextUploadedContent,
        uploadedCsvName: nextUploadedName,
        uploadedCsvImportedAt: nextImportedAt,
      };
    }
  } catch {
    // Keep the current in-memory copy if the hosted refresh check fails.
  }
}

function saveSelectedToFollowUp() {
  if (!elements.followUpList) {
    window.alert("Follow-up queue panel is not active in this layout yet.");
    return;
  }
  const selectedItem = getSelectedItem();
  if (!selectedItem) {
    window.alert("Select a task, idea, or meeting note first.");
    return;
  }

  state.followUps.unshift({
    id: crypto.randomUUID(),
    title: selectedItem.title,
    summary: selectedItem.content || selectedItem.title,
  });
  persist();
  renderFollowUps();
}

function exportIdea(idea, type) {
  if (type === "word") {
    downloadBlob(
      new Blob([buildWordExportDocument(idea.title, idea.content)], { type: "application/msword;charset=utf-8" }),
      `${safeFilename(idea.title)}.doc`,
    );
    return;
  }
  if (type === "powerpoint") {
    downloadBlob(new Blob([`${idea.title}\n\n${stripHtml(idea.content)}`], { type: "text/plain;charset=utf-8" }), `${safeFilename(idea.title)}-deck-outline.txt`);
    return;
  }
  downloadBlob(new Blob([`Title,Idea\n"${escapeCsv(idea.title)}","${escapeCsv(stripHtml(idea.content))}"`], { type: "text/csv;charset=utf-8" }), `${safeFilename(idea.title)}.csv`);
}

function openIdeasDialog() {
  renderIdeas();
  if (!elements.ideasDialog?.open) {
    elements.ideasDialog.showModal();
  }
}

function openNewIdeaEditor() {
  editingIdeaId = null;
  elements.ideaEditorTitle.textContent = "New Idea";
  elements.ideaEditorName.value = "";
  elements.ideaEditorContent.innerHTML = "";
  elements.archiveIdeaButton.textContent = "Archive";
  elements.archiveIdeaButton.classList.add("hidden");
  elements.deleteIdeaButton.classList.add("hidden");
  if (!elements.ideaEditorDialog?.open) {
    elements.ideaEditorDialog.showModal();
  }
  elements.ideaEditorName.focus();
}

function openIdeaEditor(id) {
  const idea = state.ideas.find((entry) => entry.id === id);
  if (!idea) {
    return;
  }
  editingIdeaId = id;
  elements.ideaEditorTitle.textContent = idea.title;
  elements.ideaEditorName.value = idea.title;
  elements.ideaEditorContent.innerHTML = normalizeIdeaEditorHtml(idea.content);
  elements.archiveIdeaButton.textContent = idea.archived ? "Restore" : "Archive";
  elements.archiveIdeaButton.classList.remove("hidden");
  elements.deleteIdeaButton.classList.remove("hidden");
  if (!elements.ideaEditorDialog?.open) {
    elements.ideaEditorDialog.showModal();
  }
}

function saveIdeaEditor(event) {
  event.preventDefault();
  const title = elements.ideaEditorName.value.trim();
  const content = sanitizeIdeaEditorHtml(elements.ideaEditorContent.innerHTML);
  if (!title || !content) {
    return;
  }
  const idea = state.ideas.find((entry) => entry.id === editingIdeaId);
  if (!idea) {
    const createdIdea = {
      id: crypto.randomUUID(),
      title,
      content,
      archived: false,
      createdAt: new Date().toISOString(),
    };
    state.ideas.unshift(createdIdea);
    editingIdeaId = createdIdea.id;
    state.lastCreated = { type: "idea", id: createdIdea.id };
    elements.ideaEditorTitle.textContent = createdIdea.title;
    elements.archiveIdeaButton.classList.remove("hidden");
    elements.deleteIdeaButton.classList.remove("hidden");
    persist();
    render();
    setMicStatus("Idea created.", "");
    return;
  }
  rememberCorrection({
    entityType: "idea",
    entityId: idea.id,
    action: "edit",
    from: {
      title: idea.title,
      content: idea.content,
    },
    to: {
      title,
      content,
    },
  });
  learnTitleCorrection("idea", idea.title, title);
  idea.title = title;
  idea.content = content;
  persist();
  render();
  elements.ideaEditorTitle.textContent = idea.title;
  setMicStatus("Idea updated.", "");
}

function toggleIdeaArchiveFromEditor() {
  const idea = state.ideas.find((entry) => entry.id === editingIdeaId);
  if (!idea) {
    return;
  }
  idea.archived = !idea.archived;
  elements.archiveIdeaButton.textContent = idea.archived ? "Restore" : "Archive";
  persist();
  render();
}

function deleteIdeaFromEditor() {
  const idea = state.ideas.find((entry) => entry.id === editingIdeaId);
  if (!idea) {
    return;
  }
  const confirmed = window.confirm(`Delete "${idea.title}"?`);
  if (!confirmed) {
    return;
  }
  state.ideas = state.ideas.filter((entry) => entry.id !== idea.id);
  editingIdeaId = null;
  elements.ideaEditorDialog?.close();
  persist();
  render();
}

function exportCurrentIdeaFromEditor(type) {
  const idea = state.ideas.find((entry) => entry.id === editingIdeaId);
  if (!idea) {
    return;
  }
  exportIdea(idea, type);
}

function handleIdeaEditorToolbarClick(event) {
  const button = event.target.closest("[data-editor-command]");
  if (!button) {
    return;
  }
  event.preventDefault();
  focusIdeaEditor();
  if (button.dataset.editorCommand === "removeFormat") {
    document.execCommand("removeFormat", false);
    document.execCommand("unlink", false);
    return;
  }
  document.execCommand(button.dataset.editorCommand, false);
}

function handleIdeaEditorFontButtonClick(event) {
  event.preventDefault();
  const nextFont = window.prompt("Choose a font: IBM Plex Sans, Georgia, Arial, or Courier New", "IBM Plex Sans");
  if (!nextFont) {
    return;
  }
  focusIdeaEditor();
  document.execCommand("fontName", false, nextFont);
}

function handleIdeaEditorSizeButtonClick(event) {
  event.preventDefault();
  const nextSize = window.prompt("Choose a size: 2 small, 3 normal, 4 large, or 5 XL", "3");
  if (!nextSize) {
    return;
  }
  focusIdeaEditor();
  document.execCommand("fontSize", false, nextSize);
}

function focusIdeaEditor() {
  elements.ideaEditorContent?.focus();
}

function preserveIdeaEditorSelection(event) {
  if (!event.target.closest("button")) {
    return;
  }
  event.preventDefault();
}

function normalizeIdeaEditorHtml(content) {
  const raw = String(content || "").trim();
  if (!raw) {
    return "";
  }
  if (/<[a-z][\s\S]*>/i.test(raw)) {
    return raw;
  }
  return escapeHtml(raw).replace(/\n/g, "<br>");
}

function sanitizeIdeaEditorHtml(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+="[^"]*"/gi, "")
    .replace(/\son\w+='[^']*'/gi, "")
    .replace(/\sstyle=""/gi, "")
    .trim();
}

function buildWordExportDocument(title, content) {
  const cleanedContent = normalizeHtmlForWordExport(content);
  return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8">
  <title>${escapeHtml(title)}</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    body {
      font-family: Calibri, Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.45;
      color: #1f1f1f;
      margin: 24px;
    }
    h1, h2, h3, p {
      margin: 0 0 12px;
    }
    ul, ol {
      margin: 0 0 12px 24px;
      padding: 0;
    }
    li {
      margin: 0 0 6px;
    }
    strong, b {
      font-weight: 700;
    }
    em, i {
      font-style: italic;
    }
    u {
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  ${cleanedContent || "<p></p>"}
</body>
</html>`;
}

function normalizeHtmlForWordExport(content) {
  const raw = normalizeIdeaEditorHtml(content);
  if (!raw) {
    return "";
  }

  const container = document.createElement("div");
  container.innerHTML = sanitizeIdeaEditorHtml(raw);

  container.querySelectorAll("div").forEach((div) => {
    const hasBlockChild = Array.from(div.children).some((child) => /^(DIV|P|UL|OL|LI|H1|H2|H3|H4|H5|H6)$/.test(child.tagName));
    if (hasBlockChild) {
      return;
    }

    const paragraph = document.createElement("p");
    paragraph.innerHTML = div.innerHTML.trim() || "<br>";
    div.replaceWith(paragraph);
  });

  container.querySelectorAll("br").forEach((lineBreak) => {
    const parent = lineBreak.parentElement;
    if (parent && parent.tagName === "DIV" && !parent.textContent.trim()) {
      const paragraph = document.createElement("p");
      paragraph.innerHTML = "<br>";
      parent.replaceWith(paragraph);
    }
  });

  container.querySelectorAll("*").forEach((node) => {
    if (!node.textContent.replace(/\u00a0/g, " ").trim() && !["BR", "LI", "UL", "OL"].includes(node.tagName)) {
      if (node.tagName === "P") {
        node.remove();
      }
    }
  });

  return container.innerHTML
    .replace(/&nbsp;/gi, " ")
    .replace(/<p>\s*<\/p>/gi, "")
    .replace(/(<br>\s*){3,}/gi, "<br><br>")
    .trim();
}

function stripHtml(value) {
  const container = document.createElement("div");
  container.innerHTML = String(value || "");
  return (container.textContent || container.innerText || "").trim();
}

function formatStructuredNoteForDisplay(content) {
  const raw = String(content || "").trim();
  if (!raw) {
    return "";
  }

  if (/<(ul|ol|li|p|div|br)\b/i.test(raw)) {
    return sanitizeIdeaEditorHtml(raw);
  }

  const normalized = escapeHtml(raw).replace(/\r/g, "");
  const lines = normalized.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) {
    return "";
  }

  const bulletLines = lines.filter((line) => /^[-*]\s+/.test(line));
  if (!bulletLines.length) {
    return normalized.replace(/\n/g, "<br>");
  }

  const headingLines = lines.filter((line) => !/^[-*]\s+/.test(line));
  const headingHtml = headingLines.length
    ? headingLines.map((line) => `<p>${line}</p>`).join("")
    : "";
  const bulletHtml = `<ul>${bulletLines.map((line) => `<li>${line.replace(/^[-*]\s+/, "")}</li>`).join("")}</ul>`;
  return `${headingHtml}${bulletHtml}`;
}

function formatFormattedIdeaPreview(content) {
  const raw = String(content || "").trim();
  if (!raw) {
    return "";
  }

  if (/<(ul|ol|li|p|div|br|strong|em|u)\b/i.test(raw)) {
    return sanitizeIdeaEditorHtml(raw);
  }

  const normalized = escapeHtml(raw).replace(/\r/g, "");
  const lines = normalized.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) {
    return "";
  }

  const bulletLines = lines.filter((line) => /^[-*]\s+/.test(line));
  if (!bulletLines.length) {
    return normalized.replace(/\n/g, "<br>");
  }

  const headingLines = lines.filter((line) => !/^[-*]\s+/.test(line));
  const headingHtml = headingLines.length
    ? headingLines.map((line) => `<p>${line}</p>`).join("")
    : "";
  const bulletHtml = `<ul>${bulletLines.map((line) => `<li>${line.replace(/^[-*]\s+/, "")}</li>`).join("")}</ul>`;
  return `${headingHtml}${bulletHtml}`;
}

function copyDraft() {
  if (!elements.documentPreview) {
    window.alert("Draft preview panel is not active in this layout yet.");
    return;
  }
  navigator.clipboard.writeText(elements.documentPreview.textContent)
    .then(() => window.alert("Draft copied."))
    .catch(() => window.alert("Clipboard access was blocked."));
}

function downloadDraft() {
  if (!elements.documentPreview) {
    window.alert("Draft preview panel is not active in this layout yet.");
    return;
  }
  downloadBlob(new Blob([elements.documentPreview.textContent], { type: "text/plain;charset=utf-8" }), "hal-draft.txt");
}

function downloadMarkdown() {
  if (!elements.documentPreview) {
    window.alert("Draft preview panel is not active in this layout yet.");
    return;
  }
  const selectedItem = getSelectedItem();
  if (!selectedItem) {
    window.alert("Select an item first.");
    return;
  }
  downloadBlob(new Blob([buildMarkdownDraft(selectedItem)], { type: "text/markdown;charset=utf-8" }), "hal-draft.md");
}

function downloadArchive() {
  downloadBackupArchive();
}

function clearAllData() {
  const confirmed = window.confirm("Clear HAL's local architecture data?");
  if (!confirmed) {
    return;
  }
  localStorage.removeItem(STORAGE_KEYS.app);
  window.location.reload();
}

function seedDemoData() {
  state.tasks = [
    { id: crypto.randomUUID(), title: "Follow up on pricing update", listId: state.taskLists[0].id, dueDate: todayString(), done: false, createdAt: new Date().toISOString() },
    { id: crypto.randomUUID(), title: "Review workflow idea notes", listId: state.taskLists[0].id, dueDate: tomorrowString(), done: false, createdAt: new Date().toISOString() },
  ];
  state.ideas = [
    { id: crypto.randomUUID(), title: "Workflow assistant for recurring follow-up", content: "Create a local-first assistant that captures loose thoughts and turns them into structured follow-up systems.", archived: false, createdAt: new Date().toISOString() },
  ];
  state.meetingNotes = [
    { id: crypto.randomUUID(), title: "Call with Sarah", person: "Sarah", callDate: todayString(), content: "Talking Points\n- Confirm timeline\n- Clarify deliverables\n- Define next steps", createdAt: new Date().toISOString() },
  ];
  state.calendarEvents = [
    { id: crypto.randomUUID(), title: "Supplier call", date: todayString(), time: "2:00 PM" },
  ];
  persist();
  render();
}

function openBackupDialog() {
  renderBackupSummary();
  void refreshBackupSnapshots();
  elements.backupDialog?.showModal();
}

function renderBackupSummary() {
  if (!elements.backupStatus) {
    return;
  }
  const summary = [
    `${state.tasks.length} tasks`,
    `${state.ideas.length} ideas`,
    `${state.meetingNotes.length} meeting notes`,
    `${state.reminders.length} reminders`,
  ].join(" | ");
  const mirrorLabel = serverStateSync.lastMeta?.savedAt
    ? ` Local mirror updated ${new Date(serverStateSync.lastMeta.savedAt).toLocaleString("en-US")}.`
    : " Local mirror is waiting for its first sync.";
  elements.backupStatus.textContent = `HAL is ready to back up ${summary}.${mirrorLabel}`;
}

function buildPersistedState() {
  return {
    _meta: {
      savedAt: new Date().toISOString(),
      storageVersion: 1,
      syncMode: "local-server-mirror",
    },
    theme: state.theme,
    voiceResponsesEnabled: state.voiceResponsesEnabled !== false,
    quote: state.quote,
    quickLinks: state.quickLinks,
    taskLists: state.taskLists,
    selectedTaskListId: state.selectedTaskListId,
    showCompletedTasks: Boolean(state.showCompletedTasks),
    tasks: state.tasks,
    ideas: state.ideas,
    meetingNotes: state.meetingNotes,
    calendarEvents: state.calendarEvents,
    calendarSettings: state.calendarSettings,
    lastCalendarSyncDate: state.lastCalendarSyncDate || "",
    reminders: state.reminders,
    notificationSettings: state.notificationSettings,
    followUps: state.followUps,
    audioInbox: state.audioInbox,
    selected: state.selected,
    lastCreated: state.lastCreated,
    halMemory: normalizeHalMemory(state.halMemory),
  };
}

function buildBackupEnvelope() {
  return {
    exportedAt: new Date().toISOString(),
    version: 1,
    app: "HAL",
    format: "hal-backup",
    data: buildPersistedState(),
  };
}

function downloadBackupArchive() {
  const payload = buildBackupEnvelope();
  const dateStamp = toLocalDateInputValue(new Date()).replaceAll("-", "");
  downloadBlob(
    new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" }),
    `hal-backup-${dateStamp}.json`,
  );
  setMicStatus("Backup downloaded.", "");
}

async function createLocalBackupSnapshot() {
  try {
    const response = await fetch("/api/backups/local", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildBackupEnvelope()),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not create a local snapshot.");
    }
    elements.backupStatus.textContent = `Local snapshot saved to ${payload.file}.`;
    setMicStatus("Local snapshot created.", "");
    await refreshBackupSnapshots();
  } catch (error) {
    elements.backupStatus.textContent = error.message;
    setMicStatus("HAL could not create a local snapshot.", error.message);
  }
}

function promptRestoreBackup() {
  elements.restoreBackupInput?.click();
}

async function handleRestoreBackupSelection(event) {
  const file = event.target?.files?.[0];
  if (!file) {
    return;
  }

  try {
    const raw = await file.text();
    const parsed = JSON.parse(raw);
    const payload = parsed?.format === "hal-backup" && parsed?.data ? parsed.data : parsed;
    const confirmed = window.confirm("Restore this HAL backup and replace the current local data?");
    if (!confirmed) {
      event.target.value = "";
      return;
    }
    restoreFromBackupPayload(payload);
    elements.backupStatus.textContent = `Restored backup from ${file.name}.`;
    setMicStatus("Backup restored.", "");
    render();
    elements.backupDialog?.close();
  } catch (error) {
    elements.backupStatus.textContent = "HAL could not restore that backup file.";
    setMicStatus("HAL could not restore that backup file.", error.message || "");
  } finally {
    event.target.value = "";
  }
}

function restoreFromBackupPayload(payload) {
  const restored = normalizeLoadedState(payload || {});

  const preservedRuntime = {
    pendingCapture: null,
    pendingVoiceSave: false,
    correctionContext: null,
    micMode: state.micMode,
    speechRecognition: state.speechRecognition,
    mediaRecorder: state.mediaRecorder,
    recording: state.recording,
  };

  Object.keys(state).forEach((key) => {
    delete state[key];
  });
  Object.assign(state, restored, preservedRuntime);

  editingReminderId = null;
  editingTaskId = null;
  editingIdeaId = null;
  editingMeetingNoteId = null;
  editingTaskNoteId = null;
  taskNoteEditorMode = "existing";

  persist();
  syncReminderRecurrenceUI();
  syncReminderEditorState();
  syncTaskTimeField();
  syncQuickTaskTimeField();
  syncTaskEditorState();
  syncTeamsSettingsUI();
  syncSmsSettingsUI();
  syncMyDaySettingsUI();
  renderBackupSummary();
  if (serverStateSync.hydrated) {
    scheduleServerStateSync();
  }
}

async function refreshBackupSnapshots() {
  if (!elements.backupSnapshotList) {
    return;
  }

  try {
    const response = await fetch("/api/backups");
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not read backup snapshots.");
    }
    renderBackupSnapshotList(payload.snapshots || []);
  } catch (error) {
    elements.backupSnapshotList.innerHTML = `<article class="item-card"><p class="card-body">${escapeHtml(error.message)}</p></article>`;
  }
}

function renderBackupSnapshotList(snapshots) {
  if (!elements.backupSnapshotList) {
    return;
  }

  if (!Array.isArray(snapshots) || !snapshots.length) {
    elements.backupSnapshotList.innerHTML = '<article class="item-card"><p class="card-body">No local snapshots yet.</p></article>';
    return;
  }

  elements.backupSnapshotList.innerHTML = snapshots.map((snapshot) => `
    <article class="item-card">
      <div class="card-title-row">
        <h3 class="card-title">${escapeHtml(snapshot.name || "HAL snapshot")}</h3>
        <span class="chip">${escapeHtml(snapshot.modifiedLabel || "")}</span>
      </div>
      <p class="card-body backup-path">${escapeHtml(snapshot.path || "")}</p>
    </article>
  `).join("");
}

async function initializeServerStateMirror() {
  try {
    const response = await fetch("/api/state", {
      credentials: "include",
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "HAL could not reach the local state mirror.");
    }

    const serverData = payload.state || null;
    serverStateSync.lastMeta = payload.meta || null;
    const localHasData = hasMeaningfulHalData(state);
    const serverHasData = hasMeaningfulHalData(serverData);
    const hostedMode = !isLocalHalHost();
    const localScore = calculateMeaningfulHalDataScore(state);
    const serverScore = calculateMeaningfulHalDataScore(serverData);

    if (hostedMode) {
      if (serverHasData) {
        restoreFromBackupPayload(serverData);
        render();
        updateAuthDebugState({
          lastHydration: `hosted: loaded cloud state (${serverScore})`,
        });
        setMicStatus("Loaded HAL data from the cloud.", "");
      } else if (localHasData) {
        updateAuthDebugState({
          lastHydration: `hosted: pushed local state (${localScore})`,
        });
        await saveStateToServer();
      } else {
        updateAuthDebugState({
          lastHydration: "hosted: no local or cloud data",
        });
      }
      return true;
    }

    if (!localHasData && serverHasData) {
      restoreFromBackupPayload(serverData);
      render();
      updateAuthDebugState({
        lastHydration: "local: restored server mirror",
      });
      setMicStatus("Recovered HAL data from the local mirror.", "");
    } else if (localHasData && !serverHasData) {
      updateAuthDebugState({
        lastHydration: "local: pushed local state to mirror",
      });
      await saveStateToServer();
    } else if (localHasData && serverHasData) {
      const localSavedAt = Date.parse(state._meta?.savedAt || 0);
      const serverSavedAt = Date.parse(serverData._meta?.savedAt || 0);
      const serverLooksMoreComplete = serverScore > localScore;
      if (serverLooksMoreComplete || (serverSavedAt > localSavedAt && localSavedAt > 0)) {
        restoreFromBackupPayload(serverData);
        render();
        updateAuthDebugState({
          lastHydration: "local: loaded newer server mirror",
        });
        setMicStatus("Loaded the newer HAL data from the local mirror.", "");
      } else {
        updateAuthDebugState({
          lastHydration: "local: kept local state and synced",
        });
        await saveStateToServer();
      }
    } else {
      updateAuthDebugState({
        lastHydration: "local: no meaningful state detected",
      });
    }
    return true;
  } catch (error) {
    updateAuthDebugState({
      lastHydration: "state mirror failed",
      lastError: error?.message || "state mirror failed",
    });
    setMicStatus("HAL could not load saved data yet.", "Try refreshing once. If this keeps happening, HAL may need to re-check the hosted sign-in cookie.");
    return false;
  } finally {
    serverStateSync.hydrated = true;
  }
}

function scheduleServerStateSync() {
  if (!serverStateSync.hydrated) {
    return;
  }
  window.clearTimeout(serverStateSync.saveTimer);
  serverStateSync.saveTimer = window.setTimeout(() => {
    void saveStateToServer();
  }, 250);
}

async function saveStateToServer() {
  try {
    const response = await fetch("/api/state", {
      method: "PUT",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildPersistedState()),
    });
    if (response.ok) {
      const payload = await response.json();
      serverStateSync.lastMeta = payload.meta || serverStateSync.lastMeta;
      renderBackupSummary();
    }
  } catch {
    // Keep local persistence working even if the local server mirror is unavailable.
  }
}

function hasMeaningfulHalData(source) {
  return calculateMeaningfulHalDataScore(source) > 0;
}

function calculateMeaningfulHalDataScore(source) {
  if (!source || typeof source !== "object") {
    return 0;
  }

  let score = 0;
  score += Array.isArray(source.tasks) ? source.tasks.length * 6 : 0;
  score += Array.isArray(source.ideas) ? source.ideas.length * 5 : 0;
  score += Array.isArray(source.meetingNotes) ? source.meetingNotes.length * 5 : 0;
  score += Array.isArray(source.reminders) ? source.reminders.length * 4 : 0;
  score += Array.isArray(source.calendarEvents) ? source.calendarEvents.length * 2 : 0;
  score += Array.isArray(source.followUps) ? source.followUps.length * 3 : 0;
  score += Array.isArray(source.audioInbox) ? source.audioInbox.length : 0;
  score += Array.isArray(source.quickLinks) && source.quickLinks.length > 4 ? source.quickLinks.length - 4 : 0;
  score += typeof source.quote === "string" && source.quote !== "Build the life and systems you want to live inside." ? 1 : 0;
  return score;
}

function isLocalHalHost() {
  const host = window.location.hostname || "";
  return host === "localhost" || host === "127.0.0.1";
}

function setTaskCompletion(task, isComplete) {
  task.done = isComplete;

  persist();
  render();
}

function toggleTaskHighlight(task) {
  task.highlighted = !task.highlighted;
  persist();
  render();
}

function editMeetingNote(id) {
  const note = state.meetingNotes.find((entry) => entry.id === id);
  if (!note) {
    return;
  }
  editingMeetingNoteId = id;
  elements.meetingNoteEditorTitle.textContent = note.title;
  elements.meetingNoteEditorName.value = note.title;
  elements.meetingNoteEditorDate.value = note.callDate || "";
  elements.meetingNoteEditorContent.innerHTML = normalizeIdeaEditorHtml(note.content);
  elements.deleteMeetingNoteButton.classList.remove("hidden");
  if (!elements.meetingNoteEditorDialog?.open) {
    elements.meetingNoteEditorDialog.showModal();
  }
}

function saveMeetingNoteEditor(event) {
  event.preventDefault();
  const title = elements.meetingNoteEditorName.value.trim();
  const content = sanitizeIdeaEditorHtml(elements.meetingNoteEditorContent.innerHTML);
  const callDate = (elements.meetingNoteEditorDate.value || "").trim();
  if (!title || !content) {
    return;
  }

  const note = state.meetingNotes.find((entry) => entry.id === editingMeetingNoteId);
  if (!note) {
    const createdNote = {
      id: crypto.randomUUID(),
      title,
      person: "",
      callDate,
      content,
      createdAt: new Date().toISOString(),
    };
    state.meetingNotes.unshift(createdNote);
    editingMeetingNoteId = createdNote.id;
    state.lastCreated = { type: "meeting", id: createdNote.id };
    state.selected = { type: "meeting", id: createdNote.id };
    elements.meetingNoteEditorTitle.textContent = createdNote.title;
    elements.deleteMeetingNoteButton.classList.remove("hidden");
    persist();
    render();
    elements.meetingNoteEditorDialog?.close();
    openMeetingNotesDialog();
    setMicStatus("Meeting note created.", "");
    return;
  }

  rememberCorrection({
    entityType: "meeting",
    entityId: note.id,
    action: "edit",
    from: {
      title: note.title,
      content: note.content,
      callDate: note.callDate,
    },
    to: {
      title,
      content,
      callDate,
    },
  });
  note.title = title;
  note.content = content;
  note.callDate = callDate;
  persist();
  render();
  elements.meetingNoteEditorDialog?.close();
  openMeetingNotesDialog();
  setMicStatus("Meeting note updated.", "");
}

function deleteMeetingNoteFromEditor() {
  const note = state.meetingNotes.find((entry) => entry.id === editingMeetingNoteId);
  if (!note) {
    return;
  }
  const confirmed = window.confirm(`Delete "${note.title}"?`);
  if (!confirmed) {
    return;
  }
  state.meetingNotes = state.meetingNotes.filter((entry) => entry.id !== note.id);
  editingMeetingNoteId = null;
  elements.meetingNoteEditorDialog?.close();
  persist();
  render();
}

function handleMeetingNoteEditorToolbarClick(event) {
  const button = event.target.closest("[data-editor-command]");
  if (!button) {
    return;
  }
  event.preventDefault();
  focusMeetingNoteEditor();
  if (button.dataset.editorCommand === "removeFormat") {
    document.execCommand("removeFormat", false);
    document.execCommand("unlink", false);
    return;
  }
  document.execCommand(button.dataset.editorCommand, false);
}

function handleMeetingNoteEditorFontButtonClick(event) {
  event.preventDefault();
  const nextFont = window.prompt("Choose a font: IBM Plex Sans, Georgia, Arial, or Courier New", "IBM Plex Sans");
  if (!nextFont) {
    return;
  }
  focusMeetingNoteEditor();
  document.execCommand("fontName", false, nextFont);
}

function handleMeetingNoteEditorSizeButtonClick(event) {
  event.preventDefault();
  const nextSize = window.prompt("Choose a size: 2 small, 3 normal, 4 large, or 5 XL", "3");
  if (!nextSize) {
    return;
  }
  focusMeetingNoteEditor();
  document.execCommand("fontSize", false, nextSize);
}

function focusMeetingNoteEditor() {
  elements.meetingNoteEditorContent?.focus();
}

function preserveMeetingNoteEditorSelection(event) {
  if (!event.target.closest("button")) {
    return;
  }
  event.preventDefault();
}

function setupSpeechRecognition() {
  const SpeechRecognition = window.webkitSpeechRecognition || window.SpeechRecognition;
  const diagnostics = getMicDiagnostics();
  const canRecordAudio = diagnostics.canRecordAudio;
  console.log("HAL microphone diagnostics:", diagnostics);

  if (!SpeechRecognition) {
    state.micMode = canRecordAudio ? "audio" : "none";
    elements.micStatus.textContent = canRecordAudio
      ? "Live speech-to-text is unavailable. HAL can still record audio clips."
      : "Microphone features are unavailable in this browser.";
    elements.micSupportNote.textContent = canRecordAudio
      ? "Chrome is allowing microphone access, but the browser speech recognition API is not exposed here. HAL will use audio capture fallback until we wire in a dedicated transcription service."
      : "This environment is not exposing browser speech recognition or audio capture APIs to HAL.";
    return;
  }

  try {
    const recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = "en-US";
    state.micMode = "speech";
    elements.micSupportNote.textContent = "";

    recognition.onstart = () => {
      elements.toggleRecording.classList.add("is-recording");
      elements.toggleRecording.setAttribute("aria-pressed", "true");
      setMicStatus("HAL is listening...", elements.micSupportNote.textContent);
    };

    recognition.onresult = (event) => {
      const transcript = Array.from(event.results).map((result) => result[0].transcript).join(" ");
      const targetId = state.recording?.target || "content";
      const field = document.getElementById(targetId);
      if (!field) {
        return;
      }
      const nextValue = transcript.trim();
      if (targetId === "content") {
        const saveIntent = detectVoiceSaveCommand(nextValue);
        field.value = saveIntent.cleanedText;
        previewCapture();
        if (saveIntent.shouldSave && !state.pendingVoiceSave) {
          state.pendingVoiceSave = true;
          setMicStatus("Saving...", "");
          state.speechRecognition.stop();
          return;
        }
      } else {
        field.value = nextValue;
      }
      if (targetId === "content") {
        previewCapture();
      } else if (field.type === "search") {
        field.dispatchEvent(new Event("input"));
      }
    };

    recognition.onerror = (event) => {
      state.recording = null;
      elements.toggleRecording.classList.remove("is-recording");
      elements.toggleRecording.setAttribute("aria-pressed", "false");
      setMicStatus(`Speech recognition error: ${event.error}`, `Chrome exposed speech recognition, but it failed with "${event.error}". HAL can switch to audio recording fallback if needed.`);
      console.error("HAL speech recognition error:", event.error);
    };

    recognition.onend = () => {
      state.recording = null;
      elements.toggleRecording.classList.remove("is-recording");
      elements.toggleRecording.setAttribute("aria-pressed", "false");
      if (state.pendingVoiceSave) {
        state.pendingVoiceSave = false;
        void submitCaptureFromVoice();
        return;
      }
      if (!elements.micStatus.textContent.startsWith("Speech recognition error")) {
        setMicStatus("Tap the microphone or type below.", elements.micSupportNote.textContent);
      }
    };

    state.speechRecognition = recognition;
  } catch (error) {
    state.speechRecognition = null;
    state.micMode = canRecordAudio ? "audio" : "none";
    setMicStatus("Speech recognition could not be initialized.", `Speech recognition constructor failed: ${error.message}`);
    console.error("HAL speech recognition constructor failed:", error);
  }
}

function toggleRecording() {
  if (state.micMode === "audio") {
    toggleAudioFallbackRecording();
    return;
  }
  if (!state.speechRecognition) {
    window.alert(`Speech recognition is not available here, and no fallback recorder is ready.\n\n${getMicDiagnostics().summary}`);
    return;
  }
  if (state.recording?.target === "content") {
    state.pendingVoiceSave = false;
    state.speechRecognition.stop();
    return;
  }
  state.recording = { target: "content" };
  try {
    state.speechRecognition.start();
  } catch (error) {
    state.recording = null;
    setMicStatus(`Speech recognition start failed: ${error.message}`, "HAL can fall back to audio recording while we stabilize live transcription.");
    if (getMicDiagnostics().canRecordAudio) {
      state.micMode = "audio";
    }
    console.error("HAL speech recognition start failed:", error);
  }
}

function detectVoiceSaveCommand(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) {
    return { shouldSave: false, cleanedText: "" };
  }

  const patterns = [
    /\b(?:hal[\s,]*)?save it[.!?]*$/i,
    /\b(?:hal[\s,]*)?save this[.!?]*$/i,
    /\b(?:hal[\s,]*)?that's it[.!?]*$/i,
    /\b(?:hal[\s,]*)?that is it[.!?]*$/i,
    /\b(?:hal[\s,]*)?done[.!?]*$/i,
  ];

  for (const pattern of patterns) {
    if (pattern.test(normalized)) {
      const cleanedText = normalized.replace(pattern, "").trim().replace(/[,\s]+$/, "");
      return {
        shouldSave: Boolean(cleanedText),
        cleanedText,
      };
    }
  }

  return {
    shouldSave: false,
    cleanedText: normalized,
  };
}

function startFieldDictation(targetId) {
  if (!state.speechRecognition) {
    window.alert("Voice search needs live speech recognition, which this browser session is not exposing right now.");
    return;
  }
  state.recording = { target: targetId };
  try {
    state.speechRecognition.start();
  } catch (error) {
    state.recording = null;
    window.alert(`Voice search could not start: ${error.message}`);
    console.error("HAL voice search start failed:", error);
  }
}

async function toggleAudioFallbackRecording() {
  if (state.mediaRecorder && state.mediaRecorder.state === "recording") {
    state.mediaRecorder.stop();
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const recorder = new MediaRecorder(stream);
    const chunks = [];

    recorder.onstart = () => {
      state.recording = { target: "audio-fallback" };
      elements.toggleRecording.classList.add("is-recording");
      elements.toggleRecording.setAttribute("aria-pressed", "true");
      setMicStatus("HAL is recording audio...", elements.micSupportNote.textContent);
    };

    recorder.ondataavailable = (event) => {
      if (event.data.size > 0) {
        chunks.push(event.data);
      }
    };

    recorder.onstop = () => {
      const blob = new Blob(chunks, { type: recorder.mimeType || "audio/webm" });
      const audioUrl = URL.createObjectURL(blob);
      state.audioInbox.unshift({
        id: crypto.randomUUID(),
        title: `Audio note ${state.audioInbox.length + 1}`,
        status: "Transcription pending",
        createdAt: new Date().toISOString(),
        audioUrl,
      });
      state.recording = null;
      state.mediaRecorder = null;
      elements.toggleRecording.classList.remove("is-recording");
      elements.toggleRecording.setAttribute("aria-pressed", "false");
      setMicStatus("Audio note captured.", "You can play it back or use it as a placeholder in capture.");
      stream.getTracks().forEach((track) => track.stop());
      persist();
    };

    recorder.onerror = () => {
      setMicStatus("Audio recording failed.", elements.micSupportNote.textContent);
      state.mediaRecorder = null;
      state.recording = null;
      stream.getTracks().forEach((track) => track.stop());
    };

    state.mediaRecorder = recorder;
    recorder.start();
  } catch {
    window.alert(`HAL could not access the microphone recorder.\n\n${getMicDiagnostics().summary}`);
  }
}

function setMicStatus(statusText, supportText) {
  if (elements.micStatus) {
    elements.micStatus.textContent = statusText;
  }
  if (elements.micSupportNote && supportText !== undefined) {
    elements.micSupportNote.textContent = supportText;
  }
}

function initializeTextToSpeech() {
  if (!window.speechSynthesis) {
    return;
  }

  cacheHalSpeechVoice();
  window.speechSynthesis.addEventListener("voiceschanged", cacheHalSpeechVoice);
}

function cacheHalSpeechVoice() {
  const voices = window.speechSynthesis?.getVoices?.() || [];
  halSpeechVoice = selectPreferredHalVoice(voices);
}

function selectPreferredHalVoice(voices) {
  if (!Array.isArray(voices) || !voices.length) {
    return null;
  }

  const preferredNames = [
    "microsoft david",
    "microsoft guy",
    "microsoft mark",
    "google us english",
    "daniel",
    "matthew",
    "james",
    "ryan",
    "alex",
    "guy",
    "david",
    "mark",
  ];

  const scored = voices
    .filter((voice) => /^en(-|_|$)/i.test(voice.lang || ""))
    .map((voice) => ({
      voice,
      score: getHalVoiceScore(voice, preferredNames),
    }))
    .sort((left, right) => right.score - left.score);

  return scored[0]?.voice || null;
}

function getHalVoiceScore(voice, preferredNames) {
  const name = String(voice.name || "").toLowerCase();
  const lang = String(voice.lang || "").toLowerCase();
  let score = 0;

  if (lang.startsWith("en-us")) {
    score += 6;
  } else if (lang.startsWith("en")) {
    score += 4;
  }

  preferredNames.forEach((candidate, index) => {
    if (name.includes(candidate)) {
      score += 30 - index;
    }
  });

  if (voice.default) {
    score += 3;
  }

  if (/female|woman|zira|aria|jenny|sara|sonia|libby|ava/i.test(name)) {
    score -= 12;
  }

  return score;
}

function speakHal(text) {
  if (!state.voiceResponsesEnabled || !text || !window.speechSynthesis) {
    return;
  }
  try {
    window.speechSynthesis.cancel();
    if (!halSpeechVoice) {
      cacheHalSpeechVoice();
    }
    const utterance = new SpeechSynthesisUtterance(text);
    if (halSpeechVoice) {
      utterance.voice = halSpeechVoice;
      utterance.lang = halSpeechVoice.lang;
    }
    utterance.rate = 0.9;
    utterance.pitch = 0.82;
    utterance.volume = 0.92;
    window.speechSynthesis.speak(utterance);
  } catch {}
}

function rememberPrompt(text, interpretation) {
  const history = state.halMemory.promptHistory;
  history.unshift({
    id: crypto.randomUUID(),
    text,
    interpretedAs: interpretation.type,
    createdAt: new Date().toISOString(),
  });
  state.halMemory.promptHistory = history.slice(0, 50);
}

function rememberCorrection(entry) {
  const corrections = state.halMemory.corrections;
  corrections.unshift({
    id: crypto.randomUUID(),
    ...entry,
    createdAt: new Date().toISOString(),
  });
  state.halMemory.corrections = corrections.slice(0, 100);
}

function learnTitleCorrection(kind, originalTitle, updatedTitle) {
  const original = normalizeForDuplicateCheck(originalTitle || "");
  const updated = String(updatedTitle || "").trim();
  if (!original || !updated || normalizeForDuplicateCheck(updated) === original) {
    return;
  }

  const target = kind === "task"
    ? state.halMemory.preferences.taskTitleCorrections
    : state.halMemory.preferences.ideaTitleCorrections;
  const existing = target.find((entry) => entry.original === original);
  if (existing) {
    existing.updated = updated;
    existing.updatedAt = new Date().toISOString();
    return;
  }

  target.unshift({
    original,
    updated,
    updatedAt: new Date().toISOString(),
  });
}

function learnIntentCorrection(sourcePrompt, fromType, toType) {
  const normalizedPrompt = normalizeIntentPrompt(sourcePrompt);
  if (!normalizedPrompt || !fromType || !toType || fromType === toType) {
    return;
  }

  const target = state.halMemory.preferences.intentCorrections;
  const existing = target.find((entry) => entry.normalizedPrompt === normalizedPrompt && entry.fromType === fromType);
  if (existing) {
    existing.toType = toType;
    existing.updatedAt = new Date().toISOString();
    return;
  }

  target.unshift({
    normalizedPrompt,
    fromType,
    toType,
    updatedAt: new Date().toISOString(),
  });
}

function learnRoutingPattern({
  sourcePrompt,
  fromType = "",
  toType = "",
  listId = "",
  meetingSignature = "",
  correctionText = "",
}) {
  const normalizedPrompt = normalizeIntentPrompt(sourcePrompt);
  if (!normalizedPrompt) {
    return;
  }

  const target = state.halMemory.preferences.routingPatterns;
  const keywords = extractLearningKeywords(sourcePrompt);
  const existing = target.find((entry) => (
    entry.normalizedPrompt === normalizedPrompt
    && entry.fromType === String(fromType || "")
  ));

  if (existing) {
    existing.toType = String(toType || existing.toType || "");
    existing.listId = String(listId || existing.listId || "");
    existing.meetingSignature = normalizeMeetingTargetSignature(meetingSignature || existing.meetingSignature || "");
    existing.keywords = keywords.length ? keywords : existing.keywords;
    existing.lastCorrectionText = String(correctionText || existing.lastCorrectionText || "");
    existing.count = Number(existing.count || 0) + 1;
    existing.updatedAt = new Date().toISOString();
    return;
  }

  target.unshift({
    normalizedPrompt,
    fromType: String(fromType || ""),
    toType: String(toType || ""),
    listId: String(listId || ""),
    meetingSignature: normalizeMeetingTargetSignature(meetingSignature || ""),
    keywords,
    lastCorrectionText: String(correctionText || ""),
    count: 1,
    updatedAt: new Date().toISOString(),
  });
}

function learnTaskListCorrection(sourcePrompt, listId) {
  const normalizedPrompt = normalizeIntentPrompt(sourcePrompt);
  if (!normalizedPrompt || !listId) {
    return;
  }

  const target = state.halMemory.preferences.taskListCorrections;
  const existing = target.find((entry) => entry.normalizedPrompt === normalizedPrompt);
  if (existing) {
    existing.listId = listId;
    existing.updatedAt = new Date().toISOString();
    return;
  }

  target.unshift({
    normalizedPrompt,
    listId,
    updatedAt: new Date().toISOString(),
  });

  learnRoutingPattern({
    sourcePrompt,
    fromType: "task",
    toType: "task",
    listId,
  });
}

function learnMeetingTargetCorrection(sourcePrompt, note) {
  const normalizedPrompt = normalizeIntentPrompt(sourcePrompt);
  if (!normalizedPrompt || !note) {
    return;
  }

  const target = state.halMemory.preferences.meetingTargetCorrections;
  const signature = buildMeetingTargetSignature(note);
  const existing = target.find((entry) => entry.normalizedPrompt === normalizedPrompt);
  if (existing) {
    existing.signature = signature;
    existing.updatedAt = new Date().toISOString();
    return;
  }

  target.unshift({
    normalizedPrompt,
    signature,
    updatedAt: new Date().toISOString(),
  });

  learnRoutingPattern({
    sourcePrompt,
    fromType: "meeting",
    toType: "meeting",
    meetingSignature: signature,
  });
}

function applyLearnedTitleCorrection(kind, title) {
  const normalized = normalizeForDuplicateCheck(title || "");
  if (!normalized) {
    return title;
  }
  const source = kind === "task"
    ? state.halMemory.preferences.taskTitleCorrections
    : state.halMemory.preferences.ideaTitleCorrections;
  const match = source.find((entry) => entry.original === normalized);
  return match?.updated || title;
}

function findLearnedTaskListCorrection(text) {
  const normalizedPrompt = normalizeIntentPrompt(text);
  if (!normalizedPrompt) {
    return null;
  }

  const routingPatternMatch = findBestRoutingPattern(normalizedPrompt, "task", {
    requireList: true,
  });
  if (routingPatternMatch?.listId) {
    return state.taskLists.find((list) => list.id === routingPatternMatch.listId) || null;
  }

  const bestMatch = state.halMemory.preferences.taskListCorrections
    .map((entry) => ({
      entry,
      score: scoreIntentCorrection(normalizedPrompt, entry.normalizedPrompt),
    }))
    .filter(({ score }) => score >= 0.78)
    .sort((left, right) => right.score - left.score)[0];

  if (!bestMatch?.entry?.listId) {
    return null;
  }
  return state.taskLists.find((list) => list.id === bestMatch.entry.listId) || null;
}

function findLearnedMeetingTargetCorrection(text) {
  const normalizedPrompt = normalizeIntentPrompt(text);
  if (!normalizedPrompt) {
    return null;
  }

  const routingPatternMatch = findBestRoutingPattern(normalizedPrompt, "meeting", {
    requireMeetingSignature: true,
  });
  if (routingPatternMatch?.meetingSignature) {
    return findMeetingNoteBySignature(routingPatternMatch.meetingSignature);
  }

  const bestMatch = state.halMemory.preferences.meetingTargetCorrections
    .map((entry) => ({
      entry,
      score: scoreIntentCorrection(normalizedPrompt, entry.normalizedPrompt),
    }))
    .filter(({ score }) => score >= 0.74)
    .sort((left, right) => right.score - left.score)[0];

  if (!bestMatch?.entry?.signature) {
    return null;
  }

  return findMeetingNoteBySignature(bestMatch.entry.signature);
}

function applyLearnedIntentCorrection(text, interpretation) {
  const normalizedPrompt = normalizeIntentPrompt(text);
  if (!normalizedPrompt || !interpretation?.type) {
    return interpretation;
  }

  const routingPatternMatch = findBestRoutingPattern(normalizedPrompt, interpretation.type, {
    requireToType: true,
  });
  if (routingPatternMatch?.toType && routingPatternMatch.toType !== interpretation.type) {
    return rebuildInterpretationFromLearnedType(text, routingPatternMatch.toType, routingPatternMatch);
  }

  const bestMatch = state.halMemory.preferences.intentCorrections
    .map((entry) => ({
      entry,
      score: scoreIntentCorrection(normalizedPrompt, entry.normalizedPrompt),
    }))
    .filter(({ entry, score }) => entry.fromType === interpretation.type && score >= 0.68)
    .sort((left, right) => right.score - left.score)[0];

  if (!bestMatch || bestMatch.entry.toType === interpretation.type) {
    return interpretation;
  }

  if (bestMatch.entry.toType === "task") {
    return buildTaskInterpretation(text, "");
  }
  if (bestMatch.entry.toType === "meeting") {
    return buildMeetingInterpretation(text, "");
  }
  if (bestMatch.entry.toType === "idea") {
    return buildIdeaInterpretation(text, "");
  }
  return interpretation;
}

function rebuildInterpretationFromLearnedType(text, targetType, routingPattern = null) {
  if (targetType === "task") {
    return buildTaskInterpretation(text, "", routingPattern?.listId || "");
  }
  if (targetType === "meeting") {
    if (routingPattern?.meetingSignature && isLikelyMeetingUpdatePrompt(text)) {
      const targetNote = findMeetingNoteBySignature(routingPattern.meetingSignature);
      if (targetNote) {
        return buildMeetingUpdateInterpretation(text, "", targetNote);
      }
    }
    return buildMeetingInterpretation(text, "");
  }
  if (targetType === "idea") {
    return buildIdeaInterpretation(text, "");
  }
  return interpretCapture(text, "");
}

function applyLearnedRoutingPreferences(text, interpretation) {
  if (!interpretation) {
    return interpretation;
  }

  if (interpretation.type === "task") {
    const learnedList = findLearnedTaskListCorrection(text);
    if (learnedList) {
      const adjusted = { ...interpretation, listId: learnedList.id };
      adjusted.preview = `HAL will create a task in "${learnedList.name}"${adjusted.dueDate ? ` due ${formatDisplayDate(adjusted.dueDate)}` : ""}.`;
      return adjusted;
    }
  }

  if (interpretation.type === "meeting") {
    const learnedTarget = findLearnedMeetingTargetCorrection(text);
    if (learnedTarget && isLikelyMeetingUpdatePrompt(text)) {
      return buildMeetingUpdateInterpretation(text, "", learnedTarget);
    }
  }

  return interpretation;
}

function findBestRoutingPattern(textOrNormalizedPrompt, interpretationType = "", options = {}) {
  const normalizedPrompt = normalizeIntentPrompt(textOrNormalizedPrompt);
  if (!normalizedPrompt) {
    return null;
  }

  const bestMatch = state.halMemory.preferences.routingPatterns
    .map((entry) => ({
      entry,
      score: scoreRoutingPattern(normalizedPrompt, interpretationType, entry),
    }))
    .filter(({ entry, score }) => {
      if (options.requireToType && !entry.toType) {
        return false;
      }
      if (options.requireList && !entry.listId) {
        return false;
      }
      if (options.requireMeetingSignature && !entry.meetingSignature) {
        return false;
      }
      return score >= 0.72;
    })
    .sort((left, right) => right.score - left.score)[0];

  return bestMatch?.entry || null;
}

function scoreRoutingPattern(normalizedPrompt, interpretationType, entry) {
  const promptScore = scoreIntentCorrection(normalizedPrompt, entry.normalizedPrompt);
  const promptTokens = tokenizeIntentPrompt(normalizedPrompt);
  const keywordScore = scoreKeywordOverlap(promptTokens, entry.keywords || []);
  const typeBoost = entry.fromType && interpretationType && entry.fromType === interpretationType ? 0.18 : 0;
  const confidenceBoost = Math.min(Number(entry.count || 1), 5) * 0.04;
  return promptScore + keywordScore + typeBoost + confidenceBoost;
}

function extractLearningKeywords(value) {
  return normalizeLearningKeywords(tokenizeIntentPrompt(value).slice(0, 8));
}

function normalizeLearningKeywords(items) {
  return Array.from(new Set(
    (Array.isArray(items) ? items : [])
      .map((item) => String(item || "").trim().toLowerCase())
      .filter(Boolean),
  )).slice(0, 8);
}

function scoreKeywordOverlap(sourceTokens, targetKeywords) {
  if (!sourceTokens.length || !targetKeywords.length) {
    return 0;
  }
  const sourceSet = new Set(sourceTokens);
  const shared = targetKeywords.filter((token) => sourceSet.has(token)).length;
  return shared / Math.max(sourceTokens.length, targetKeywords.length);
}

function createDefaultHalMemory() {
  return {
    promptHistory: [],
    corrections: [],
    preferences: {
      taskTitleCorrections: [],
      ideaTitleCorrections: [],
      intentCorrections: [],
      taskListCorrections: [],
      meetingTargetCorrections: [],
      routingPatterns: [],
    },
  };
}

function normalizeCorrectionList(items, limit) {
  return Array.isArray(items)
    ? items.slice(0, limit).filter(Boolean)
    : [];
}

function normalizeTitleCorrections(items) {
  return Array.isArray(items)
    ? items
      .filter((item) => item && item.original && item.updated)
      .map((item) => ({
        original: normalizeForDuplicateCheck(item.original),
        updated: String(item.updated).trim(),
        updatedAt: item.updatedAt || new Date().toISOString(),
      }))
      .filter((item) => item.original && item.updated)
    : [];
}

function normalizeIntentCorrections(items) {
  return Array.isArray(items)
    ? items
      .filter((item) => item && item.normalizedPrompt && item.fromType && item.toType)
      .map((item) => ({
        normalizedPrompt: normalizeIntentPrompt(item.normalizedPrompt),
        fromType: String(item.fromType),
        toType: String(item.toType),
        updatedAt: item.updatedAt || new Date().toISOString(),
      }))
      .filter((item) => item.normalizedPrompt && item.fromType && item.toType)
    : [];
}

function normalizeTaskListCorrections(items) {
  return Array.isArray(items)
    ? items
      .filter((item) => item && item.normalizedPrompt && item.listId)
      .map((item) => ({
        normalizedPrompt: normalizeIntentPrompt(item.normalizedPrompt),
        listId: String(item.listId),
        updatedAt: item.updatedAt || new Date().toISOString(),
      }))
      .filter((item) => item.normalizedPrompt && item.listId)
    : [];
}

function normalizeMeetingTargetCorrections(items) {
  return Array.isArray(items)
    ? items
      .filter((item) => item && item.normalizedPrompt && item.signature)
      .map((item) => ({
        normalizedPrompt: normalizeIntentPrompt(item.normalizedPrompt),
        signature: normalizeMeetingTargetSignature(item.signature),
        updatedAt: item.updatedAt || new Date().toISOString(),
      }))
      .filter((item) => item.normalizedPrompt && item.signature)
    : [];
}

function normalizeRoutingPatterns(items) {
  return Array.isArray(items)
    ? items
      .filter((item) => item && item.normalizedPrompt)
      .map((item) => ({
        normalizedPrompt: normalizeIntentPrompt(item.normalizedPrompt),
        fromType: String(item.fromType || ""),
        toType: String(item.toType || ""),
        listId: String(item.listId || ""),
        meetingSignature: normalizeMeetingTargetSignature(item.meetingSignature || ""),
        keywords: normalizeLearningKeywords(item.keywords),
        lastCorrectionText: String(item.lastCorrectionText || ""),
        count: Math.max(1, Number(item.count || 1)),
        updatedAt: item.updatedAt || new Date().toISOString(),
      }))
      .filter((item) => item.normalizedPrompt)
    : [];
}

function normalizeHalMemory(memory) {
  const defaults = createDefaultHalMemory();
  return {
    promptHistory: normalizeCorrectionList(memory?.promptHistory, 50),
    corrections: normalizeCorrectionList(memory?.corrections, 100),
    preferences: {
      ...defaults.preferences,
      taskTitleCorrections: normalizeTitleCorrections(memory?.preferences?.taskTitleCorrections),
      ideaTitleCorrections: normalizeTitleCorrections(memory?.preferences?.ideaTitleCorrections),
      intentCorrections: normalizeIntentCorrections(memory?.preferences?.intentCorrections),
      taskListCorrections: normalizeTaskListCorrections(memory?.preferences?.taskListCorrections),
      meetingTargetCorrections: normalizeMeetingTargetCorrections(memory?.preferences?.meetingTargetCorrections),
      routingPatterns: normalizeRoutingPatterns(memory?.preferences?.routingPatterns),
    },
  };
}

async function interpretCaptureWithLearning(text, manualTitle) {
  const interpretation = interpretCapture(text, manualTitle);
  const adjusted = applyLearnedRoutingPreferences(text, applyLearnedIntentCorrection(text, { ...interpretation }));

  if (adjusted.type === "task" && adjusted.title) {
    adjusted.title = applyLearnedTitleCorrection("task", adjusted.title);
  } else if (adjusted.type === "idea" && adjusted.title) {
    adjusted.title = applyLearnedTitleCorrection("idea", adjusted.title);
  }

  if (shouldBypassAiInterpretation(adjusted)) {
    return adjusted;
  }

  const aiInterpretation = await requestAiInterpretation(text, adjusted);
  return mergeAiInterpretation(text, adjusted, aiInterpretation);
}

function shouldBypassAiInterpretation(interpretation) {
  if (interpretation?.forceLocal) {
    return true;
  }
  return [
    "follow-up",
    "meeting-update",
    "reminder-delete",
    "reminder-delete-missing",
    "reminder-edit",
    "reminder-edit-missing",
  ].includes(interpretation.type);
}

async function requestAiInterpretation(text, localInterpretation) {
  try {
    const response = await fetch("/api/ai/interpret", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        text,
        today: todayString(),
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        taskLists: state.taskLists.map((list) => ({ id: list.id, name: list.name })),
        localInterpretation: {
          type: localInterpretation.type,
          title: localInterpretation.title || "",
          label: localInterpretation.label || "",
          preview: localInterpretation.preview || "",
          dueDate: localInterpretation.dueDate || "",
          dueTime: localInterpretation.dueTime || "",
          person: localInterpretation.person || "",
        },
        memory: {
          recentPrompts: state.halMemory.promptHistory.slice(0, 8),
          corrections: state.halMemory.corrections.slice(0, 12),
          preferences: state.halMemory.preferences,
        },
      }),
    });

    if (!response.ok) {
      return null;
    }

    const payload = await response.json();
    if (!payload?.enabled || !payload?.interpretation) {
      return null;
    }

    return payload.interpretation;
  } catch {
    return null;
  }
}

function mergeAiInterpretation(text, localInterpretation, aiInterpretation) {
  if (!aiInterpretation || aiInterpretation.intentType === "use_local") {
    return localInterpretation;
  }

  const titleOverride = String(aiInterpretation.title || "").trim();
  const formattedOverride = String(aiInterpretation.formatted || "").trim();
  const personOverride = String(aiInterpretation.person || "").trim();

  if (aiInterpretation.intentType === "task") {
    return buildTaskInterpretation(text, titleOverride || localInterpretation.title, localInterpretation.listId);
  }

  if (aiInterpretation.intentType === "meeting") {
    return buildMeetingInterpretation(text, titleOverride || localInterpretation.title, formattedOverride, personOverride);
  }

  if (aiInterpretation.intentType === "reminder") {
    return buildReminderInterpretation(text, titleOverride || localInterpretation.title);
  }

  if (aiInterpretation.intentType === "idea") {
    return buildIdeaInterpretation(text, titleOverride || localInterpretation.title, formattedOverride);
  }

  return localInterpretation;
}

function getMicDiagnostics() {
  const hasSpeechRecognition = Boolean(window.SpeechRecognition || window.webkitSpeechRecognition);
  const hasMediaDevices = Boolean(navigator.mediaDevices);
  const hasGetUserMedia = Boolean(navigator.mediaDevices?.getUserMedia);
  const hasMediaRecorder = Boolean(window.MediaRecorder);
  const isSecureContextAvailable = window.isSecureContext;
  const protocol = window.location.protocol;
  const host = window.location.host;

  return {
    hasSpeechRecognition,
    hasMediaDevices,
    hasGetUserMedia,
    hasMediaRecorder,
    isSecureContextAvailable,
    canRecordAudio: hasMediaDevices && hasGetUserMedia && hasMediaRecorder && isSecureContextAvailable,
    summary: `speech=${hasSpeechRecognition}, mediaDevices=${hasMediaDevices}, getUserMedia=${hasGetUserMedia}, mediaRecorder=${hasMediaRecorder}, secureContext=${isSecureContextAvailable}, origin=${protocol}//${host}`,
  };
}

function interpretCapture(text, manualTitle) {
  const explicitTitle = manualTitle || extractExplicitTitle(text);
  const normalizedText = explicitTitle ? stripExplicitTitleInstruction(text) : text;
  const lower = normalizedText.toLowerCase();
  const targetTaskList = findTaskListFromText(normalizedText);
  const meetingUpdateTarget = findMatchingMeetingNote(normalizedText);

  if (isReminderDeleteIntent(lower)) {
    const reminderTarget = findMatchingReminder(normalizedText);
    if (reminderTarget) {
      return {
        type: "reminder-delete",
        label: "Reminder",
        target: reminderTarget,
        preview: `HAL will delete the reminder "${reminderTarget.text}".`,
      };
    }
    return {
      type: "reminder-delete-missing",
      label: "Reminder",
      title: "",
      preview: "HAL could not find a matching reminder to delete.",
    };
  }

  if (isReminderEditIntent(lower)) {
    const reminderTarget = findMatchingReminder(normalizedText);
    if (reminderTarget) {
      return {
        type: "reminder-edit",
        label: "Reminder",
        target: reminderTarget,
        preview: `HAL will open the reminder "${reminderTarget.text}" for editing.`,
      };
    }
    return {
      type: "reminder-edit-missing",
      label: "Reminder",
      title: "",
      preview: "HAL could not find a matching reminder to edit.",
    };
  }

  if (isReminderIntent(lower)) {
    return buildReminderInterpretation(normalizedText, explicitTitle);
  }

  if (isMeetingUpdateIntent(lower, meetingUpdateTarget)) {
    return buildMeetingUpdateInterpretation(normalizedText, explicitTitle, meetingUpdateTarget);
  }

  if (isTaskIntent(lower, targetTaskList)) {
    return buildTaskInterpretation(normalizedText, explicitTitle, targetTaskList?.id);
  }

  if (isMeetingIntent(lower)) {
    return buildMeetingInterpretation(normalizedText, explicitTitle);
  }

  return buildIdeaInterpretation(normalizedText, explicitTitle);
}

function buildTaskInterpretation(text, manualTitle, listIdOverride = "") {
  const lower = text.toLowerCase();
  const dueDate = parseDueDate(lower);
  const dueTime = parseTaskTimeDetails(lower);
  const targetTaskList = listIdOverride
    ? state.taskLists.find((list) => list.id === listIdOverride)
    : findTaskListFromText(text);
  const listId = targetTaskList?.id || state.taskLists[0].id;
  const title = manualTitle || cleanTaskTitle(text);

  return {
    type: "task",
    label: "Task",
    title,
    dueDate,
    dueTime: dueTime.explicit ? dueTime.time : "",
    listId,
    forceLocal: Boolean(targetTaskList || hasExplicitTaskLanguage(lower)),
    preview: `HAL will create a task in "${getTaskListName(listId)}"${dueDate ? ` due ${formatDisplayDate(dueDate)}` : ""}.`,
  };
}

function buildReminderInterpretation(text, manualTitle) {
  const lower = text.toLowerCase();
  const recurrence = parseReminderRecurrence(lower);
  const title = manualTitle || cleanReminderText(text);
  const reminderTime = parseReminderTimeDetails(lower);
  const reminderDate = parseDueDate(lower) || todayString();
  const delivery = inferReminderDelivery(text);

  if (!reminderTime.explicit) {
    return {
      type: "follow-up",
      question: "What time would you like to be reminded?",
      supportText: "You can say something like 3:15 PM, 8 AM, or this afternoon.",
      placeholder: 'Example: "3:15 PM"',
      pendingCapture: {
        kind: "reminder-time",
        draft: {
          title,
          date: reminderDate,
          delivery,
          frequency: recurrence ? "recurring" : "once",
          recurrence,
        },
      },
    };
  }

  return {
    type: "reminder",
    label: "Reminder",
    title,
    at: `${reminderDate}T${reminderTime.time}`,
    delivery,
    frequency: recurrence ? "recurring" : "once",
    recurrence,
    preview: `HAL will create a reminder${recurrence ? ` (${getReminderScheduleLabel({ frequency: "recurring", recurrence })})` : ""}.`,
  };
}

function buildMeetingInterpretation(text, manualTitle, formattedOverride = "", personOverride = "") {
  const lower = text.toLowerCase();
  const dueDate = parseDueDate(lower);
  const person = personOverride || extractPerson(lower);
  const title = manualTitle || `Call with ${person || "Contact"}`;
  const formatted = formattedOverride || formatMeetingNote(text);

  return {
    type: "meeting",
    label: "Meeting Notes",
    title,
    dueDate,
    person: person || "",
    formatted,
    forceLocal: hasExplicitMeetingLanguage(lower),
    preview: `HAL will create a meeting note${dueDate ? ` for ${formatDisplayDate(dueDate)}` : ""} using talking-point formatting.`,
  };
}

function buildMeetingUpdateInterpretation(text, manualTitle, targetNote) {
  const content = manualTitle
    ? ""
    : cleanMeetingUpdateText(text);
  const formatted = formatMeetingNote(content);

  return {
    type: "meeting-update",
    label: "Meeting Notes",
    target: targetNote,
    title: targetNote?.title || "",
    formatted,
    rawAppendText: content,
    forceLocal: true,
    preview: targetNote
      ? `HAL will add this to "${targetNote.title}".`
      : "HAL will add this to the matching meeting note.",
  };
}

function buildIdeaInterpretation(text, manualTitle, formattedOverride = "") {
  const lower = text.toLowerCase();
  const dueDate = parseDueDate(lower);
  const title = manualTitle || generateIdeaTitle(text);
  const formatted = formattedOverride || formatIdea(text);

  return {
    type: "idea",
    label: "Idea",
    title,
    dueDate,
    formatted,
    preview: "HAL will store this in Ideas with an auto-generated title you can edit later.",
  };
}

function detectCategoryCorrectionIntent(text) {
  const lower = String(text || "").toLowerCase().trim();
  if (!lower) {
    return null;
  }

  let targetType = extractCorrectionTargetType(lower);
  const targetTaskList = findTaskListFromText(lower);
  if (!targetType && targetTaskList && /\b(move|put|assign|switch)\b/.test(lower)) {
    targetType = "task";
  }
  if (!targetType) {
    return null;
  }
  const sourceType = extractCorrectionSourceType(lower);

  const correctionPatterns = [
    /\bthat should have been\b/,
    /\bit should have been\b/,
    /\bthis should have been\b/,
    /\bthat was supposed to be\b/,
    /\bit was supposed to be\b/,
    /\byou should have created\b/,
    /\byou made\b/,
    /\bmake that\b/,
    /\bmove that\b/,
    /\bput that\b/,
    /\bthat's\b/,
    /\bthat is\b/,
    /\bthis is\b/,
    /\bno[, ]+\b/,
    /\bactually\b/,
    /\binstead\b/,
    /\bbelongs in\b/,
    /\bnot an idea\b/,
    /\bnot a task\b/,
    /\bnot meeting notes\b/,
  ];

  return correctionPatterns.some((pattern) => pattern.test(lower))
    ? { targetType, sourceType, targetTaskListId: targetTaskList?.id || "" }
    : null;
}

function detectSectionCleanupIntent(text) {
  const lower = String(text || "").toLowerCase().trim();
  if (!lower) {
    return null;
  }

  const targetType = extractCorrectionTargetType(lower);
  if (!targetType) {
    return null;
  }

  const cleanupPatterns = [
    /\b(delete|remove|clear)\s+(it|that)\b/,
    /\b(delete|remove|clear)\s+.*\bfrom\b/,
  ];

  if (!cleanupPatterns.some((pattern) => pattern.test(lower))) {
    return null;
  }

  return { targetType };
}

function extractCorrectionTargetType(text) {
  const explicitDestinationMatch = text.match(/\b(move|put|belongs)\s+(it|that)?\s*(to|in)\s+(my\s+)?(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (explicitDestinationMatch) {
    return normalizeCorrectionTypeLabel(explicitDestinationMatch[5]);
  }

  const shouldHaveBeenMatch = text.match(/\bshould have been\s+(a|an)?\s*(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (shouldHaveBeenMatch) {
    return normalizeCorrectionTypeLabel(shouldHaveBeenMatch[2]);
  }

  const wasSupposedToBeMatch = text.match(/\bwas supposed to be\s+(a|an)?\s*(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (wasSupposedToBeMatch) {
    return normalizeCorrectionTypeLabel(wasSupposedToBeMatch[2]);
  }

  const makeThatMatch = text.match(/\bmake that\s+(a|an)?\s*(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (makeThatMatch) {
    return normalizeCorrectionTypeLabel(makeThatMatch[2]);
  }

  const wasAMatch = text.match(/\bthat was\s+(a|an)?\s*(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (wasAMatch) {
    return normalizeCorrectionTypeLabel(wasAMatch[2]);
  }

  if (/\bmeeting notes\b|\bmeeting note\b|\bcall notes\b|\bcall note\b|\bnotes\b|\bnote\b/.test(text)) {
    return "meeting";
  }
  if (/\btask\b|\btasks\b/.test(text)) {
    return "task";
  }
  if (/\bidea\b|\bideas\b/.test(text)) {
    return "idea";
  }
  return "";
}

function extractCorrectionSourceType(text) {
  const explicitNegativeMatch = text.match(/\bnot\s+(a|an)?\s*(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (explicitNegativeMatch) {
    return normalizeCorrectionTypeLabel(explicitNegativeMatch[2]);
  }

  const madeMatch = text.match(/\byou made\b[^.]*\b(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b/);
  if (madeMatch) {
    return normalizeCorrectionTypeLabel(madeMatch[1]);
  }

  const leadingTypeMatch = text.match(/\b(the|this|that)\s+(meeting notes|meeting note|call notes|call note|notes|note|tasks|task|ideas|idea)\b[^.]*\bshould have been\b/);
  if (leadingTypeMatch) {
    return normalizeCorrectionTypeLabel(leadingTypeMatch[2]);
  }

  if (/\bidea\b[^.]*\bshould have been\b|\bidea\b[^.]*\bwas supposed to be\b|\byou made\b[^.]*\bidea\b/.test(text)) {
    return "idea";
  }
  if (/\btask\b[^.]*\bshould have been\b|\btask\b[^.]*\bwas supposed to be\b|\byou made\b[^.]*\btask\b/.test(text)) {
    return "task";
  }
  if (/\bmeeting notes?\b[^.]*\bshould have been\b|\bmeeting notes?\b[^.]*\bwas supposed to be\b|\byou made\b[^.]*\bmeeting notes?\b/.test(text)) {
    return "meeting";
  }
  return "";
}

function normalizeCorrectionTypeLabel(label) {
  const lower = String(label || "").toLowerCase().trim();
  if (/\bmeeting notes\b|\bmeeting note\b|\bcall notes\b|\bcall note\b|\bnotes\b|\bnote\b/.test(lower)) {
    return "meeting";
  }
  if (/\btasks\b|\btask\b/.test(lower)) {
    return "task";
  }
  if (/\bideas\b|\bidea\b/.test(lower)) {
    return "idea";
  }
  return "";
}

function applyCategoryCorrection(intent, userText) {
  const target = intent?.targetType;
  const source = resolveCorrectionSource(intent);
  if (!target || !source) {
    setMicStatus("HAL could not find a recent item to correct.", "");
    speakHal("HAL could not find a recent item to correct.");
    return;
  }

  if (source.type === target) {
    if (target === "task" && intent?.targetTaskListId && source.item.listId !== intent.targetTaskListId) {
      source.item.listId = intent.targetTaskListId;
      state.selectedTaskListId = intent.targetTaskListId;
      learnTaskListCorrection(getSourcePromptForItem(source) || userText, intent.targetTaskListId);
      learnRoutingPattern({
        sourcePrompt: getSourcePromptForItem(source) || userText,
        fromType: "task",
        toType: "task",
        listId: intent.targetTaskListId,
        correctionText: userText,
      });
      state.lastCreated = { type: "task", id: source.id };
      state.selected = { type: "task", id: source.id };
      state.correctionContext = {
        sourceType: "task",
        sourceId: source.id,
        targetType: "task",
        targetId: source.id,
      };
      persist();
      render();
      setMicStatus(`Moved to ${getTaskListName(intent.targetTaskListId)}.`, "");
      speakHal(`Moved to ${getTaskListName(intent.targetTaskListId)}.`);
      return;
    }
    setMicStatus(`That item is already in ${getSectionLabel(target)}.`, "");
    speakHal(`That item is already in ${getSectionLabel(target)}.`);
    return;
  }

  const sourcePrompt = getSourcePromptForItem(source) || source.item.title || "";
  learnIntentCorrection(sourcePrompt, source.type, target);
  rememberCorrection({
    entityType: source.type,
    entityId: source.id,
    action: "reclassify",
    fromType: source.type,
    toType: target,
    prompt: sourcePrompt,
    userText,
  });

  const replacement = buildReplacementItemFromCorrection(source, target, sourcePrompt, userText, intent);
  if (!replacement) {
    setMicStatus("HAL could not correct that item yet.", "");
    speakHal("HAL could not correct that item yet.");
    return;
  }

  learnRoutingPattern({
    sourcePrompt,
    fromType: source.type,
    toType: target,
    listId: target === "task" ? replacement.listId || "" : "",
    meetingSignature: target === "meeting" ? buildMeetingTargetSignature(replacement) : "",
    correctionText: userText,
  });

  removeItemFromSection(source.type, source.id);
  insertCorrectedItem(target, replacement);
  state.lastCreated = { type: target, id: replacement.id };
  state.selected = { type: target, id: replacement.id };
  state.correctionContext = {
    sourceType: source.type,
    sourceId: source.id,
    targetType: target,
    targetId: replacement.id,
  };

  if (target === "task") {
    state.selectedTaskListId = replacement.listId;
    learnTaskListCorrection(sourcePrompt || userText, replacement.listId);
    if (elements.tasksDateFilter) {
      elements.tasksDateFilter.value = replacement.dueDate || "";
    }
    if (elements.myDayDate && replacement.dueDate) {
      elements.myDayDate.value = replacement.dueDate;
    }
  }

  if (target === "meeting") {
    learnMeetingTargetCorrection(sourcePrompt || userText, replacement);
  }

  setMicStatus(`Moved to ${getSectionLabel(target)}.`, "");
  speakHal(`Moved to ${getSectionLabel(target)}.`);
}

function applySectionCleanup(intent) {
  const targetType = intent?.targetType;
  const candidate = resolveCleanupSource(targetType);
  if (!targetType || !candidate) {
    setMicStatus("HAL could not find the item to remove.", "");
    speakHal("HAL could not find the item to remove.");
    return;
  }

  removeItemFromSection(candidate.type, candidate.id);
  if (state.selected?.type === candidate.type && state.selected?.id === candidate.id) {
    state.selected = { type: null, id: null };
  }
  if (state.lastCreated?.type === candidate.type && state.lastCreated?.id === candidate.id) {
    state.lastCreated = { type: null, id: null };
  }
  state.correctionContext = null;
  setMicStatus(`Removed from ${getSectionLabel(candidate.type)}.`, "");
  speakHal(`Removed from ${getSectionLabel(candidate.type)}.`);
}

function resolveCorrectionSource(intent) {
  const preferredType = intent?.sourceType;
  const context = state.correctionContext;

  if (preferredType) {
    const preferredItem = getMostRecentCorrectableItemOfType(preferredType);
    if (preferredItem) {
      return preferredItem;
    }
  }

  const lastCreated = getMostRecentCorrectableItem();
  if (lastCreated && (!intent?.targetType || lastCreated.type !== intent.targetType)) {
    return lastCreated;
  }

  if (context?.sourceType && context?.sourceId) {
    const contextItem = getItemBySection(context.sourceType, context.sourceId);
    if (contextItem) {
      return { type: context.sourceType, id: context.sourceId, item: contextItem };
    }
  }

  return lastCreated;
}

function resolveCleanupSource(targetType) {
  const context = state.correctionContext;
  if (context?.sourceType === targetType && context?.sourceId) {
    const contextItem = getItemBySection(context.sourceType, context.sourceId);
    if (contextItem) {
      return { type: context.sourceType, id: context.sourceId, item: contextItem };
    }
  }

  return getMostRecentCorrectableItemOfType(targetType);
}

function buildMeetingTargetSignature(note) {
  return normalizeMeetingTargetSignature(`${note?.title || ""} ${note?.person || ""}`);
}

function normalizeMeetingTargetSignature(value) {
  return normalizeIntentPrompt(value);
}

function findMeetingNoteBySignature(signature) {
  const normalizedSignature = normalizeMeetingTargetSignature(signature);
  if (!normalizedSignature) {
    return null;
  }

  let bestMatch = null;
  let bestScore = 0;
  state.meetingNotes.forEach((note) => {
    const score = scoreIntentCorrection(normalizedSignature, buildMeetingTargetSignature(note));
    if (score > bestScore) {
      bestScore = score;
      bestMatch = note;
    }
  });

  return bestScore >= 0.72 ? bestMatch : null;
}

function getMostRecentCorrectableItem() {
  if (state.lastCreated?.type && state.lastCreated?.id) {
    const item = getItemBySection(state.lastCreated.type, state.lastCreated.id);
    if (item && ["task", "idea", "meeting"].includes(state.lastCreated.type)) {
      return { type: state.lastCreated.type, id: state.lastCreated.id, item };
    }
  }

  if (state.selected?.type && state.selected?.id) {
    const item = getItemBySection(state.selected.type, state.selected.id);
    if (item && ["task", "idea", "meeting"].includes(state.selected.type)) {
      return { type: state.selected.type, id: state.selected.id, item };
    }
  }

  return null;
}

function getMostRecentCorrectableItemOfType(type) {
  if (!type) {
    return null;
  }

  if (state.lastCreated?.type === type && state.lastCreated?.id) {
    const item = getItemBySection(type, state.lastCreated.id);
    if (item) {
      return { type, id: state.lastCreated.id, item };
    }
  }

  if (state.selected?.type === type && state.selected?.id) {
    const item = getItemBySection(type, state.selected.id);
    if (item) {
      return { type, id: state.selected.id, item };
    }
  }

  const collection = type === "task"
    ? state.tasks
    : type === "idea"
      ? state.ideas
      : state.meetingNotes;
  const item = collection[0];
  return item ? { type, id: item.id, item } : null;
}

function getItemBySection(type, id) {
  if (type === "task") {
    return state.tasks.find((item) => item.id === id) || null;
  }
  if (type === "idea") {
    return state.ideas.find((item) => item.id === id) || null;
  }
  if (type === "meeting") {
    return state.meetingNotes.find((item) => item.id === id) || null;
  }
  return null;
}

function getSourcePromptForItem(source) {
  if (source.type === "task" || source.type === "idea") {
    return source.item.sourcePrompt || "";
  }
  if (source.type === "meeting") {
    return source.item.sourcePrompt || source.item.content || source.item.title || "";
  }
  return "";
}

function buildReplacementItemFromCorrection(source, targetType, sourcePrompt, correctionText = "", intent = null) {
  if (targetType === "task") {
    const correctionTaskList = intent?.targetTaskListId
      ? state.taskLists.find((list) => list.id === intent.targetTaskListId)
      : findTaskListFromText(correctionText);
    const taskDraft = buildTaskInterpretation(sourcePrompt || source.item.title, "", correctionTaskList?.id || "");
    return {
      id: crypto.randomUUID(),
      title: taskDraft.title,
      listId: correctionTaskList?.id || taskDraft.listId || source.item.listId || state.selectedTaskListId || state.taskLists[0].id,
      dueDate: taskDraft.dueDate || source.item.dueDate || source.item.callDate || "",
      dueTime: taskDraft.dueTime || source.item.dueTime || "",
      done: false,
      notes: source.type === "task" ? source.item.notes || "" : "",
      highlighted: source.type === "task" ? Boolean(source.item.highlighted) : false,
      sourcePrompt,
      createdAt: new Date().toISOString(),
    };
  }

  if (targetType === "idea") {
    const formatted = source.type === "meeting"
      ? source.item.content
      : source.type === "task"
        ? ""
        : source.item.content;
    const ideaDraft = buildIdeaInterpretation(sourcePrompt || source.item.title, source.item.title, formatted);
    return {
      id: crypto.randomUUID(),
      title: ideaDraft.title,
      content: ideaDraft.formatted,
      archived: false,
      sourcePrompt,
      createdAt: new Date().toISOString(),
    };
  }

  if (targetType === "meeting") {
    const sourceBody = source.type === "idea"
      ? stripHtml(source.item.content || "")
      : source.type === "task"
        ? ""
        : source.item.content || "";
    const formatted = source.type === "meeting"
      ? source.item.content
      : formatMeetingNote(sourceBody || sourcePrompt || source.item.title);
    const meetingDraft = buildMeetingInterpretation(sourcePrompt || source.item.title, source.item.title, formatted, source.item.person || "");
    return {
      id: crypto.randomUUID(),
      title: meetingDraft.title,
      person: meetingDraft.person,
      callDate: meetingDraft.dueDate || source.item.dueDate || "",
      content: meetingDraft.formatted,
      sourcePrompt,
      createdAt: new Date().toISOString(),
    };
  }

  return null;
}

function removeItemFromSection(type, id) {
  if (type === "task") {
    state.tasks = state.tasks.filter((item) => item.id !== id);
  } else if (type === "idea") {
    state.ideas = state.ideas.filter((item) => item.id !== id);
  } else if (type === "meeting") {
    state.meetingNotes = state.meetingNotes.filter((item) => item.id !== id);
  }
}

function insertCorrectedItem(type, item) {
  if (type === "task") {
    state.tasks.unshift(item);
  } else if (type === "idea") {
    state.ideas.unshift(item);
  } else if (type === "meeting") {
    state.meetingNotes.unshift(item);
  }
}

function getSectionLabel(type) {
  if (type === "task") {
    return "Tasks";
  }
  if (type === "meeting") {
    return "Meeting Notes";
  }
  return "Ideas";
}

function isTaskIntent(text, referencedTaskList = null) {
  const collapsed = text.replace(/\s+/g, " ").trim();
  const compact = text.replace(/\s+/g, "");
  if (hasExplicitMeetingLanguage(collapsed) && !referencedTaskList && !hasExplicitTaskLanguage(collapsed)) {
    return false;
  }
  return (
    /create.*task/.test(compact) ||
    hasExplicitTaskLanguage(collapsed) ||
    /\bput\s+.*\bin\s+my\b.*\btask\s+list\b/.test(collapsed) ||
    /\bin\s+(my|the)\b.*\btask\s+list\b/.test(collapsed) ||
    /task.*for.*me/.test(compact) ||
    collapsed.startsWith("task ") ||
    /\badd\s+(this\s+)?to\s+my\s+tasks\b/.test(collapsed) ||
    /\bremind\s+me\s+to\b/.test(collapsed) ||
    /\bfor me to\b/.test(collapsed) ||
    Boolean(referencedTaskList)
  );
}

function isReminderIntent(text) {
  const collapsed = text.replace(/\s+/g, " ").trim();
  const compact = text.replace(/\s+/g, "");
  return (
    /\bcreate\s+a?\s*reminder\b/.test(collapsed) ||
    /createreminder/.test(compact) ||
    /\bset\s+a?\s*reminder\b/.test(collapsed) ||
    /\bnew\s+reminder\b/.test(collapsed) ||
    /\bmake\s+(me\s+)?a?\s*reminder\b/.test(collapsed) ||
    /\badd\s+(a\s+)?reminder\b/.test(collapsed) ||
    /\breminder\b/.test(collapsed) ||
    /\bremind me\b/.test(collapsed)
  );
}

function isReminderDeleteIntent(text) {
  return (
    /\b(delete|remove|cancel|clear)\s+(my\s+)?reminder\b/.test(text) ||
    /\b(delete|remove|cancel|clear)\s+(my\s+)?reminder\s+for\b/.test(text) ||
    /\bcan\s+(my\s+)?reminder\b/.test(text) ||
    /\bcan\s+(my\s+)?reminder\s+for\b/.test(text)
  );
}

function isReminderEditIntent(text) {
  return (
    /\b(edit|change|update|modify)\s+(my\s+)?reminder\b/.test(text) ||
    /\b(edit|change|update|modify)\s+(my\s+)?reminder\s+for\b/.test(text)
  );
}

function isMeetingIntent(text) {
  return (
    /\bcreate\s+(a\s+)?meeting\s+note\b/.test(text) ||
    /\bcreate\s+notes?\s+for\b/.test(text) ||
    /\bmake\s+(a\s+)?meeting\s+note\b/.test(text) ||
    /\bmake\s+notes?\s+for\b/.test(text) ||
    /\bmeeting\s+note\b/.test(text) ||
    /\bcall with\b/.test(text) ||
    /\bmeeting notes\b/.test(text) ||
    /\bnotes for my call\b/.test(text) ||
    /\bnotes for my meeting\b/.test(text) ||
    /\btalking points\b/.test(text)
  );
}

function isMeetingUpdateIntent(text, targetNote = null) {
  return Boolean(targetNote) && (
    /\b(add|append|put|drop|include)\b[^.]*\b(meeting notes?|notes?)\b/.test(text) ||
    /\b(add|append|put|drop|include)\b[^.]*\b(call with|meeting with)\b/.test(text) ||
    /\bupdate\b[^.]*\b(meeting notes?|notes?)\b/.test(text) ||
    /\bfor my call with\b/.test(text) && /\b(add|append|put|include)\b/.test(text) ||
    /\bto my\b[^.]*\bnotes?\b/.test(text)
  );
}

function isLikelyMeetingUpdatePrompt(text) {
  const lower = String(text || "").toLowerCase();
  return (
    /\b(add|append|put|drop|include|update)\b/.test(lower) ||
    /\bto my\b[^.]*\b(meeting notes?|notes?)\b/.test(lower)
  );
}

function hasExplicitTaskLanguage(text) {
  return (
    /\bcreate\s+(a\s+)?task\b/.test(text) ||
    /\badd\s+(a\s+)?task\b/.test(text) ||
    /\bmake\s+(a\s+)?task\b/.test(text) ||
    /\bnew\s+task\b/.test(text) ||
    /\btask\s+for\s+me\b/.test(text) ||
    /\btask\s+to\b/.test(text) ||
    /\badd\s+(this\s+)?to\s+my\s+tasks\b/.test(text) ||
    /\bmove\b[^.]*\bto\b[^.]*\btask\s+list\b/.test(text)
  );
}

function hasExplicitMeetingLanguage(text) {
  return (
    /\bmeeting\s+note\b/.test(text) ||
    /\bmeeting\s+notes\b/.test(text) ||
    /\bnotes?\s+for\s+my\s+call\b/.test(text) ||
    /\bnotes?\s+for\s+my\s+meeting\b/.test(text) ||
    /\bcall with\b/.test(text) ||
    /\bmeeting with\b/.test(text) ||
    /\btalking points\b/.test(text)
  );
}

function getSelectedItem() {
  if (state.selected.type === "idea") {
    const idea = state.ideas.find((item) => item.id === state.selected.id);
    return idea ? { title: idea.title, content: idea.content, sectionLabel: "Idea", meta: [idea.archived ? "Archived" : "Active", formatDateTime(idea.createdAt)] } : null;
  }
  if (state.selected.type === "meeting") {
    const note = state.meetingNotes.find((item) => item.id === state.selected.id);
    return note ? { title: note.title, content: note.content, sectionLabel: "Meeting Notes", meta: [note.callDate ? formatDisplayDate(note.callDate) : "No date", note.person || "No contact"] } : null;
  }
  if (state.selected.type === "task") {
    const task = state.tasks.find((item) => item.id === state.selected.id);
    return task ? { title: task.title, content: task.done ? "This task is marked complete." : "This task is still open.", sectionLabel: "Task", meta: [formatTaskDue(task), getTaskListName(task.listId)] } : null;
  }
  return null;
}

function buildSelectedDraft(item) {
  return `${item.title}\n${"=".repeat(item.title.length)}\n\nSection: ${item.sectionLabel}\n${item.meta.join("\n")}\n\n${item.content}`;
}

function buildMarkdownDraft(item) {
  return `# ${item.title}\n\n- Section: ${item.sectionLabel}\n- ${item.meta.join("\n- ")}\n\n## Content\n\n${item.content}`;
}

function buildPrompts(item) {
  if (!item) {
    return [];
  }
  if (item.sectionLabel === "Idea") {
    return [
      { title: "Clarify the problem", body: "What friction or inefficiency is this idea trying to remove?" },
      { title: "Choose a small test", body: "What is the fastest version of this idea you could try first?" },
      { title: "Decide the output", body: "Should this become a proposal, checklist, process doc, or presentation?" },
    ];
  }
  if (item.sectionLabel === "Meeting Notes") {
    return [
      { title: "Outcome first", body: "What needs to happen in the meeting for it to be a win?" },
      { title: "Top talking points", body: "What are the three most important points to cover?" },
      { title: "Expiration rule", body: "Once this meeting is past, HAL can auto-hide or expire the note later." },
    ];
  }
  return [
    { title: "Break it down", body: "Should this task be split into smaller steps?" },
    { title: "Assign a date", body: "If this belongs in My Day later, give it a due date." },
    { title: "Move lists if needed", body: "Would this be easier to find in a different task list?" },
  ];
}

function getTasksForDate(date) {
  return state.tasks.filter((task) => task.dueDate === date);
}

function getPastDueTasks() {
  const today = todayString();
  return state.tasks.filter((task) => !task.done && task.dueDate && task.dueDate < today);
}

function getTaskListName(id) {
  return state.taskLists.find((list) => list.id === id)?.name || "Regular Tasks";
}

function formatTaskDue(task) {
  if (!task?.dueDate) {
    return "No date";
  }
  return task.dueTime ? `${formatDisplayDate(task.dueDate)} at ${formatTimeValue(task.dueTime)}` : formatDisplayDate(task.dueDate);
}

function formatTimeValue(value) {
  const [hourText, minuteText] = String(value || "").split(":");
  const hour = Number(hourText);
  const minute = Number(minuteText || "0");
  if (Number.isNaN(hour)) {
    return value;
  }
  return new Intl.DateTimeFormat("en-US", { hour: "numeric", minute: "2-digit" }).format(new Date(2026, 0, 1, hour, minute));
}

function ensureDefaultDate() {
  if (!elements.myDayDate.value) {
    elements.myDayDate.value = todayString();
  }
  if (!state.selectedTaskListId && state.taskLists.length) {
    state.selectedTaskListId = state.taskLists[0].id;
  }
}

function parseDueDate(text) {
  const baseDate = new Date();
  const weekdayNames = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

  if (/\btomorrow\b/.test(text)) {
    const date = new Date(baseDate);
    date.setDate(date.getDate() + 1);
    return toLocalDateInputValue(date);
  }

  if (/\btoday\b/.test(text) || /\btonight\b/.test(text) || /\bthis afternoon\b/.test(text) || /\bthis morning\b/.test(text) || /\bthis evening\b/.test(text)) {
    return toLocalDateInputValue(baseDate);
  }

  if (/\bthis weekend\b/.test(text)) {
    return toLocalDateInputValue(thisWeekendDate(baseDate));
  }

  const nextWeekdayMatch = text.match(/\bnext (monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/);
  if (nextWeekdayMatch) {
    const targetIndex = weekdayNames.indexOf(nextWeekdayMatch[1]);
    return toLocalDateInputValue(nextWeekdayDate(baseDate, targetIndex, true));
  }

  const weekdayMatch = text.match(/\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/);
  if (weekdayMatch) {
    const targetIndex = weekdayNames.indexOf(weekdayMatch[1]);
    return toLocalDateInputValue(nextWeekdayDate(baseDate, targetIndex, false));
  }

  return "";
}

function extractPerson(text) {
  const match = text.match(/call with ([a-z\s]+)/i);
  return match ? toTitleCase(match[1].trim()) : "";
}

function cleanTaskTitle(text) {
  return text
    .replace(/^create\s+a?\s*task\b/i, "")
    .replace(/create\s*a?\s*task\s+for\s+me\s+to/i, "")
    .replace(/create\s*a?\s*task\s+to/i, "")
    .replace(/create\s*a?\s*task\s+in\s+my\s+.+?\s+(task\s+)?list\s+(for me to|to)/i, "")
    .replace(/create\s*a?\s*task\s+in\s+the\s+.+?\s+(task\s+)?list\s+(for me to|to)/i, "")
    .replace(/^a\s+task\s+to\b/i, "")
    .replace(/^task\s+to\b/i, "")
    .replace(/task\s+for\s+me\s+to/i, "")
    .replace(/\bin\s+my\s+.+?\s+(task\s+)?list\b/i, "")
    .replace(/\bin\s+the\s+.+?\s+(task\s+)?list\b/i, "")
    .replace(/opinionon/gi, "opinion on")
    .replace(/\bat\s+\d{1,2}(:\d{2})?\s*(am|pm)?\b/i, "")
    .replace(/\bit'?s?\s+due\s+tomorrow\b/i, "")
    .replace(/\bit'?s?\s+due\s+today\b/i, "")
    .replace(/\bit'?s?\s+due\s+tonight\b/i, "")
    .replace(/\bit'?s?\s+due\s+this afternoon\b/i, "")
    .replace(/\bit'?s?\s+due\s+this morning\b/i, "")
    .replace(/\bit'?s?\s+due\s+this evening\b/i, "")
    .replace(/\bit'?s?\s+due\s+this weekend\b/i, "")
    .replace(/\bit'?s?\s+due\s+next\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, "")
    .replace(/\bdue\s+tomorrow\b/i, "")
    .replace(/\bdue\s+today\b/i, "")
    .replace(/\bdue\s+tonight\b/i, "")
    .replace(/\bdue\s+this afternoon\b/i, "")
    .replace(/\bdue\s+this morning\b/i, "")
    .replace(/\bdue\s+this evening\b/i, "")
    .replace(/\bdue\s+this weekend\b/i, "")
    .replace(/\bdue\s+next\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, "")
    .replace(/\btomorrow\b/i, "")
    .replace(/\btoday\b/i, "")
    .replace(/\btonight\b/i, "")
    .replace(/\bthis afternoon\b/i, "")
    .replace(/\bthis morning\b/i, "")
    .replace(/\bthis evening\b/i, "")
    .replace(/\bthis weekend\b/i, "")
    .replace(/\bnext\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, "")
    .replace(/\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+(morning|afternoon|evening|night)\b/i, "$1")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\.$/, "") || "New task";
}

function extractExplicitTitle(text) {
  const match = String(text || "").match(/\bcall it\s+["“]?(.+?)["”]?\s*$/i);
  return match ? match[1].trim().replace(/[.?!]+$/, "") : "";
}

function stripExplicitTitleInstruction(text) {
  return String(text || "")
    .replace(/\s*,?\s*\bcall it\s+["“]?.+?["”]?\s*$/i, "")
    .trim();
}

function findTaskListFromText(text) {
  const lower = text.toLowerCase();
  const compactText = compactTaskListReference(lower);

  const strictNamedPatterns = [
    /\btask\s+list\s+called\s+(.+?)(?:\s+to\b|\s+for\b|\s+about\b|$)/,
    /\bin\s+my\s+task\s+list\s+called\s+(.+?)(?:\s+to\b|\s+for\b|\s+about\b|$)/,
    /\bin\s+the\s+task\s+list\s+called\s+(.+?)(?:\s+to\b|\s+for\b|\s+about\b|$)/,
  ];

  for (const pattern of strictNamedPatterns) {
    const match = lower.match(pattern);
    if (!match) {
      continue;
    }
    const requested = normalizeTaskListReference(match[1]);
    const namedMatch = state.taskLists.find((list) => normalizeTaskListReference(list.name) === requested)
      || state.taskLists.find((list) => normalizeTaskListReference(list.name).includes(requested))
      || state.taskLists.find((list) => requested.includes(normalizeTaskListReference(list.name)));
    if (namedMatch) {
      return namedMatch;
    }
  }

  const directMatch = state.taskLists.find((list) => {
    const listName = list.name.toLowerCase().trim();
    const compactListName = compactTaskListReference(listName);
    return (
      lower.includes(`${listName} task list`) ||
      lower.includes(`${listName} list`) ||
      lower.includes(`my ${listName}`) ||
      lower.includes(`the ${listName}`) ||
      compactText.includes(`${compactListName}tasklist`) ||
      compactText.includes(`${compactListName}list`) ||
      compactText.includes(compactListName)
    );
  });
  if (directMatch) {
    return directMatch;
  }

  const calledMatch = lower.match(/\btask\s+list\s+called\s+(.+?)(?:\s+to\b|\s+for\b|$)/);
  if (calledMatch) {
    const requested = normalizeTaskListReference(calledMatch[1]);
    const namedMatch = state.taskLists.find((list) => normalizeForDuplicateCheck(list.name) === requested)
      || state.taskLists.find((list) => normalizeForDuplicateCheck(list.name).includes(requested))
      || state.taskLists.find((list) => requested.includes(normalizeForDuplicateCheck(list.name)));
    if (namedMatch) {
      return namedMatch;
    }
  }

  const match = lower.match(/\bin\s+(?:my|the)\s+(.+?)\s+(?:task\s+)?list\b/);
  if (match) {
    const requested = normalizeTaskListReference(match[1]);
    return state.taskLists.find((list) => normalizeForDuplicateCheck(list.name) === requested)
      || state.taskLists.find((list) => normalizeForDuplicateCheck(list.name).includes(requested))
      || state.taskLists.find((list) => requested.includes(normalizeForDuplicateCheck(list.name)))
      || null;
  }

  const normalizedText = normalizeTaskListReference(text);
  let bestMatch = null;
  let bestScore = 0;
  state.taskLists.forEach((list) => {
    const score = scoreTaskListMatch(normalizedText, normalizeTaskListReference(list.name));
    if (score > bestScore) {
      bestScore = score;
      bestMatch = list;
    }
  });

  return bestScore >= 0.45 ? bestMatch : null;
}

function normalizeTaskListReference(value) {
  return normalizeForDuplicateCheck(value).replace(/\btask\b/g, "").replace(/\s+/g, " ").trim();
}

function compactTaskListReference(value) {
  return normalizeTaskListReference(value).replace(/\s+/g, "");
}

function scoreTaskListMatch(source, target) {
  if (!source || !target) {
    return 0;
  }

  const compactSource = compactTaskListReference(source);
  const compactTarget = compactTaskListReference(target);
  if (compactSource.includes(compactTarget) || compactTarget.includes(compactSource)) {
    return 1;
  }

  const sourceTokens = source.split(" ").filter(Boolean);
  const targetTokens = target.split(" ").filter(Boolean);
  if (!sourceTokens.length || !targetTokens.length) {
    return 0;
  }

  const targetSet = new Set(targetTokens);
  const shared = sourceTokens.filter((token) => targetSet.has(token)).length;
  return shared / Math.max(sourceTokens.length, targetTokens.length);
}

function findMatchingMeetingNote(text) {
  const lower = String(text || "").toLowerCase().trim();
  if (!lower || !state.meetingNotes.length) {
    return null;
  }

  const directPerson = extractPerson(lower);
  if (directPerson) {
    const directMatch = state.meetingNotes.find((note) => {
      const title = normalizeForDuplicateCheck(note.title);
      const person = normalizeForDuplicateCheck(note.person || "");
      const target = normalizeForDuplicateCheck(directPerson);
      return title.includes(target) || person.includes(target);
    });
    if (directMatch) {
      return directMatch;
    }
  }

  const normalizedText = normalizeIntentPrompt(stripMeetingReferencePhrases(lower));
  if (!normalizedText) {
    return null;
  }

  let bestMatch = null;
  let bestScore = 0;
  state.meetingNotes.forEach((note) => {
    const noteSignature = normalizeIntentPrompt(`${note.title} ${note.person || ""}`);
    const score = scoreIntentCorrection(normalizedText, noteSignature);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = note;
    }
  });

  return bestScore >= 0.45 ? bestMatch : null;
}

function cleanMeetingUpdateText(text) {
  return String(text || "")
    .replace(/\b(add|append|put|drop|include)\b\s+(this\s+)?(to|in)\s+(my\s+)?(existing\s+)?meeting\s+notes?\s+for\s+[^,.:]+[:,]?\s*/i, "")
    .replace(/\b(add|append|put|drop|include)\b\s+(this\s+)?(to|in)\s+(my\s+)?(existing\s+)?notes?\s+for\s+my\s+(call|meeting)\s+with\s+[^,.:]+[:,]?\s*/i, "")
    .replace(/\b(add|append|put|drop|include)\b\s+(this\s+)?(to|in)\s+(my\s+)?(existing\s+)?meeting\s+notes?\b[:,]?\s*/i, "")
    .replace(/\b(add|append|put|drop|include)\b\s+(this\s+)?(to|in)\s+(my\s+)?(existing\s+)?notes?\b[:,]?\s*/i, "")
    .replace(/\bupdate\s+(my\s+)?meeting\s+notes?\s+for\s+[^,.:]+[:,]?\s*/i, "")
    .replace(/\bupdate\s+(my\s+)?notes?\s+for\s+my\s+(call|meeting)\s+with\s+[^,.:]+[:,]?\s*/i, "")
    .trim();
}

function stripMeetingReferencePhrases(text) {
  return String(text || "")
    .replace(/\b(add|append|put|drop|include|update)\b/g, " ")
    .replace(/\b(this|that|it|my|existing)\b/g, " ")
    .replace(/\bmeeting\s+notes?\b/g, " ")
    .replace(/\bnotes?\b/g, " ")
    .replace(/\bfor\s+my\s+(call|meeting)\s+with\b/g, " ")
    .replace(/\bcall\s+with\b/g, " ")
    .replace(/\bmeeting\s+with\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function appendMeetingNoteContent(existingContent, additionContent) {
  const existing = String(existingContent || "").trim();
  const addition = String(additionContent || "").trim();
  if (!addition) {
    return existing;
  }

  if (/<[a-z][\s\S]*>/i.test(existing)) {
    const existingHtml = sanitizeIdeaEditorHtml(existing);
    const additionHtml = formatStructuredNoteForDisplay(addition);
    if (!existingHtml) {
      return additionHtml;
    }
    return `${existingHtml}${additionHtml ? additionHtml : ""}`;
  }

  const normalizedAddition = addition.replace(/^Talking Points\s*\n?/i, "").trim();
  if (!existing) {
    return addition;
  }
  if (!normalizedAddition) {
    return existing;
  }
  if (/^Talking Points\b/i.test(existing)) {
    return `${existing}\n${normalizedAddition}`;
  }
  return `${existing}\n\n${addition}`;
}

function compareCalendarEvents(left, right) {
  return parseCalendarEventTime(left?.time) - parseCalendarEventTime(right?.time);
}

function parseCalendarEventTime(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) {
    return Number.MAX_SAFE_INTEGER;
  }

  const match = raw.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)?/);
  if (!match) {
    return Number.MAX_SAFE_INTEGER - 1;
  }

  let hours = Number(match[1]);
  const minutes = Number(match[2] || "0");
  const meridiem = match[3];
  if (meridiem === "pm" && hours < 12) {
    hours += 12;
  }
  if (meridiem === "am" && hours === 12) {
    hours = 0;
  }

  return (hours * 60) + minutes;
}

function parseClientCsvCalendarEvents(csv) {
  const lines = String(csv || "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return [];
  }

  return lines.slice(1)
    .map(parseClientCsvCalendarRow)
    .filter(Boolean)
    .sort((left, right) => compareCalendarEvents(left, right));
}

function parseClientCsvCalendarRow(line) {
  const values = splitSimpleCsvLineClient(line);
  if (values.length < 3) {
    return null;
  }

  const title = values[0]?.trim();
  const date = normalizeClientCsvDate(values[1]);
  const time = normalizeClientCsvTime(values[2]);

  if (!title || !date) {
    return null;
  }

  return {
    id: crypto.randomUUID(),
    title,
    date,
    time,
    source: "csv",
  };
}

function splitSimpleCsvLineClient(line) {
  const values = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];

    if (character === '"') {
      const nextCharacter = line[index + 1];
      if (inQuotes && nextCharacter === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === "," && !inQuotes) {
      values.push(current);
      current = "";
      continue;
    }

    current += character;
  }

  values.push(current);
  return values;
}

function normalizeClientCsvDate(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const month = slashMatch[1].padStart(2, "0");
    const day = slashMatch[2].padStart(2, "0");
    return `${slashMatch[3]}-${month}-${day}`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }

  return toLocalDateInputValue(parsed);
}

function normalizeClientCsvTime(value) {
  return String(value || "").trim();
}

function normalizeIntentPrompt(value) {
  return tokenizeIntentPrompt(value).join(" ");
}

function tokenizeIntentPrompt(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token && !INTENT_LEARNING_STOP_WORDS.has(token))
    .map((token) => token.endsWith("s") && token.length > 3 ? token.slice(0, -1) : token);
}

function scoreIntentCorrection(source, target) {
  const sourceTokens = tokenizeIntentPrompt(source);
  const targetTokens = tokenizeIntentPrompt(target);
  if (!sourceTokens.length || !targetTokens.length) {
    return 0;
  }

  const targetSet = new Set(targetTokens);
  const shared = sourceTokens.filter((token) => targetSet.has(token)).length;
  return shared / Math.max(sourceTokens.length, targetTokens.length);
}

function parseTaskTimeDetails(text) {
  return parseReminderTimeDetails(text);
}

function cleanReminderText(text) {
  return text
    .replace(/\b(create|set)\s+a?\s*reminder\s+(for|to)\b/i, "")
    .replace(/\bcreate\s+a?\s*reminder\b/i, "")
    .replace(/\bset\s+a?\s*reminder\b/i, "")
    .replace(/\bmake\s+(me\s+)?a?\s*reminder\s+(for|to)\b/i, "")
    .replace(/\badd\s+(a\s+)?reminder\s+(for|to)\b/i, "")
    .replace(/\bnew\s+reminder\s+(for|to)\b/i, "")
    .replace(/\bremind me to\b/i, "")
    .replace(/\b(delete|remove|cancel|clear)\s+(my\s+)?reminder\s+for\b/i, "")
    .replace(/\b(delete|remove|cancel|clear)\s+(my\s+)?reminder\b/i, "")
    .replace(/\bcan\s+(my\s+)?reminder\s+for\b/i, "")
    .replace(/\bcan\s+(my\s+)?reminder\b/i, "")
    .replace(/\bmy\s+reminder\s+for\b/i, "")
    .replace(/\bmy\s+reminder\b/i, "")
    .replace(/\breminder\s+for\b/i, "")
    .replace(/\breminder\b/i, "")
    .replace(/\beach\s+day\b/i, "")
    .replace(/\bevery\s+day\b/i, "")
    .replace(/\beach\s+week\b/i, "")
    .replace(/\bevery\s+week\b/i, "")
    .replace(/\beach\s+month\b/i, "")
    .replace(/\bevery\s+month\b/i, "")
    .replace(/\bweekdays\b/i, "")
    .replace(/\ball days of the week\b/i, "")
    .replace(/\beach weekday\b/i, "")
    .replace(/\bevery weekday\b/i, "")
    .replace(/\bmonday through friday\b/i, "")
    .replace(/\b(morning|afternoon|evening|tonight)\b/i, "")
    .replace(/\btomorrow\b/i, "")
    .replace(/\btoday\b/i, "")
    .replace(/\bnext\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, "")
    .replace(/\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i, "")
    .replace(/\bat\s+\d{1,2}(:\d{2})?\s*(am|pm)?\b/i, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\.$/, "") || "New reminder";
}

function parseReminderRecurrence(text) {
  if (/\b(first|second|third|fourth)\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+of\s+every\s+month\b/.test(text)) {
    const match = text.match(/\b(first|second|third|fourth)\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+of\s+every\s+month\b/);
    return {
      mode: "monthlyWeekday",
      ordinal: match[1],
      weekday: match[2],
    };
  }

  if (/\b(every|each)\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/.test(text)) {
    const match = text.match(/\b(every|each)\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/);
    return {
      mode: "weeklyWeekday",
      weekday: match[2],
    };
  }

  if (/\b(weekdays|every weekday|each weekday|monday through friday)\b/.test(text)) {
    return { mode: "interval", interval: "weekdays" };
  }

  if (/\b(all days of the week|every single day of the week)\b/.test(text)) {
    return { mode: "interval", interval: "allDays" };
  }

  if (/\b(every|each)\s+week\b/.test(text)) {
    return { mode: "interval", interval: "weekly" };
  }

  if (/\b(every|each)\s+month\b/.test(text)) {
    return { mode: "interval", interval: "monthly" };
  }

  if (/\b(every|each)\s+day\b/.test(text)) {
    return { mode: "interval", interval: "daily" };
  }

  return null;
}

function parseReminderDateTime(text) {
  const date = parseDueDate(text) || todayString();
  const timeDetails = parseReminderTimeDetails(text);
  return `${date}T${timeDetails.time}`;
}

function parseReminderTimeDetails(text) {
  const explicitMatch = text.match(/\bat\s+(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\b/i) || text.match(/\b(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/i);
  if (explicitMatch) {
    let hours = Number(explicitMatch[1]);
    const minutes = explicitMatch[2] || "00";
    const meridiem = explicitMatch[3]?.toLowerCase();
    if (meridiem === "pm" && hours < 12) {
      hours += 12;
    }
    if (meridiem === "am" && hours === 12) {
      hours = 0;
    }
    return {
      time: `${String(hours).padStart(2, "0")}:${minutes}`,
      explicit: true,
    };
  }

  if (/\bmorning\b/.test(text)) {
    return { time: "09:00", explicit: true };
  }
  if (/\bafternoon\b/.test(text)) {
    return { time: "15:00", explicit: true };
  }
  if (/\bevening\b/.test(text) || /\btonight\b/.test(text)) {
    return { time: "18:00", explicit: true };
  }
  return { time: "09:00", explicit: false };
}

function findMatchingReminder(text) {
  const cleanedTarget = cleanReminderText(text);
  const normalizedTarget = normalizeReminderText(cleanedTarget);
  const targetTokens = tokenizeReminderText(cleanedTarget);
  const recurrence = parseReminderRecurrence(text.toLowerCase());

  let reminders = state.reminders.slice();
  if (recurrence?.interval) {
    reminders = reminders.filter((reminder) => reminder.recurrence?.interval === recurrence.interval);
  } else if (recurrence?.mode === "weeklyWeekday") {
    reminders = reminders.filter((reminder) => reminder.recurrence?.mode === "weeklyWeekday" && reminder.recurrence.weekday === recurrence.weekday);
  } else if (recurrence?.mode === "monthlyWeekday") {
    reminders = reminders.filter((reminder) => reminder.recurrence?.mode === "monthlyWeekday" && reminder.recurrence.ordinal === recurrence.ordinal && reminder.recurrence.weekday === recurrence.weekday);
  }

  const exactMatch = reminders.find((reminder) => normalizeReminderText(reminder.text) === normalizedTarget);
  if (exactMatch) {
    return exactMatch;
  }

  const inclusionMatch = reminders.find((reminder) => normalizeReminderText(reminder.text).includes(normalizedTarget))
    || reminders.find((reminder) => normalizedTarget.includes(normalizeReminderText(reminder.text)));
  if (inclusionMatch) {
    return inclusionMatch;
  }

  let bestMatch = null;
  let bestScore = 0;
  reminders.forEach((reminder) => {
    const score = scoreReminderMatch(targetTokens, tokenizeReminderText(reminder.text));
    if (score > bestScore) {
      bestScore = score;
      bestMatch = reminder;
    }
  });

  return bestScore >= 0.5 ? bestMatch : null;
}

function normalizeReminderText(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizeReminderText(value) {
  return normalizeReminderText(value)
    .split(" ")
    .map((token) => token.trim())
    .filter((token) => token && !REMINDER_STOP_WORDS.has(token))
    .map((token) => token.endsWith("s") && token.length > 3 ? token.slice(0, -1) : token);
}

function scoreReminderMatch(targetTokens, candidateTokens) {
  if (!targetTokens.length || !candidateTokens.length) {
    return 0;
  }

  const candidateSet = new Set(candidateTokens);
  const sharedCount = targetTokens.filter((token) => candidateSet.has(token)).length;
  return sharedCount / Math.max(targetTokens.length, candidateTokens.length);
}

function findDuplicateTask(title, listId, dueDate) {
  const normalizedTitle = normalizeForDuplicateCheck(title);
  return state.tasks.find((task) => (
    !task.done &&
    task.listId === listId &&
    (task.dueDate || "") === (dueDate || "") &&
    normalizeForDuplicateCheck(task.title) === normalizedTitle
  )) || null;
}

function normalizeForDuplicateCheck(value) {
  const cleaned = cleanTaskTitle(String(value));
  return cleaned
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function generateIdeaTitle(text) {
  const cleaned = stripStructuredIdeaIntro(text)
    .replace(/i'?m brainstorming an idea/i, "")
    .replace(/brainstorming/i, "")
    .replace(/idea/i, "")
    .replace(/\bthe following\b[:\s-]*/i, "")
    .trim();
  const words = cleaned.split(/\s+/).filter(Boolean).slice(0, 6);
  return toTitleCase(words.join(" ")) || "New idea";
}

function formatMeetingNote(text) {
  const plain = stripStructuredMeetingIntro(text).trim();
  if (!plain) {
    return "";
  }

  const bullets = extractStructuredBulletItems(text, "meeting")
    .map((part) => `- ${part}`);

  return bullets.length ? `Talking Points\n${bullets.join("\n")}` : plain;
}

function formatIdea(text) {
  const plain = stripStructuredIdeaIntro(text).trim();
  if (!plain) {
    return "";
  }

  const bullets = extractStructuredBulletItems(text, "idea");
  if (bullets.length) {
    return `Idea Points\n${bullets.map((part) => `- ${part}`).join("\n")}`;
  }

  return plain;
}

function extractStructuredBulletItems(text, mode = "idea") {
  const source = mode === "meeting"
    ? stripStructuredMeetingIntro(text)
    : stripStructuredIdeaIntro(text);
  const normalized = source
    .replace(/\r/g, "\n")
    .replace(/\bthe following\b[:\s-]*/gi, "")
    .replace(/\n+/g, "\n")
    .trim();

  let items = normalized
    .split(/\n|;|,|•/g)
    .map((part) => part.trim())
    .filter(Boolean);

  if (items.length <= 1) {
    items = normalized
      .split(/[.!?]+/g)
      .map((part) => part.trim())
      .filter(Boolean);
  }

  return items
    .map((part) => part.replace(/^[-*]\s*/, "").trim())
    .filter(Boolean);
}

function stripStructuredMeetingIntro(text) {
  return String(text || "")
    .replace(/create a meeting note for me for my call with [^,]+[:,]?\s*/i, "")
    .replace(/create a meeting note for my call with [^,]+[:,]?\s*/i, "")
    .replace(/create meeting notes? for me for my call with [^,]+[:,]?\s*/i, "")
    .replace(/create meeting notes? for my call with [^,]+[:,]?\s*/i, "")
    .replace(/create notes? for my call with [^,]+[:,]?\s*/i, "")
    .replace(/i need to create some notes for my call with [^,]+[:,]?\s*/i, "")
    .replace(/i need notes for my call with [^,]+[:,]?\s*/i, "")
    .replace(/i need notes for my meeting with [^,]+[:,]?\s*/i, "")
    .replace(/make a note to talk about (the )?following[:\s-]*/i, "")
    .replace(/make notes? to talk about (the )?following[:\s-]*/i, "")
    .replace(/notes? for my call with [^,]+[:,]?\s*/i, "")
    .trim();
}

function stripStructuredIdeaIntro(text) {
  return String(text || "")
    .replace(/create an? idea for me[:,]?\s*/i, "")
    .replace(/create an? idea[:,]?\s*/i, "")
    .replace(/make an? idea for me[:,]?\s*/i, "")
    .replace(/make an? idea[:,]?\s*/i, "")
    .replace(/i'?m brainstorming an idea/i, "")
    .replace(/brainstorm (the )?following[:\s-]*/i, "")
    .replace(/brainstorming (the )?following[:\s-]*/i, "")
    .replace(/ideas? for (the )?following[:\s-]*/i, "")
    .replace(/make a note of (the )?following[:\s-]*/i, "")
    .trim();
}

function todayString() {
  return toLocalDateInputValue(new Date());
}

function tomorrowString() {
  const date = new Date();
  date.setDate(date.getDate() + 1);
  return toLocalDateInputValue(date);
}

function formatDisplayDate(value) {
  return new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", year: "numeric" }).format(new Date(`${value}T00:00:00`));
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }).format(new Date(value));
}

function getReminderScheduleLabel(reminder) {
  if (reminder.frequency !== "recurring" || !reminder.recurrence) {
    return "one time";
  }

  if (reminder.recurrence.mode === "monthlyWeekday") {
    return `${capitalize(reminder.recurrence.ordinal)} ${capitalize(reminder.recurrence.weekday)} of every month`;
  }

  if (reminder.recurrence.mode === "weeklyWeekday") {
    return `every ${capitalize(reminder.recurrence.weekday)}`;
  }

  const interval = reminder.recurrence.interval || "daily";
  if (interval === "daily") {
    return "every day";
  }
  if (interval === "allDays") {
    return "all days of the week";
  }
  if (interval === "weekdays") {
    return "weekdays";
  }
  if (interval === "weekly") {
    return "every week";
  }
  return "every month";
}

function getReminderDeliveryLabel(reminder) {
  if (reminder.delivery === "teams") {
    return "Teams";
  }
  if (reminder.delivery === "sms") {
    return "Text message";
  }
  if (reminder.delivery === "teamsSms") {
    return "Teams + text";
  }
  if (reminder.delivery === "both") {
    return "In app + text";
  }
  return "In app";
}

function inferReminderDelivery(text) {
  const lower = String(text || "").toLowerCase();
  if (/\b(teams and text|teams plus text)\b/.test(lower)) {
    return "teamsSms";
  }
  if (/\b(send to teams|notify me in teams|teams reminder|teams notification)\b/.test(lower)) {
    return "teams";
  }
  if (/\b(text me and remind me|send me a text and remind me|both)\b/.test(lower)) {
    return "both";
  }
  if (/\b(text me|send (me )?a text|sms me|message me)\b/.test(lower)) {
    return "sms";
  }
  return "inApp";
}

function normalizePhoneNumber(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (!digits) {
    return "";
  }
  if (digits.length === 10) {
    return `+1${digits}`;
  }
  if (digits.startsWith("1") && digits.length === 11) {
    return `+${digits}`;
  }
  if (value.trim().startsWith("+")) {
    return value.trim();
  }
  return `+${digits}`;
}

function capitalize(value) {
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : "";
}

function toLocalDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function nextWeekdayDate(baseDate, targetDayIndex, forceNextWeek) {
  const date = new Date(baseDate);
  const currentDayIndex = date.getDay();
  let diff = (targetDayIndex - currentDayIndex + 7) % 7;

  if (diff === 0 || forceNextWeek) {
    diff += 7;
  }

  date.setDate(date.getDate() + diff);
  return date;
}

function thisWeekendDate(baseDate) {
  const date = new Date(baseDate);
  const day = date.getDay();

  if (day === 6 || day === 0) {
    return date;
  }

  return nextWeekdayDate(baseDate, 6, false);
}

function toTitleCase(value) {
  return value.replace(/\w\S*/g, (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase());
}

function truncate(value, maxLength) {
  return value.length <= maxLength ? value : `${value.slice(0, maxLength - 1)}...`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeCsv(value) {
  return String(value).replaceAll('"', '""');
}

function safeFilename(value) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "hal-export";
}

function on(element, eventName, handler) {
  if (element) {
    element.addEventListener(eventName, handler);
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function persist() {
  localStorage.setItem(STORAGE_KEYS.app, JSON.stringify(buildPersistedState()));
  scheduleServerStateSync();
}

function normalizeTaskLists(taskLists, fallbackId) {
  const lists = Array.isArray(taskLists) && taskLists.length
    ? taskLists.map((list, index) => ({
      id: list.id || crypto.randomUUID(),
      name: index === 0
        ? ((list.name || "Regular Tasks").replace(/^My Tasks$/i, "Regular Tasks"))
        : (list.name || `Task List ${index + 1}`),
    }))
    : [{ id: fallbackId || crypto.randomUUID(), name: "Regular Tasks" }];

  if (!lists[0]) {
    lists[0] = { id: fallbackId || crypto.randomUUID(), name: "Regular Tasks" };
  }
  return lists;
}

function normalizeTasks(tasks, taskLists, defaultListId) {
  const validListIds = new Set((taskLists || []).map((list) => list.id));
  return Array.isArray(tasks)
    ? tasks.map((task) => ({
      ...task,
      listId: validListIds.has(task.listId) ? task.listId : defaultListId,
      dueDate: task.dueDate || "",
      dueTime: task.dueTime || "",
      notes: typeof task.notes === "string" ? task.notes : "",
      highlighted: Boolean(task.highlighted),
    }))
    : [];
}

function loadState() {
  const fallback = {
    theme: "dark",
    voiceResponsesEnabled: true,
    quote: "Build the life and systems you want to live inside.",
    quickLinks: [
      { id: crypto.randomUUID(), label: "EX", url: "https://www.office.com/launch/excel" },
      { id: crypto.randomUUID(), label: "PP", url: "https://www.office.com/launch/powerpoint" },
      { id: crypto.randomUUID(), label: "OL", url: "https://outlook.office.com/calendar/" },
      { id: crypto.randomUUID(), label: "PB", url: "https://app.powerbi.com/" },
    ],
    taskLists: [{ id: crypto.randomUUID(), name: "Regular Tasks" }],
    selectedTaskListId: "",
    showCompletedTasks: false,
    tasks: [],
    ideas: [],
    meetingNotes: [],
    calendarEvents: [],
    calendarSettings: {
      provider: "csv",
      googleMode: "both",
      googleWorkCalendarName: "Work",
      googlePersonalCalendarName: "Personal",
      uploadedCsvContent: "",
      uploadedCsvName: "",
      uploadedCsvImportedAt: "",
    },
    lastCalendarSyncDate: "",
    reminders: [],
    notificationSettings: {
      teamsEnabled: false,
      smsEnabled: false,
      phoneNumber: "",
    },
    followUps: [],
    audioInbox: [],
    selected: { type: null, id: null },
    lastCreated: { type: null, id: null },
    _meta: null,
    halMemory: createDefaultHalMemory(),
    micMode: "unknown",
    speechRecognition: null,
    mediaRecorder: null,
    recording: null,
  };

  try {
    const raw = localStorage.getItem(STORAGE_KEYS.app);
    if (!raw) {
      fallback.selectedTaskListId = fallback.taskLists[0].id;
      return fallback;
    }
    const parsed = JSON.parse(raw);
    return normalizeLoadedState(parsed, fallback);
  } catch {
    fallback.selectedTaskListId = fallback.taskLists[0].id;
    return fallback;
  }
}

function normalizeLoadedState(parsed = {}, fallback = null) {
  const defaults = fallback || {
    theme: "dark",
    voiceResponsesEnabled: true,
    quote: "Build the life and systems you want to live inside.",
    quickLinks: [],
    taskLists: [{ id: crypto.randomUUID(), name: "Regular Tasks" }],
    selectedTaskListId: "",
    showCompletedTasks: false,
    tasks: [],
    ideas: [],
    meetingNotes: [],
    calendarEvents: [],
    calendarSettings: {
      provider: "csv",
      googleMode: "both",
      googleWorkCalendarName: "Work",
      googlePersonalCalendarName: "Personal",
    },
    lastCalendarSyncDate: "",
    reminders: [],
    notificationSettings: {
      teamsEnabled: false,
      smsEnabled: false,
      phoneNumber: "",
    },
    followUps: [],
    audioInbox: [],
    selected: { type: null, id: null },
    lastCreated: { type: null, id: null },
    halMemory: createDefaultHalMemory(),
    micMode: "unknown",
    speechRecognition: null,
    mediaRecorder: null,
    recording: null,
  };

  const normalizedTaskLists = normalizeTaskLists(parsed.taskLists, defaults.taskLists[0].id);
  const normalizedTasks = normalizeTasks(parsed.tasks, normalizedTaskLists, normalizedTaskLists[0].id);
  return {
    ...defaults,
    ...parsed,
    theme: parsed.theme || defaults.theme,
    voiceResponsesEnabled: parsed.voiceResponsesEnabled !== false,
    taskLists: normalizedTaskLists,
    selectedTaskListId: parsed.selectedTaskListId || normalizedTaskLists[0]?.id || defaults.taskLists[0].id,
    showCompletedTasks: Boolean(parsed.showCompletedTasks),
    notificationSettings: {
      teamsEnabled: Boolean(parsed.notificationSettings?.teamsEnabled),
      smsEnabled: Boolean(parsed.notificationSettings?.smsEnabled),
      phoneNumber: normalizePhoneNumber(parsed.notificationSettings?.phoneNumber || ""),
    },
    calendarSettings: {
      provider: "csv",
      googleMode: parsed.calendarSettings?.googleMode || "both",
      googleWorkCalendarName: parsed.calendarSettings?.googleWorkCalendarName || "Work",
      googlePersonalCalendarName: parsed.calendarSettings?.googlePersonalCalendarName || "Personal",
      uploadedCsvContent: parsed.calendarSettings?.uploadedCsvContent || "",
      uploadedCsvName: parsed.calendarSettings?.uploadedCsvName || "",
      uploadedCsvImportedAt: parsed.calendarSettings?.uploadedCsvImportedAt || "",
    },
    calendarEvents: Array.isArray(parsed.calendarEvents)
      ? parsed.calendarEvents.map((event) => ({
        ...event,
        source: event.source || "manual",
      }))
      : [],
    lastCalendarSyncDate: String(parsed.lastCalendarSyncDate || ""),
    tasks: normalizedTasks,
    audioInbox: Array.isArray(parsed.audioInbox) ? parsed.audioInbox : [],
    lastCreated: parsed.lastCreated || { type: null, id: null },
    _meta: parsed._meta || null,
    halMemory: normalizeHalMemory(parsed.halMemory),
    micMode: "unknown",
    speechRecognition: null,
    mediaRecorder: null,
    recording: null,
  };
}
