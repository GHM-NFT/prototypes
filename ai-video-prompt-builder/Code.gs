const AI_VIDEO_CONFIG = Object.freeze({
  builderSheet: '00_Live_Builder',
  experimentSheet: '04_Experiments',
  archiveSheet: '07_Final_Prompts',
  namedRanges: Object.freeze({
    final: 'PROMPT_FINAL',
    canonical: 'CANONICAL_BRIEF',
    midjourney: 'MIDJOURNEY_PROMPT',
    grok: 'GROK_PROMPT',
    video: 'VIDEO_PROMPT',
    avatar: 'HEDRA_HEYGEN_PROMPT',
    depth: 'IMMERSITY_PROMPT',
  }),
  toolUrls: Object.freeze({
    'ChatGPT': 'https://chatgpt.com/',
    'Canonical Brief': 'https://chatgpt.com/',
    'Midjourney': 'https://www.midjourney.com/imagine',
    'Grok Imagine': 'https://grok.com/',
    'Veo': 'https://labs.google/fx/tools/flow',
    'Kling': 'https://klingai.com/',
    'Luma Dream Machine': 'https://lumalabs.ai/app',
    'Runway': 'https://app.runwayml.com/',
    'HeyGen': 'https://app.heygen.com/',
    'Hedra': 'https://www.hedra.com/',
    'Immersity': 'https://www.immersity.ai/',
    'LeiaPix': 'https://www.immersity.ai/',
    'Adobe Firefly': 'https://firefly.adobe.com/',
    'Pika': 'https://pika.art/',
    'DeeVid': 'https://deevid.ai/',
  }),
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Video Tools')
    .addItem('Open Prompt Panel', 'showPromptSidebar')
    .addItem('Open Selected Tool', 'openSelectedTool')
    .addSeparator()
    .addItem('Archive Current Prompt', 'archiveCurrentPrompt')
    .addItem('Create Experiment Draft', 'createExperimentDraft')
    .addToUi();
}

function showPromptSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI Video Prompt Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPromptPayload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const builder = requireSheet_(ss, AI_VIDEO_CONFIG.builderSheet);
  const tool = builder.getRange('B5').getDisplayValue().trim();

  const prompts = {};
  Object.entries(AI_VIDEO_CONFIG.namedRanges).forEach(([key, rangeName]) => {
    prompts[key] = readNamedRange_(ss, rangeName);
  });

  return {
    tool,
    targetUrl: resolveToolUrl_(tool),
    prompts,
    generatedAt: Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    ),
  };
}

function openSelectedTool() {
  const payload = getPromptPayload();
  const safeUrl = JSON.stringify(payload.targetUrl);
  const html = HtmlService.createHtmlOutput(
    `<script>window.open(${safeUrl}, '_blank', 'noopener,noreferrer');google.script.host.close();</script>`
  ).setWidth(120).setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening tool');
}

function archiveCurrentPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const builder = requireSheet_(ss, AI_VIDEO_CONFIG.builderSheet);
  const archive = requireSheet_(ss, AI_VIDEO_CONFIG.archiveSheet);
  const tool = builder.getRange('B5').getDisplayValue().trim();
  const prompt = readNamedRange_(ss, AI_VIDEO_CONFIG.namedRanges.final);
  const row = Math.max(4, archive.getLastRow() + 1);
  const now = new Date();
  const sequence = String(row - 3).padStart(4, '0');
  const id = `PRM-${Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMdd')}-${sequence}`;

  archive.getRange(row, 1, 1, 11).setValues([[
    id,
    now,
    '',
    tool,
    'Current adapter',
    prompt,
    '',
    '',
    'Draft',
    'Archived from 00_Live_Builder via Apps Script.',
    activeUser_(),
  ]]);

  archive.getRange(row, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  SpreadsheetApp.getUi().alert(`Archived as ${id}.`);
}

function createExperimentDraft() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const builder = requireSheet_(ss, AI_VIDEO_CONFIG.builderSheet);
  const log = requireSheet_(ss, AI_VIDEO_CONFIG.experimentSheet);
  const row = Math.max(4, log.getLastRow() + 1);
  const now = new Date();
  const sequence = String(row - 3).padStart(4, '0');
  const id = `EXP-${Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMdd')}-${sequence}`;

  const values = [[
    id,
    now,
    '',
    builder.getRange('B8').getDisplayValue(),
    builder.getRange('B5').getDisplayValue(),
    '',
    builder.getRange('B6').getDisplayValue(),
    readNamedRange_(ss, AI_VIDEO_CONFIG.namedRanges.final),
    `Seed: ${builder.getRange('B25').getDisplayValue()}`,
    builder.getRange('B21').getDisplayValue(),
    builder.getRange('B22').getDisplayValue(),
    '',
    '',
    '',
    '',
    '',
    'Needs Retest',
    'No',
    'Draft created from the live builder. Add output link and evaluation after generation.',
    activeUser_(),
  ]];

  log.getRange(row, 1, 1, 20).setValues(values);
  log.getRange(row, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  SpreadsheetApp.getUi().alert(`Experiment draft ${id} created.`);
}

function resolveToolUrl_(tool) {
  return AI_VIDEO_CONFIG.toolUrls[tool] || AI_VIDEO_CONFIG.toolUrls.ChatGPT;
}

function readNamedRange_(ss, rangeName) {
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    throw new Error(`Missing named range: ${rangeName}`);
  }
  const value = range.getDisplayValues().flat().find((cell) => String(cell).trim() !== '');
  return value ? String(value) : '';
}

function requireSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Missing sheet: ${sheetName}`);
  }
  return sheet;
}

function activeUser_() {
  const email = Session.getActiveUser().getEmail();
  return email || 'Current user';
}
