const express = require('express');
const path = require('path');
const { google } = require('googleapis');

function createApiApp(options = {}) {
  const app = express();
  const io = options.io || null;

  app.use(express.json());
  app.use((req, res, next) => {
    if (req.url === '/api') {
      req.url = '/';
    } else if (req.url.startsWith('/api/')) {
      req.url = req.url.slice(4);
    }

    next();
  });

  function emitRealtime(eventName, payload) {
    if (io) {
      io.emit(eventName, payload);
    }
  }

let players = [];
let lastLogText = '';
let logHistory = [];
let playerSheetReady = false;
let playerSheetReadyPromise = null;
let logSheetReady = false;
let rateSheetReady = false;
let blackjackSessionSheetReady = false;
let baccaratSessionSheetReady = false;
let russianRouletteSessionSheetReady = false;
let pvpSessionSheetReady = false;
let formulaHighLowSessionSheetReady = false;
let settingsSheetReady = false;
let blackjackSessions = new Map();
let baccaratSessions = new Map();
let russianRouletteSessions = new Map();
let pvpSessions = new Map();
let formulaHighLowSessions = new Map();
const COLORS = ['red', 'blue', 'green', 'yellow', 'white'];
const CHIP_LABELS = {
  red: 'Red',
  blue: 'Blue',
  green: 'Green',
  yellow: 'Yellow',
  white: 'White'
};
const GAME_LIMITS = [
  { type: 'roulette', key: 'rouletteLimit', label: 'Roulette' },
  { type: 'highlow', key: 'highlowLimit', label: 'High Low' },
  { type: 'baccarat', key: 'baccaratLimit', label: 'Baccarat' },
  { type: 'blackjack', key: 'blackjackLimit', label: 'Blackjack' },
  { type: 'redblack', key: 'redblackLimit', label: 'Red Black' }
];
const GAME_LIMIT_BY_TYPE = new Map(GAME_LIMITS.map((config) => [config.type, config]));
const DEFAULT_DAILY_GAME_LIMIT = 3;
const GAME_LIMIT_RESET_TIME_ZONE = 'Asia/Seoul';
const GAME_LIMIT_RESET_KEY = 'lastGameLimitResetDate';
const DEFAULT_RATES = { red: 1, blue: 3, green: 5, yellow: 10, white: 15 };
const READ_CACHE_TTL_MS = 3000;
const ROW_NUMBER_CACHE_TTL_MS = 60000;

let rates = { ...DEFAULT_RATES };
let playersLoadedAt = 0;
let playersCacheInvalidatedAt = 0;
let playersLoadPromise = null;
let playerRowNumbers = new Map();
let playerRowNumbersLoadedAt = 0;
let ratesLoadedAt = 0;
let ratesLoadPromise = null;
let dailyGameLimitResetDate = '';
let dailyGameLimitResetPromise = null;

// =========================
// 🔑 Google Sheets 설정
// =========================
function normalizePrivateKey(value) {
  let key = String(value || '').trim();

  if (
    (key.startsWith('"') && key.endsWith('"')) ||
    (key.startsWith("'") && key.endsWith("'"))
  ) {
    key = key.slice(1, -1);
  }

  return key.replace(/\\n/g, '\n');
}

function getGoogleAuthConfig() {
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON?.trim();
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL?.trim();
  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.trim();

  if (serviceAccountJson) {
    return {
      credentials: JSON.parse(serviceAccountJson),
      scopes
    };
  }

  if (clientEmail && privateKey) {
    return {
      credentials: {
        client_email: clientEmail,
        private_key: normalizePrivateKey(privateKey)
      },
      scopes
    };
  }

  if (process.env.VERCEL) {
    throw new Error('Google Sheets 인증 환경변수가 없습니다. Vercel에 GOOGLE_SERVICE_ACCOUNT_JSON 또는 GOOGLE_CLIENT_EMAIL/GOOGLE_PRIVATE_KEY를 설정하세요.');
  }

  return {
    keyFile: path.join(process.cwd(), 'key.json'),
    scopes
  };
}

const auth = new google.auth.GoogleAuth(getGoogleAuthConfig());

const sheets = google.sheets({ version: 'v4', auth });

// 네 시트 ID
const SPREADSHEET_ID = '1d9s84o9LrVdncnWCNC85Vjnml4uyjp-rkMgqXkfesls';

const PLAYER_SHEET_NAME = '플레이어';
const PLAYER_SHEET_REF = `'${PLAYER_SHEET_NAME}'`;
const PLAYER_END_COLUMN = 'M';
const PLAYER_RANGE = `${PLAYER_SHEET_REF}!A:${PLAYER_END_COLUMN}`;
const PLAYER_MIGRATION_RANGE = `${PLAYER_SHEET_REF}!A:N`;
const PLAYER_WRITE_RANGE = `${PLAYER_SHEET_REF}!A1`;
const PLAYER_ID_RANGE = `${PLAYER_SHEET_REF}!A:A`;
const PLAYER_HEADER_RANGE = `${PLAYER_SHEET_REF}!A1:${PLAYER_END_COLUMN}1`;
const PLAYER_GAME_LIMIT_RESET_RANGE_START_COLUMN = 'I';
const PLAYER_GAME_LIMIT_RESET_RANGE_END_COLUMN = 'M';
const PLAYER_HEADER = ['id', 'name', 'team', ...COLORS, ...GAME_LIMITS.map((game) => game.key)];
const SETTINGS_SHEET_NAME = '설정';
const SETTINGS_SHEET_REF = `'${SETTINGS_SHEET_NAME}'`;
const SETTINGS_HEADER_RANGE = `${SETTINGS_SHEET_REF}!A1:B1`;
const SETTINGS_RANGE = `${SETTINGS_SHEET_REF}!A:B`;
const SETTINGS_HEADER = ['key', 'value'];
const RATE_SHEET_NAME = '칩가치';
const RATE_SHEET_REF = `'${RATE_SHEET_NAME}'`;
const RATE_HEADER_RANGE = `${RATE_SHEET_REF}!A1:B1`;
const RATE_RANGE = `${RATE_SHEET_REF}!A:B`;
const RATE_WRITE_RANGE = `${RATE_SHEET_REF}!A1`;
const RATE_HEADER = ['color', 'value'];
const LOG_SHEET_NAME = '로그';
const LOG_SHEET_REF = `'${LOG_SHEET_NAME}'`;
const LOG_HEADER_RANGE = `${LOG_SHEET_REF}!A1:F1`;
const LOG_RANGE = `${LOG_SHEET_REF}!A:F`;
const LOG_ID_RANGE = `${LOG_SHEET_REF}!A:A`;
const LOG_APPEND_RANGE = `${LOG_SHEET_REF}!A:F`;
const LOG_HEADER = ['id', 'type', 'text', 'publicText', 'createdAt', 'groupId'];
const DEFAULT_LOG_LIMIT = 100;
const MAX_LOG_LIMIT = 1000;
const BLACKJACK_SESSION_SHEET_NAME = '블랙잭세션';
const BLACKJACK_SESSION_SHEET_REF = `'${BLACKJACK_SESSION_SHEET_NAME}'`;
const BLACKJACK_SESSION_HEADER_RANGE = `${BLACKJACK_SESSION_SHEET_REF}!A1:L1`;
const BLACKJACK_SESSION_RANGE = `${BLACKJACK_SESSION_SHEET_REF}!A:L`;
const BLACKJACK_SESSION_ID_RANGE = `${BLACKJACK_SESSION_SHEET_REF}!A:A`;
const BLACKJACK_SESSION_HEADER = [
  'playerId',
  'color',
  'bet',
  'deck',
  'playerCards',
  'dealerCards',
  'lastDraw',
  'dealerDraws',
  'done',
  'result',
  'updatedAt',
  'logGroupId'
];
const BACCARAT_SESSION_SHEET_NAME = '바카라세션';
const BACCARAT_SESSION_SHEET_REF = `'${BACCARAT_SESSION_SHEET_NAME}'`;
const BACCARAT_SESSION_HEADER_RANGE = `${BACCARAT_SESSION_SHEET_REF}!A1:N1`;
const BACCARAT_SESSION_RANGE = `${BACCARAT_SESSION_SHEET_REF}!A:N`;
const BACCARAT_SESSION_ID_RANGE = `${BACCARAT_SESSION_SHEET_REF}!A:A`;
const BACCARAT_SESSION_HEADER = [
  'playerId',
  'side',
  'color',
  'bet',
  'shoe',
  'playerCards',
  'bankerCards',
  'playerThirdCard',
  'bankerThirdCard',
  'playerAction',
  'done',
  'outcome',
  'updatedAt',
  'logGroupId'
];
const RUSSIAN_ROULETTE_SESSION_SHEET_NAME = '러시안룰렛세션';
const RUSSIAN_ROULETTE_SESSION_SHEET_REF = `'${RUSSIAN_ROULETTE_SESSION_SHEET_NAME}'`;
const RUSSIAN_ROULETTE_SESSION_HEADER_RANGE = `${RUSSIAN_ROULETTE_SESSION_SHEET_REF}!A1:M1`;
const RUSSIAN_ROULETTE_SESSION_RANGE = `${RUSSIAN_ROULETTE_SESSION_SHEET_REF}!A:M`;
const RUSSIAN_ROULETTE_SESSION_ID_RANGE = `${RUSSIAN_ROULETTE_SESSION_SHEET_REF}!A:A`;
const RUSSIAN_ROULETTE_SESSION_HEADER = [
  'id',
  'participantIds',
  'activeIds',
  'eliminated',
  'color',
  'bet',
  'pot',
  'round',
  'done',
  'winnerId',
  'result',
  'lastAction',
  'updatedAt'
];
const PVP_SESSION_SHEET_NAME = 'PVP세션';
const PVP_SESSION_SHEET_REF = `'${PVP_SESSION_SHEET_NAME}'`;
const PVP_SESSION_HEADER_RANGE = `${PVP_SESSION_SHEET_REF}!A1:Q1`;
const PVP_SESSION_RANGE = `${PVP_SESSION_SHEET_REF}!A:Q`;
const PVP_SESSION_ID_RANGE = `${PVP_SESSION_SHEET_REF}!A:A`;
const PVP_SESSION_HEADER = [
  'id',
  'gameType',
  'playerOneId',
  'playerTwoId',
  'color',
  'bet',
  'deck',
  'shoe',
  'playerOneCards',
  'playerTwoCards',
  'playerOneAction',
  'playerTwoAction',
  'turn',
  'done',
  'winnerKey',
  'result',
  'updatedAt'
];
const FORMULA_HIGH_LOW_SESSION_SHEET_NAME = '콤비네이션베팅세션';
const FORMULA_HIGH_LOW_SESSION_SHEET_REF = `'${FORMULA_HIGH_LOW_SESSION_SHEET_NAME}'`;
const FORMULA_HIGH_LOW_SESSION_HEADER_RANGE = `${FORMULA_HIGH_LOW_SESSION_SHEET_REF}!A1:N1`;
const FORMULA_HIGH_LOW_SESSION_RANGE = `${FORMULA_HIGH_LOW_SESSION_SHEET_REF}!A:N`;
const FORMULA_HIGH_LOW_SESSION_ID_RANGE = `${FORMULA_HIGH_LOW_SESSION_SHEET_REF}!A:A`;
const FORMULA_HIGH_LOW_SESSION_HEADER = [
  'id',
  'participantIds',
  'color',
  'bet',
  'target',
  'operatorCards',
  'participantStates',
  'stage',
  'winnerIds',
  'result',
  'memo',
  'pot',
  'roundNumber',
  'updatedAt'
];

// =========================
// 📥 시트 → 서버
// =========================
function withPlayerRowNumber(player, rowNumber) {
  Object.defineProperty(player, '_rowNumber', {
    value: rowNumber,
    enumerable: false,
    writable: true
  });
  return player;
}

function normalizeHeaderCell(value) {
  return String(value || '').trim();
}

function headerMatches(header, expectedHeader) {
  return expectedHeader.every((name, index) => normalizeHeaderCell(header[index]) === name);
}

function getRowValueByHeader(row, header, name) {
  const index = header.findIndex((cell) => normalizeHeaderCell(cell) === name);
  return index >= 0 ? row[index] : undefined;
}

function parsePlayerRowByHeader(row, header, rowNumber) {
  const normalizedHeader = header.map(normalizeHeaderCell);
  const hasNamedHeader = ['id', 'name'].every((name) => normalizedHeader.includes(name));

  if (hasNamedHeader) {
    const limitValues = Object.fromEntries(
      GAME_LIMITS.map((game) => [game.key, readGameLimitValue(getRowValueByHeader(row, header, game.key))])
    );

    return withPlayerRowNumber({
      id: Number(getRowValueByHeader(row, header, 'id') || 0),
      name: getRowValueByHeader(row, header, 'name') || '',
      team: getRowValueByHeader(row, header, 'team') || '',
      ...Object.fromEntries(COLORS.map((color) => [color, readChipValue(getRowValueByHeader(row, header, color))])),
      ...limitValues
    }, rowNumber);
  }

  return withPlayerRowNumber({
    id: Number(row[0] || 0),
    name: row[1] || '',
    team: '',
    red: readChipValue(row[2]),
    blue: readChipValue(row[3]),
    green: readChipValue(row[4]),
    yellow: readChipValue(row[5]),
    white: readChipValue(row[6]),
    ...Object.fromEntries(GAME_LIMITS.map((game) => [game.key, null]))
  }, rowNumber);
}

function parsePlayerRow(row, rowNumber) {
  const limitValues = Object.fromEntries(
    GAME_LIMITS.map((game, index) => [game.key, readGameLimitValue(row[8 + index])])
  );

  return withPlayerRowNumber({
    id: Number(row[0] || 0),
    name: row[1] || '',
    team: row[2] || '',
    red: readChipValue(row[3]),
    blue: readChipValue(row[4]),
    green: readChipValue(row[5]),
    yellow: readChipValue(row[6]),
    white: readChipValue(row[7]),
    ...limitValues
  }, rowNumber);
}

function playerToRow(player) {
  return [
    player.id,
    player.name,
    player.team || '',
    ...COLORS.map((color) => player[color] || 0),
    ...GAME_LIMITS.map((game) => formatGameLimitValue(player[game.key]))
  ];
}

function uniquePlayerIds(playerIds) {
  return [...new Set(playerIds.map(Number))]
    .filter((id) => Number.isFinite(id));
}

function isReadCacheFresh(loadedAt) {
  return loadedAt > 0 && Date.now() - loadedAt < READ_CACHE_TTL_MS;
}

function isRowNumberCacheFresh() {
  return playerRowNumbersLoadedAt > 0 && Date.now() - playerRowNumbersLoadedAt < ROW_NUMBER_CACHE_TTL_MS;
}

function rememberPlayerRowNumber(id, rowNumber) {
  const playerId = Number(id);
  const row = Number(rowNumber);

  if (Number.isFinite(playerId) && Number.isFinite(row) && row > 1) {
    playerRowNumbers.set(playerId, row);
  }
}

function rememberPlayerRowNumbers(nextPlayers, options = {}) {
  if (options.replace) {
    playerRowNumbers = new Map();
  }

  nextPlayers.forEach((player) => {
    rememberPlayerRowNumber(player.id, player._rowNumber);
  });
  playerRowNumbersLoadedAt = Date.now();
}

function invalidatePlayersCache() {
  playersLoadedAt = 0;
  playersCacheInvalidatedAt = Date.now();
}

function markPlayersCacheFresh() {
  playersLoadedAt = Date.now();
}

function markRatesCacheFresh() {
  ratesLoadedAt = Date.now();
}

async function ensurePlayerSheetHeader() {
  if (playerSheetReady) {
    return;
  }

  if (playerSheetReadyPromise) {
    return playerSheetReadyPromise;
  }

  playerSheetReadyPromise = (async () => {
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      fields: 'sheets.properties.title'
    });
    const sheetExists = (spreadsheet.data.sheets || [])
      .some((sheet) => sheet.properties?.title === PLAYER_SHEET_NAME);

    if (!sheetExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: { title: PLAYER_SHEET_NAME }
              }
            }
          ]
        }
      });
    }

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: PLAYER_MIGRATION_RANGE
    });
    const values = res.data.values || [];
    const header = values[0] || [];
    const hasHeader = headerMatches(header, PLAYER_HEADER);

    if (!hasHeader) {
      const migratedRows = [
        PLAYER_HEADER,
        ...values
          .slice(1)
          .filter((row) => row.some((cell) => normalizeHeaderCell(cell)))
          .map((row, index) => playerToRow(parsePlayerRowByHeader(row, header, index + 2)))
      ];

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: PLAYER_WRITE_RANGE,
        valueInputOption: 'RAW',
        requestBody: { values: migratedRows }
      });

      await sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PLAYER_SHEET_REF}!N:N`
      });
    } else if (normalizeHeaderCell(header[13]) === 'russianrouletteLimit') {
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `${PLAYER_SHEET_REF}!N:N`
      });
    }

    playerSheetReady = true;
  })();

  try {
    await playerSheetReadyPromise;
  } finally {
    playerSheetReadyPromise = null;
  }
}

async function ensureSettingsSheetHeader() {
  if (settingsSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === SETTINGS_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: SETTINGS_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: SETTINGS_HEADER_RANGE
  });
  const header = (res.data.values || [])[0] || [];

  if (!headerMatches(header, SETTINGS_HEADER)) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: SETTINGS_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [SETTINGS_HEADER] }
    });
  }

  settingsSheetReady = true;
}

async function readSettingValue(key) {
  await ensureSettingsSheetHeader();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: SETTINGS_RANGE
  });
  const rows = res.data.values || [];
  const match = rows.find((row, index) => index > 0 && normalizeHeaderCell(row[0]) === key);

  return match?.[1] || '';
}

async function saveSettingValue(key, value) {
  await ensureSettingsSheetHeader();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: SETTINGS_RANGE
  });
  const rows = res.data.values || [];
  const rowIndex = rows.findIndex((row, index) => index > 0 && normalizeHeaderCell(row[0]) === key);

  if (rowIndex >= 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SETTINGS_SHEET_REF}!A${rowIndex + 1}:B${rowIndex + 1}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[key, value]] }
    });
    return;
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: SETTINGS_RANGE,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [[key, value]] }
  });
}

async function getLastPlayerDataRowNumber() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_ID_RANGE
  });
  let lastRowNumber = 1;

  (res.data.values || []).forEach((row, index) => {
    if (index > 0 && normalizeHeaderCell(row[0])) {
      lastRowNumber = index + 1;
    }
  });

  return lastRowNumber;
}

async function resetDailyGameLimitRange() {
  const lastRowNumber = await getLastPlayerDataRowNumber();

  if (lastRowNumber < 2) {
    return false;
  }

  const rowCount = lastRowNumber - 1;
  const values = Array.from(
    { length: rowCount },
    () => GAME_LIMITS.map(() => DEFAULT_DAILY_GAME_LIMIT)
  );

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${PLAYER_SHEET_REF}!${PLAYER_GAME_LIMIT_RESET_RANGE_START_COLUMN}2:${PLAYER_GAME_LIMIT_RESET_RANGE_END_COLUMN}${lastRowNumber}`,
    valueInputOption: 'RAW',
    requestBody: { values }
  });

  players.forEach((player) => {
    GAME_LIMITS.forEach((game) => {
      player[game.key] = DEFAULT_DAILY_GAME_LIMIT;
    });
  });

  invalidatePlayersCache();
  return true;
}

async function ensureDailyGameLimitReset() {
  const today = getKoreaDateKey();

  if (dailyGameLimitResetDate === today) {
    return false;
  }

  if (dailyGameLimitResetPromise) {
    return dailyGameLimitResetPromise;
  }

  dailyGameLimitResetPromise = (async () => {
    const [lastResetDate] = await Promise.all([
      readSettingValue(GAME_LIMIT_RESET_KEY),
      ensurePlayerSheetHeader()
    ]);

    if (lastResetDate === today) {
      dailyGameLimitResetDate = today;
      return false;
    }

    await resetDailyGameLimitRange();
    await saveSettingValue(GAME_LIMIT_RESET_KEY, today);
    dailyGameLimitResetDate = today;
    return true;
  })()
    .finally(() => {
      dailyGameLimitResetPromise = null;
    });

  return dailyGameLimitResetPromise;
}

async function loadSheet() {
  await ensurePlayerSheetHeader();
  await ensureDailyGameLimitReset();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_RANGE
  });

  const rows = res.data.values || [];

  if (rows.length <= 1) {
    players = [];
    rememberPlayerRowNumbers(players, { replace: true });
    return;
  }

  players = rows.slice(1).map((row, index) => parsePlayerRow(row, index + 2));
  rememberPlayerRowNumbers(players, { replace: true });
}

async function loadSheetCached() {
  if (isReadCacheFresh(playersLoadedAt)) {
    return players;
  }

  if (playersLoadPromise) {
    return playersLoadPromise;
  }

  const loadStartedAt = Date.now();

  playersLoadPromise = loadSheet()
    .then(() => {
      if (playersCacheInvalidatedAt <= loadStartedAt) {
        markPlayersCacheFresh();
      }

      return players;
    })
    .finally(() => {
      playersLoadPromise = null;
    });

  return playersLoadPromise;
}

// =========================
// 📤 서버 → 시트
// =========================
async function saveSheet() {
  await ensurePlayerSheetHeader();

  const values = [
    PLAYER_HEADER,
    ...players.map(playerToRow)
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_WRITE_RANGE,
    valueInputOption: 'RAW',
    requestBody: { values }
  });

  invalidatePlayersCache();
}

async function getPlayerRowNumbers(playerIds, options = {}) {
  await ensurePlayerSheetHeader();

  const ids = uniquePlayerIds(playerIds);
  const idSet = new Set(ids);
  const rowNumbers = new Map();

  if (!ids.length) {
    return rowNumbers;
  }

  if (!options.forceRefresh && isRowNumberCacheFresh()) {
    ids.forEach((id) => {
      const rowNumber = playerRowNumbers.get(id);

      if (rowNumber) {
        rowNumbers.set(id, rowNumber);
      }
    });

    if (rowNumbers.size === ids.length) {
      return rowNumbers;
    }
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_ID_RANGE
  });

  playerRowNumbers = new Map();
  (res.data.values || []).forEach((row, index) => {
    const id = Number(row[0]);

    if (index > 0 && Number.isFinite(id)) {
      const rowNumber = index + 1;
      playerRowNumbers.set(id, rowNumber);

      if (idSet.has(id) && !rowNumbers.has(id)) {
        rowNumbers.set(id, rowNumber);
      }
    }
  });
  playerRowNumbersLoadedAt = Date.now();

  return rowNumbers;
}

async function loadPlayersByIds(playerIds) {
  invalidatePlayersCache();

  const ids = uniquePlayerIds(playerIds);

  if (!ids.length) {
    players = [];
    return players;
  }

  await ensureDailyGameLimitReset();

  const canUseCachedRowNumbers = isRowNumberCacheFresh() && ids.every((id) => playerRowNumbers.get(id));
  let rowNumbers = await getPlayerRowNumbers(ids);
  let targets = ids
    .map((id) => ({ id, rowNumber: rowNumbers.get(id) }))
    .filter((target) => target.rowNumber);

  if (!targets.length) {
    players = [];
    return players;
  }

  async function readTargetPlayers() {
    const res = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: targets.map((target) => sheetRowRange(PLAYER_SHEET_REF, PLAYER_END_COLUMN, target.rowNumber))
    });

    return (res.data.valueRanges || [])
      .map((valueRange, index) => parsePlayerRow((valueRange.values || [])[0] || [], targets[index].rowNumber))
      .filter((player) => Number.isFinite(player.id) && targets.some((target) => target.id === player.id));
  }

  players = await readTargetPlayers();

  if (canUseCachedRowNumbers && players.length < targets.length) {
    rowNumbers = await getPlayerRowNumbers(ids, { forceRefresh: true });
    targets = ids
      .map((id) => ({ id, rowNumber: rowNumbers.get(id) }))
      .filter((target) => target.rowNumber);
    players = targets.length ? await readTargetPlayers() : [];
  }

  rememberPlayerRowNumbers(players);

  return players;
}

async function savePlayers(changedPlayers) {
  const byId = new Map();

  changedPlayers
    .filter(Boolean)
    .forEach((player) => {
      const id = Number(player.id);

      if (Number.isFinite(id)) {
        byId.set(id, player);
      }
    });

  const targetPlayers = [...byId.values()];

  if (!targetPlayers.length) {
    return;
  }

  const missingRowIds = targetPlayers
    .filter((player) => !player._rowNumber)
    .map((player) => player.id);
  const rowNumbers = await getPlayerRowNumbers(missingRowIds);

  const data = targetPlayers.map((player) => {
    const rowNumber = player._rowNumber || rowNumbers.get(Number(player.id));

    if (!rowNumber) {
      throw new Error(`플레이어 행을 찾을 수 없습니다: ${player.id}`);
    }

    return {
      range: sheetRowRange(PLAYER_SHEET_REF, PLAYER_END_COLUMN, rowNumber),
      values: [playerToRow(player)]
    };
  });

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'RAW',
      data
    }
  });

  data.forEach((entry, index) => {
    const rowNumber = getRowNumberFromRange(entry.range);
    rememberPlayerRowNumber(targetPlayers[index].id, rowNumber);
  });
  invalidatePlayersCache();
}

async function ensureLogSheetHeader() {
  if (logSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === LOG_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: LOG_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: LOG_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = LOG_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: LOG_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [LOG_HEADER] }
    });
  }

  logSheetReady = true;
}

async function ensureRateSheetHeader() {
  if (rateSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === RATE_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: RATE_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RATE_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = RATE_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: RATE_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [RATE_HEADER] }
    });
  }

  rateSheetReady = true;
}

async function loadRates() {
  await ensureRateSheetHeader();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RATE_RANGE
  });
  const rows = res.data.values || [];

  if (rows.length <= 1) {
    await saveRates(DEFAULT_RATES);
    return rates;
  }

  const nextRates = { ...DEFAULT_RATES };

  rows.slice(1).forEach((row) => {
    const color = String(row[0] || '').trim().toLowerCase();
    const value = Number(row[1]);

    if (COLORS.includes(color) && Number.isFinite(value) && value > 0) {
      nextRates[color] = value;
    }
  });

  rates = nextRates;
  markRatesCacheFresh();
  return rates;
}

async function loadRatesCached() {
  if (isReadCacheFresh(ratesLoadedAt)) {
    return rates;
  }

  if (ratesLoadPromise) {
    return ratesLoadPromise;
  }

  ratesLoadPromise = loadRates()
    .finally(() => {
      ratesLoadPromise = null;
    });

  return ratesLoadPromise;
}

async function saveRates(nextRates) {
  await ensureRateSheetHeader();

  rates = Object.fromEntries(
    COLORS.map((color) => [color, Number(nextRates[color])])
  );

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: RATE_WRITE_RANGE,
    valueInputOption: 'RAW',
    requestBody: {
      values: [
        RATE_HEADER,
        ...COLORS.map((color) => [color, rates[color]])
      ]
    }
  });

  markRatesCacheFresh();
  return rates;
}

async function ensureBlackjackSessionSheetHeader() {
  if (blackjackSessionSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === BLACKJACK_SESSION_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: BLACKJACK_SESSION_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: BLACKJACK_SESSION_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = BLACKJACK_SESSION_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: BLACKJACK_SESSION_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [BLACKJACK_SESSION_HEADER] }
    });
  }

  blackjackSessionSheetReady = true;
}

function parseJsonArray(value) {
  try {
    const parsed = JSON.parse(value || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function parseJsonObject(value) {
  try {
    const parsed = JSON.parse(value || '{}');
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (err) {
    return {};
  }
}

function parseIdArray(value) {
  return parseJsonArray(value)
    .map(Number)
    .filter((id) => Number.isFinite(id));
}

function parseEliminatedPlayers(value) {
  return parseJsonArray(value)
    .map((entry) => ({
      id: Number(entry?.id),
      round: Number(entry?.round)
    }))
    .filter((entry) => Number.isFinite(entry.id) && Number.isFinite(entry.round));
}

function sheetRowRange(sheetRef, endColumn, rowNumber) {
  return `${sheetRef}!A${rowNumber}:${endColumn}${rowNumber}`;
}

async function findRowNumberByFirstColumn(idRange, id) {
  const targetId = Number(id);

  if (!Number.isFinite(targetId)) {
    return null;
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: idRange
  });
  const rows = res.data.values || [];

  for (let index = 0; index < rows.length; index += 1) {
    if (Number(rows[index]?.[0]) === targetId) {
      return index + 1;
    }
  }

  return null;
}

async function findRowNumberByFirstColumnText(idRange, id) {
  const targetId = String(id || '');

  if (!targetId) {
    return null;
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: idRange
  });
  const rows = res.data.values || [];

  for (let index = 0; index < rows.length; index += 1) {
    if (String(rows[index]?.[0] || '') === targetId) {
      return index + 1;
    }
  }

  return null;
}

async function readSingleRow(range) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range
  });
  return (res.data.values || [])[0] || [];
}

async function updateSingleRow(range, row) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });
}

async function appendSingleRow(range, row) {
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
  return res.data.updates?.updatedRange || '';
}

async function clearSingleRow(range) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range
  });
}

function getRowNumberFromRange(range) {
  const match = String(range || '').match(/![A-Z]+(\d+):/);
  const rowNumber = match ? Number(match[1]) : NaN;
  return Number.isFinite(rowNumber) ? rowNumber : null;
}

function withSessionRowNumber(session, rowNumber) {
  Object.defineProperty(session, '_rowNumber', {
    value: rowNumber,
    enumerable: false,
    writable: true
  });
  return session;
}

function parseBlackjackSessionRow(row) {
  return {
    playerId: Number(row[0]),
    color: row[1] || '',
    bet: toChipAmount(row[2]),
    deck: parseJsonArray(row[3]),
    playerCards: parseJsonArray(row[4]),
    dealerCards: parseJsonArray(row[5]),
    lastDraw: row[6] || '',
    dealerDraws: parseJsonArray(row[7]),
    done: row[8] === 'true',
    result: row[9] || '',
    logGroupId: row[11] || ''
  };
}

function blackjackSessionToRow(session) {
  return [
    session.playerId,
    session.color,
    session.bet,
    JSON.stringify(session.deck || []),
    JSON.stringify(session.playerCards || []),
    JSON.stringify(session.dealerCards || []),
    session.lastDraw || '',
    JSON.stringify(session.dealerDraws || []),
    session.done ? 'true' : 'false',
    session.result || '',
    new Date().toISOString(),
    session.logGroupId || ''
  ];
}

async function getBlackjackSessionRowNumber(playerId) {
  await ensureBlackjackSessionSheetHeader();
  return findRowNumberByFirstColumn(BLACKJACK_SESSION_ID_RANGE, playerId);
}

async function getBlackjackSession(playerId) {
  const rowNumber = await getBlackjackSessionRowNumber(playerId);

  if (!rowNumber) {
    blackjackSessions.delete(String(playerId));
    return null;
  }

  const row = await readSingleRow(sheetRowRange(BLACKJACK_SESSION_SHEET_REF, 'L', rowNumber));
  const session = parseBlackjackSessionRow(row);

  if (!Number.isFinite(session.playerId)) {
    blackjackSessions.delete(String(playerId));
    return null;
  }

  withSessionRowNumber(session, rowNumber);
  blackjackSessions.set(String(session.playerId), session);
  return session;
}

async function saveBlackjackSession(session) {
  let rowNumber = session._rowNumber || await getBlackjackSessionRowNumber(session.playerId);
  const row = blackjackSessionToRow(session);

  if (rowNumber) {
    await updateSingleRow(sheetRowRange(BLACKJACK_SESSION_SHEET_REF, 'L', rowNumber), row);
  } else {
    const updatedRange = await appendSingleRow(BLACKJACK_SESSION_RANGE, row);
    rowNumber = getRowNumberFromRange(updatedRange);
  }

  if (rowNumber) {
    withSessionRowNumber(session, rowNumber);
  }
  blackjackSessions.set(String(session.playerId), session);
}

async function deleteBlackjackSession(playerId) {
  const cachedSession = blackjackSessions.get(String(playerId));
  const rowNumber = cachedSession?._rowNumber || await getBlackjackSessionRowNumber(playerId);

  if (rowNumber) {
    await clearSingleRow(sheetRowRange(BLACKJACK_SESSION_SHEET_REF, 'L', rowNumber));
  }

  blackjackSessions.delete(String(playerId));
}

async function ensureBaccaratSessionSheetHeader() {
  if (baccaratSessionSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === BACCARAT_SESSION_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: BACCARAT_SESSION_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: BACCARAT_SESSION_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = BACCARAT_SESSION_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: BACCARAT_SESSION_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [BACCARAT_SESSION_HEADER] }
    });
  }

  baccaratSessionSheetReady = true;
}

function parseBaccaratSessionRow(row) {
  return {
    playerId: Number(row[0]),
    side: row[1] || '',
    color: row[2] || '',
    bet: toChipAmount(row[3]),
    shoe: parseJsonArray(row[4]),
    playerCards: parseJsonArray(row[5]),
    bankerCards: parseJsonArray(row[6]),
    playerThirdCard: row[7] || '',
    bankerThirdCard: row[8] || '',
    playerAction: row[9] || '',
    done: row[10] === 'true',
    outcome: row[11] || '',
    logGroupId: row[13] || ''
  };
}

function baccaratSessionToRow(session) {
  return [
    session.playerId,
    session.side,
    session.color,
    session.bet,
    JSON.stringify(session.shoe || []),
    JSON.stringify(session.playerCards || []),
    JSON.stringify(session.bankerCards || []),
    session.playerThirdCard || '',
    session.bankerThirdCard || '',
    session.playerAction || '',
    session.done ? 'true' : 'false',
    session.outcome || '',
    new Date().toISOString(),
    session.logGroupId || ''
  ];
}

async function getBaccaratSessionRowNumber(playerId) {
  await ensureBaccaratSessionSheetHeader();
  return findRowNumberByFirstColumn(BACCARAT_SESSION_ID_RANGE, playerId);
}

async function getBaccaratSession(playerId) {
  const rowNumber = await getBaccaratSessionRowNumber(playerId);

  if (!rowNumber) {
    baccaratSessions.delete(String(playerId));
    return null;
  }

  const row = await readSingleRow(sheetRowRange(BACCARAT_SESSION_SHEET_REF, 'N', rowNumber));
  const session = parseBaccaratSessionRow(row);

  if (!Number.isFinite(session.playerId)) {
    baccaratSessions.delete(String(playerId));
    return null;
  }

  withSessionRowNumber(session, rowNumber);
  baccaratSessions.set(String(session.playerId), session);
  return session;
}

async function saveBaccaratSession(session) {
  let rowNumber = session._rowNumber || await getBaccaratSessionRowNumber(session.playerId);
  const row = baccaratSessionToRow(session);

  if (rowNumber) {
    await updateSingleRow(sheetRowRange(BACCARAT_SESSION_SHEET_REF, 'N', rowNumber), row);
  } else {
    const updatedRange = await appendSingleRow(BACCARAT_SESSION_RANGE, row);
    rowNumber = getRowNumberFromRange(updatedRange);
  }

  if (rowNumber) {
    withSessionRowNumber(session, rowNumber);
  }
  baccaratSessions.set(String(session.playerId), session);
}

async function deleteBaccaratSession(playerId) {
  const cachedSession = baccaratSessions.get(String(playerId));
  const rowNumber = cachedSession?._rowNumber || await getBaccaratSessionRowNumber(playerId);

  if (rowNumber) {
    await clearSingleRow(sheetRowRange(BACCARAT_SESSION_SHEET_REF, 'N', rowNumber));
  }

  baccaratSessions.delete(String(playerId));
}

async function ensureRussianRouletteSessionSheetHeader() {
  if (russianRouletteSessionSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === RUSSIAN_ROULETTE_SESSION_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: RUSSIAN_ROULETTE_SESSION_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RUSSIAN_ROULETTE_SESSION_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = RUSSIAN_ROULETTE_SESSION_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: RUSSIAN_ROULETTE_SESSION_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [RUSSIAN_ROULETTE_SESSION_HEADER] }
    });
  }

  russianRouletteSessionSheetReady = true;
}

function parseRussianRouletteSessionRow(row) {
  const winnerId = Number(row[9]);

  return {
    id: row[0] || '',
    participantIds: parseIdArray(row[1]),
    activeIds: parseIdArray(row[2]),
    eliminated: parseEliminatedPlayers(row[3]),
    color: row[4] || '',
    bet: toChipAmount(row[5]),
    pot: toChipAmount(row[6]),
    round: toChipAmount(row[7]),
    done: row[8] === 'true',
    winnerId: Number.isFinite(winnerId) ? winnerId : null,
    result: row[10] || '',
    lastAction: row[11] || ''
  };
}

function russianRouletteSessionToRow(session) {
  return [
    session.id,
    JSON.stringify(session.participantIds || []),
    JSON.stringify(session.activeIds || []),
    JSON.stringify(session.eliminated || []),
    session.color,
    session.bet,
    session.pot,
    session.round,
    session.done ? 'true' : 'false',
    session.winnerId || '',
    session.result || '',
    session.lastAction || '',
    new Date().toISOString()
  ];
}

async function getRussianRouletteSessionRowNumber(sessionId) {
  await ensureRussianRouletteSessionSheetHeader();
  return findRowNumberByFirstColumnText(RUSSIAN_ROULETTE_SESSION_ID_RANGE, sessionId);
}

async function getRussianRouletteSession(sessionId) {
  const rowNumber = await getRussianRouletteSessionRowNumber(sessionId);

  if (!rowNumber) {
    russianRouletteSessions.delete(String(sessionId));
    return null;
  }

  const row = await readSingleRow(sheetRowRange(RUSSIAN_ROULETTE_SESSION_SHEET_REF, 'M', rowNumber));
  const session = parseRussianRouletteSessionRow(row);

  if (!session.id) {
    russianRouletteSessions.delete(String(sessionId));
    return null;
  }

  withSessionRowNumber(session, rowNumber);
  russianRouletteSessions.set(String(session.id), session);
  return session;
}

async function saveRussianRouletteSession(session) {
  let rowNumber = session._rowNumber || await getRussianRouletteSessionRowNumber(session.id);
  const row = russianRouletteSessionToRow(session);

  if (rowNumber) {
    await updateSingleRow(sheetRowRange(RUSSIAN_ROULETTE_SESSION_SHEET_REF, 'M', rowNumber), row);
  } else {
    const updatedRange = await appendSingleRow(RUSSIAN_ROULETTE_SESSION_RANGE, row);
    rowNumber = getRowNumberFromRange(updatedRange);
  }

  if (rowNumber) {
    withSessionRowNumber(session, rowNumber);
  }
  russianRouletteSessions.set(String(session.id), session);
}

async function deleteRussianRouletteSession(sessionId) {
  const cachedSession = russianRouletteSessions.get(String(sessionId));
  const rowNumber = cachedSession?._rowNumber || await getRussianRouletteSessionRowNumber(sessionId);

  if (rowNumber) {
    await clearSingleRow(sheetRowRange(RUSSIAN_ROULETTE_SESSION_SHEET_REF, 'M', rowNumber));
  }

  russianRouletteSessions.delete(String(sessionId));
}

async function ensurePvpSessionSheetHeader() {
  if (pvpSessionSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === PVP_SESSION_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: PVP_SESSION_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: PVP_SESSION_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = PVP_SESSION_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: PVP_SESSION_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [PVP_SESSION_HEADER] }
    });
  }

  pvpSessionSheetReady = true;
}

function parsePvpSessionRow(row) {
  return {
    id: row[0] || '',
    gameType: row[1] || '',
    playerOneId: Number(row[2]),
    playerTwoId: Number(row[3]),
    color: row[4] || '',
    bet: toChipAmount(row[5]),
    deck: parseJsonArray(row[6]),
    shoe: parseJsonArray(row[7]),
    playerOneCards: parseJsonArray(row[8]),
    playerTwoCards: parseJsonArray(row[9]),
    playerOneAction: row[10] || '',
    playerTwoAction: row[11] || '',
    turn: row[12] || 'playerOne',
    done: row[13] === 'true',
    winnerKey: row[14] || '',
    result: row[15] || ''
  };
}

function pvpSessionToRow(session) {
  return [
    session.id,
    session.gameType,
    session.playerOneId,
    session.playerTwoId,
    session.color,
    session.bet,
    JSON.stringify(session.deck || []),
    JSON.stringify(session.shoe || []),
    JSON.stringify(session.playerOneCards || []),
    JSON.stringify(session.playerTwoCards || []),
    session.playerOneAction || '',
    session.playerTwoAction || '',
    session.turn || '',
    session.done ? 'true' : 'false',
    session.winnerKey || '',
    session.result || '',
    new Date().toISOString()
  ];
}

async function getPvpSessionRowNumber(sessionId) {
  await ensurePvpSessionSheetHeader();
  return findRowNumberByFirstColumnText(PVP_SESSION_ID_RANGE, sessionId);
}

async function getPvpSession(sessionId) {
  const rowNumber = await getPvpSessionRowNumber(sessionId);

  if (!rowNumber) {
    pvpSessions.delete(String(sessionId));
    return null;
  }

  const row = await readSingleRow(sheetRowRange(PVP_SESSION_SHEET_REF, 'Q', rowNumber));
  const session = parsePvpSessionRow(row);

  if (!session.id) {
    pvpSessions.delete(String(sessionId));
    return null;
  }

  withSessionRowNumber(session, rowNumber);
  pvpSessions.set(String(session.id), session);
  return session;
}

async function savePvpSession(session) {
  let rowNumber = session._rowNumber || await getPvpSessionRowNumber(session.id);
  const row = pvpSessionToRow(session);

  if (rowNumber) {
    await updateSingleRow(sheetRowRange(PVP_SESSION_SHEET_REF, 'Q', rowNumber), row);
  } else {
    const updatedRange = await appendSingleRow(PVP_SESSION_RANGE, row);
    rowNumber = getRowNumberFromRange(updatedRange);
  }

  if (rowNumber) {
    withSessionRowNumber(session, rowNumber);
  }
  pvpSessions.set(String(session.id), session);
}

async function deletePvpSession(sessionId) {
  const cachedSession = pvpSessions.get(String(sessionId));
  const rowNumber = cachedSession?._rowNumber || await getPvpSessionRowNumber(sessionId);

  if (rowNumber) {
    await clearSingleRow(sheetRowRange(PVP_SESSION_SHEET_REF, 'Q', rowNumber));
  }

  pvpSessions.delete(String(sessionId));
}

async function ensureFormulaHighLowSessionSheetHeader() {
  if (formulaHighLowSessionSheetReady) {
    return;
  }

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties.title'
  });
  const sheetExists = (spreadsheet.data.sheets || [])
    .some((sheet) => sheet.properties?.title === FORMULA_HIGH_LOW_SESSION_SHEET_NAME);

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: FORMULA_HIGH_LOW_SESSION_SHEET_NAME }
            }
          }
        ]
      }
    });
  }

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: FORMULA_HIGH_LOW_SESSION_HEADER_RANGE
  });
  const values = res.data.values || [];
  const header = values[0] || [];
  const hasHeader = FORMULA_HIGH_LOW_SESSION_HEADER.every((name, index) => header[index] === name);

  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: FORMULA_HIGH_LOW_SESSION_HEADER_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: [FORMULA_HIGH_LOW_SESSION_HEADER] }
    });
  }

  formulaHighLowSessionSheetReady = true;
}

function parseFormulaHighLowSessionRow(row) {
  return {
    id: row[0] || '',
    participantIds: parseIdArray(row[1]),
    color: row[2] || '',
    bet: toChipAmount(row[3]),
    target: Number(row[4]),
    operatorCards: parseJsonArray(row[5]),
    participantStates: parseJsonArray(row[6]),
    stage: row[7] || 'input',
    winnerIds: parseIdArray(row[8]),
    result: row[9] || '',
    memo: row[10] || '',
    pot: toChipAmount(row[11]),
    roundNumber: readGameLimitValue(row[12])
  };
}

function formulaHighLowSessionToRow(session) {
  return [
    session.id,
    JSON.stringify(session.participantIds || []),
    session.color,
    session.bet,
    Number.isFinite(Number(session.target)) ? session.target : '',
    JSON.stringify(session.operatorCards || []),
    JSON.stringify(session.participantStates || []),
    session.stage || '',
    JSON.stringify(session.winnerIds || []),
    session.result || '',
    session.memo || '',
    session.pot || 0,
    Number.isFinite(Number(session.roundNumber)) ? session.roundNumber : '',
    new Date().toISOString()
  ];
}

async function getFormulaHighLowSessionRowNumber(sessionId) {
  await ensureFormulaHighLowSessionSheetHeader();
  return findRowNumberByFirstColumnText(FORMULA_HIGH_LOW_SESSION_ID_RANGE, sessionId);
}

async function getFormulaHighLowSession(sessionId) {
  const rowNumber = await getFormulaHighLowSessionRowNumber(sessionId);

  if (!rowNumber) {
    formulaHighLowSessions.delete(String(sessionId));
    return null;
  }

  const row = await readSingleRow(sheetRowRange(FORMULA_HIGH_LOW_SESSION_SHEET_REF, 'N', rowNumber));
  const session = parseFormulaHighLowSessionRow(row);

  if (!session.id) {
    formulaHighLowSessions.delete(String(sessionId));
    return null;
  }

  withSessionRowNumber(session, rowNumber);
  formulaHighLowSessions.set(String(session.id), session);
  return session;
}

async function saveFormulaHighLowSession(session) {
  let rowNumber = session._rowNumber || await getFormulaHighLowSessionRowNumber(session.id);
  const row = formulaHighLowSessionToRow(session);

  if (rowNumber) {
    await updateSingleRow(sheetRowRange(FORMULA_HIGH_LOW_SESSION_SHEET_REF, 'N', rowNumber), row);
  } else {
    const updatedRange = await appendSingleRow(FORMULA_HIGH_LOW_SESSION_RANGE, row);
    rowNumber = getRowNumberFromRange(updatedRange);
  }

  if (rowNumber) {
    withSessionRowNumber(session, rowNumber);
  }
  formulaHighLowSessions.set(String(session.id), session);
}

async function deleteFormulaHighLowSession(sessionId) {
  const cachedSession = formulaHighLowSessions.get(String(sessionId));
  const rowNumber = cachedSession?._rowNumber || await getFormulaHighLowSessionRowNumber(sessionId);

  if (rowNumber) {
    await clearSingleRow(sheetRowRange(FORMULA_HIGH_LOW_SESSION_SHEET_REF, 'N', rowNumber));
  }

  formulaHighLowSessions.delete(String(sessionId));
}

function parseLogRow(row) {
  return {
    id: Number(row[0]) || Date.parse(row[4]) || 0,
    type: row[1] || '',
    text: row[2] || '',
    publicText: row[3] || row[2] || '',
    createdAt: row[4] || new Date(Number(row[0]) || Date.now()).toISOString(),
    groupId: row[5] || ''
  };
}

function normalizeLogLimit(value) {
  if (value === undefined || value === null || value === '') {
    return null;
  }

  const limit = Math.floor(Number(value));
  return Number.isFinite(limit) && limit > 0
    ? Math.min(limit, MAX_LOG_LIMIT)
    : null;
}

async function loadLimitedLogRows(limit) {
  const idRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: LOG_ID_RANGE
  });
  const rowCount = (idRes.data.values || []).length;

  if (rowCount <= 1) {
    return [];
  }

  const endRow = rowCount;
  const startRow = Math.max(2, endRow - limit + 1);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${LOG_SHEET_REF}!A${startRow}:F${endRow}`
  });

  return res.data.values || [];
}

async function loadLogHistory(limitValue) {
  await ensureLogSheetHeader();

  const limit = normalizeLogLimit(limitValue);
  let rows;

  if (limit) {
    rows = await loadLimitedLogRows(limit);
  } else {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: LOG_RANGE
    });
    rows = (res.data.values || []).slice(1);
  }

  logHistory = rows
    .filter((row) => row.length)
    .map(parseLogRow)
    .reverse();

  lastLogText = logHistory[0]?.text || '';
  return logHistory;
}

async function appendLogToSheet(log) {
  await ensureLogSheetHeader();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: LOG_APPEND_RANGE,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[log.id, log.type, log.text, log.publicText, log.createdAt, log.groupId || '']]
    }
  });
}

function getErrorDetail(err) {
  return err.response?.data?.error?.message || err.message || String(err);
}

function getEnvDebugInfo() {
  const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '';
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL || '';
  const privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  const normalizedPrivateKey = normalizePrivateKey(privateKey);

  return {
    vercel: Boolean(process.env.VERCEL),
    hasServiceAccountJson: Boolean(serviceAccountJson.trim()),
    serviceAccountJsonLength: serviceAccountJson.length,
    serviceAccountJsonLooksJson: serviceAccountJson.trim().startsWith('{'),
    hasClientEmail: Boolean(clientEmail.trim()),
    clientEmailLength: clientEmail.length,
    clientEmailLooksLikeServiceAccount: clientEmail.includes('iam.gserviceaccount.com'),
    hasPrivateKey: Boolean(privateKey.trim()),
    privateKeyLength: privateKey.length,
    privateKeyHasBegin: normalizedPrivateKey.includes('-----BEGIN PRIVATE KEY-----'),
    privateKeyHasEnd: normalizedPrivateKey.includes('-----END PRIVATE KEY-----'),
    privateKeyHasRealNewlines: normalizedPrivateKey.includes('\n'),
    usingCredentialMode: serviceAccountJson.trim()
      ? 'GOOGLE_SERVICE_ACCOUNT_JSON'
      : clientEmail.trim() && privateKey.trim()
        ? 'GOOGLE_CLIENT_EMAIL/GOOGLE_PRIVATE_KEY'
        : 'none'
  };
}

// =========================
// 🧰 공통 함수
// =========================
function chipStr(p) {
  return COLORS.map((color) => `${CHIP_LABELS[color]}:${p[color] || 0}`).join(' ');
}

function validateColor(color) {
  return COLORS.includes(color);
}

function toChipAmount(value) {
  return Math.floor(Number(value));
}

function readChipValue(value) {
  return Number.isFinite(Number(value)) ? toChipAmount(value) : 0;
}

function readGameLimitValue(value) {
  const text = String(value ?? '').trim();

  if (!text) {
    return null;
  }

  const limit = toChipAmount(text);
  return Number.isFinite(limit) ? Math.max(0, limit) : null;
}

function formatGameLimitValue(value) {
  return Number.isFinite(Number(value)) ? Math.max(0, toChipAmount(value)) : '';
}

function getKoreaDateKey(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: GAME_LIMIT_RESET_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(date);
  const byType = Object.fromEntries(parts.map((part) => [part.type, part.value]));

  return `${byType.year}-${byType.month}-${byType.day}`;
}

function getPlayerById(playerId) {
  return players.find((p) => p.id === Number(playerId));
}

function getPlayerTeam(player) {
  const team = String(player?.team || '').trim();
  return team || '팀 없음';
}

function getPlayerChipValue(player, rateMap = rates) {
  return COLORS.reduce((total, color) => total + ((player?.[color] || 0) * (rateMap[color] || 0)), 0);
}

function getLimitGameType(gameType) {
  return String(gameType || '').toLowerCase();
}

function getGameLimitConfig(gameType) {
  return GAME_LIMIT_BY_TYPE.get(getLimitGameType(gameType)) || null;
}

function getGameLimitValue(player, gameType) {
  const config = getGameLimitConfig(gameType);
  return config ? player?.[config.key] : null;
}

function createHttpError(status, message) {
  const err = new Error(message);
  err.status = status;
  return err;
}

function assertGameLimitAvailable(player, gameType) {
  const config = getGameLimitConfig(gameType);

  if (!config) {
    return;
  }

  const limit = getGameLimitValue(player, gameType);

  if (Number.isFinite(limit) && limit <= 0) {
    throw createHttpError(400, `${player.name}의 ${config.label} 플레이 가능 횟수가 없습니다.`);
  }
}

function consumeGameLimit(player, gameType) {
  const config = getGameLimitConfig(gameType);

  if (!config) {
    return;
  }

  const limit = getGameLimitValue(player, gameType);

  if (Number.isFinite(limit)) {
    player[config.key] = Math.max(0, limit - 1);
  }
}

function getGameLimits(player) {
  return Object.fromEntries(
    GAME_LIMITS.map((game) => [game.type, player?.[game.key] ?? null])
  );
}

function buildTeamTotals(playerList = players, rateMap = rates) {
  const teams = new Map();

  playerList.filter(Boolean).forEach((player) => {
    const teamName = getPlayerTeam(player);

    if (!teams.has(teamName)) {
      teams.set(teamName, {
        team: teamName,
        playerCount: 0,
        chipTotals: Object.fromEntries(COLORS.map((color) => [color, 0])),
        totalValue: 0
      });
    }

    const team = teams.get(teamName);
    team.playerCount += 1;

    COLORS.forEach((color) => {
      const amount = player[color] || 0;
      team.chipTotals[color] += amount;
      team.totalValue += amount * (rateMap[color] || 0);
    });
  });

  return [...teams.values()].sort((a, b) => b.totalValue - a.totalValue || a.team.localeCompare(b.team));
}

function serializePlayers(playerList = players) {
  return playerList
    .filter(Boolean)
    .map((player) => ({
      id: player.id,
      name: player.name,
      team: player.team || '',
      ...Object.fromEntries(COLORS.map((color) => [color, player[color] || 0])),
      totalValue: getPlayerChipValue(player),
      gameLimits: getGameLimits(player),
      ...Object.fromEntries(GAME_LIMITS.map((game) => [game.key, player[game.key] ?? null]))
    }));
}

function createLogGroupId(type, key = '') {
  const suffix = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  return [type, key, suffix].filter(Boolean).join(':');
}

function ensureLogGroupId(session, type, key = '') {
  if (!session.logGroupId) {
    session.logGroupId = createLogGroupId(type, key || session.id || session.playerId);
  }

  return session.logGroupId;
}

async function addLog(type, text, publicText = text, options = {}) {
  const log = {
    id: Date.now(),
    type,
    text,
    publicText,
    createdAt: new Date().toISOString(),
    groupId: options.groupId || ''
  };

  await appendLogToSheet(log);

  lastLogText = text;
  logHistory = [log, ...logHistory].slice(0, DEFAULT_LOG_LIMIT);
  return log;
}

function createBlackjackDeck() {
  const suits = ['S', 'H', 'D', 'C'];
  const ranks = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
  return suits.flatMap((suit) => ranks.map((rank) => `${suit}-${rank}`));
}

function drawBlackjackCard(deck) {
  const index = Math.floor(Math.random() * deck.length);
  const [card] = deck.splice(index, 1);
  return card;
}

function getBlackjackCardRank(card) {
  return String(card).split('-').pop();
}

function getBlackjackHandValue(cards) {
  let value = 0;
  let aces = 0;

  cards.forEach((card) => {
    const rank = getBlackjackCardRank(card);

    if (rank === 'A') {
      aces += 1;
      value += 11;
    } else if (['J', 'Q', 'K'].includes(rank)) {
      value += 10;
    } else {
      value += Number(rank);
    }
  });

  while (value > 21 && aces > 0) {
    value -= 10;
    aces -= 1;
  }

  return value;
}

function getBlackjackState(session) {
  return {
    playerId: session.playerId,
    color: session.color,
    bet: session.bet,
    openCard: session.dealerCards[0] || '',
    playerCards: session.playerCards,
    dealerCards: session.dealerCards,
    playerValue: getBlackjackHandValue(session.playerCards),
    dealerValue: getBlackjackHandValue(session.dealerCards),
    lastDraw: session.lastDraw || '',
    dealerDraws: session.dealerDraws || [],
    remainingCards: session.deck ? session.deck.length : 0,
    done: session.done || false,
    result: session.result || '',
    resultChipDelta: Number.isFinite(session.resultChipDelta) ? session.resultChipDelta : null,
    resultChipBalance: Number.isFinite(session.resultChipBalance) ? session.resultChipBalance : null,
    resultChipText: session.resultChipText || ''
  };
}

function blackjackStateText(session) {
  const state = getBlackjackState(session);
  return `플레이어 패: ${state.playerCards.join(', ')} (${state.playerValue})
딜러 패: ${state.dealerCards.join(', ')} (${state.dealerValue})
이번 결과: ${state.lastDraw || '-'}
딜러 추가 카드: ${state.dealerDraws.length ? state.dealerDraws.join(', ') : '-'}
남은 카드: ${state.remainingCards}`;
}

function getSignedChipAmount(amount) {
  const chipAmount = toChipAmount(amount);
  return chipAmount > 0 ? `+${chipAmount}` : String(chipAmount);
}

function getChipLabel(color) {
  return CHIP_LABELS[color] || color;
}

function makeChipResultText(color, delta, balance) {
  const label = getChipLabel(color);
  return `${label} ${getSignedChipAmount(delta)} / 현재 잔고 ${label} ${toChipAmount(balance)}`;
}

function makePlayerChipResultText(player, color, delta) {
  return makeChipResultText(color, delta, player?.[color] || 0);
}

function makePlayerChipResultFromBefore(player, color, beforeBalance) {
  const balance = player?.[color] || 0;
  return makeChipResultText(color, balance - beforeBalance, balance);
}

function makeNamedChipResultText(player, color, delta) {
  return `${player.name}: ${makePlayerChipResultText(player, color, delta)}`;
}

function chipResultLine(resultChipText) {
  return resultChipText ? `\n결과 칩: ${resultChipText}` : '';
}

function setSessionChipResult(player, session, delta) {
  session.resultChipDelta = toChipAmount(delta);
  session.resultChipBalance = player[session.color] || 0;
  session.resultChipText = makeChipResultText(session.color, session.resultChipDelta, session.resultChipBalance);
}

function setBlackjackChipResult(player, session, delta) {
  setSessionChipResult(player, session, delta);
}

function blackjackChoiceText(session) {
  return session.done ? '' : '\n선택 가능: Hit 또는 Stand';
}

function blackjackChipResultText(session) {
  return session.done ? chipResultLine(session.resultChipText) : '';
}

function makeBlackjackProgressLog(player, session, action, resultText) {
  return `블랙잭 진행
플레이어: ${player.name}
칩: ${session.color} ${session.bet}
행동: ${action}
${blackjackStateText(session)}${blackjackChoiceText(session)}
결과: ${resultText}
현재: ${chipStr(player)}${blackjackChipResultText(session)}`;
}

function getPublicDealerCards(session) {
  if (session.done) {
    return session.dealerCards.join(', ');
  }

  return `${session.dealerCards[0] || '-'}, Hidden`;
}

function makeBlackjackPublicLog(player, session, action, resultText) {
  const state = getBlackjackState(session);

  if (!session.done) {
    return `플레이어 패: ${state.playerCards.join(', ')} (${state.playerValue})
딜러 패: ${getPublicDealerCards(session)}
선택 가능: Hit 또는 Stand`;
  }

  const dealerValue = session.done ? ` (${state.dealerValue})` : '';

  return `플레이어: ${player.name}
베팅: ${session.color} ${session.bet}
플레이어 패: ${state.playerCards.join(', ')} (${state.playerValue})
딜러 패: ${getPublicDealerCards(session)}${dealerValue}
결과: ${resultText}${blackjackChipResultText(session)}`;
}

function finishBlackjackSession(player, session) {
  const playerValue = getBlackjackHandValue(session.playerCards);

  session.dealerDraws = [];

  if (playerValue <= 21) {
    while (getBlackjackHandValue(session.dealerCards) < 17) {
      const card = drawBlackjackCard(session.deck);
      session.dealerCards.push(card);
      session.dealerDraws.push(card);
    }
  }

  const dealerValue = getBlackjackHandValue(session.dealerCards);

  if (playerValue > 21) {
    player[session.color] -= session.bet;
    session.result = '패배';
    setBlackjackChipResult(player, session, -session.bet);
  } else if (dealerValue > 21 || playerValue > dealerValue) {
    player[session.color] += session.bet;
    session.result = '승리';
    setBlackjackChipResult(player, session, session.bet);
  } else if (playerValue < dealerValue) {
    player[session.color] -= session.bet;
    session.result = '패배';
    setBlackjackChipResult(player, session, -session.bet);
  } else {
    session.result = '무승부';
    setBlackjackChipResult(player, session, 0);
  }

  session.done = true;
  blackjackSessions.delete(String(session.playerId));
  return session.result;
}

function getPvpBlackjackWinnerKey(playerOneValue, playerTwoValue) {
  const playerOneBust = playerOneValue > 21;
  const playerTwoBust = playerTwoValue > 21;

  if (playerOneBust && playerTwoBust) {
    return 'tie';
  }

  if (playerOneBust) {
    return 'playerTwo';
  }

  if (playerTwoBust) {
    return 'playerOne';
  }

  if (playerOneValue > playerTwoValue) {
    return 'playerOne';
  }

  if (playerTwoValue > playerOneValue) {
    return 'playerTwo';
  }

  return 'tie';
}

const ROULETTE_RED_NUMBERS = new Set([1, 3, 5, 7, 9, 12, 14, 16, 18, 19, 21, 23, 25, 27, 30, 32, 34, 36]);

function parseRouletteNumbers(value) {
  return String(value || '')
    .split(/[^0-9]+/)
    .filter(Boolean)
    .map(Number)
    .filter((number) => Number.isInteger(number) && number >= 0 && number <= 36);
}

function getRouletteDozenNumbers(value) {
  const normalized = String(value || '').toLowerCase();
  const dozen = normalized.includes('2') || normalized.includes('second')
    ? 2
    : normalized.includes('3') || normalized.includes('third')
      ? 3
      : 1;
  const start = (dozen - 1) * 12 + 1;
  return Array.from({ length: 12 }, (_, index) => start + index);
}

function getRouletteColumnNumbers(value) {
  const column = Math.min(Math.max(Number(String(value || '1').match(/[1-3]/)?.[0] || 1), 1), 3);
  return Array.from({ length: 12 }, (_, index) => column + index * 3);
}

function getRouletteStreetNumbers(value) {
  const numbers = parseRouletteNumbers(value);
  const first = numbers.length ? numbers[0] : 1;
  const start = Math.max(1, Math.min(34, first - ((first - 1) % 3)));
  return [start, start + 1, start + 2];
}

function getRouletteSixLineNumbers(value) {
  const numbers = parseRouletteNumbers(value);
  const first = numbers.length ? numbers[0] : 1;
  const start = Math.max(1, Math.min(31, first - ((first - 1) % 3)));
  return [start, start + 1, start + 2, start + 3, start + 4, start + 5];
}

function evaluateRouletteBet(type, detail, result) {
  const betType = String(type || '').toLowerCase();
  let coveredNumbers = [];
  let payout = 0;
  let requiredCount = 0;

  if (betType === 'number' || betType === 'straight') {
    coveredNumbers = parseRouletteNumbers(detail).slice(0, 1);
    payout = 35;
    requiredCount = 1;
  } else if (betType === 'split') {
    coveredNumbers = parseRouletteNumbers(detail).slice(0, 2);
    payout = 17;
    requiredCount = 2;
  } else if (betType === 'street') {
    coveredNumbers = getRouletteStreetNumbers(detail);
    payout = 11;
  } else if (betType === 'corner') {
    coveredNumbers = parseRouletteNumbers(detail).slice(0, 4);
    payout = 8;
    requiredCount = 4;
  } else if (betType === 'sixline') {
    coveredNumbers = getRouletteSixLineNumbers(detail);
    payout = 5;
  } else if (betType === 'dozen') {
    coveredNumbers = getRouletteDozenNumbers(detail);
    payout = 2;
  } else if (betType === 'column') {
    coveredNumbers = getRouletteColumnNumbers(detail);
    payout = 2;
  } else if (betType === 'red') {
    coveredNumbers = [...ROULETTE_RED_NUMBERS];
    payout = 1;
  } else if (betType === 'black') {
    coveredNumbers = Array.from({ length: 36 }, (_, index) => index + 1)
      .filter((number) => !ROULETTE_RED_NUMBERS.has(number));
    payout = 1;
  } else if (betType === 'even') {
    coveredNumbers = Array.from({ length: 18 }, (_, index) => (index + 1) * 2);
    payout = 1;
  } else if (betType === 'odd') {
    coveredNumbers = Array.from({ length: 18 }, (_, index) => index * 2 + 1);
    payout = 1;
  } else if (betType === 'low' || betType === '1-18') {
    coveredNumbers = Array.from({ length: 18 }, (_, index) => index + 1);
    payout = 1;
  } else if (betType === 'high' || betType === '19-36') {
    coveredNumbers = Array.from({ length: 18 }, (_, index) => index + 19);
    payout = 1;
  }

  return {
    win: coveredNumbers.includes(result),
    payout,
    coveredNumbers,
    valid: payout > 0 && (!requiredCount || coveredNumbers.length === requiredCount)
  };
}

function playRoulette(player, color, bet, extra) {
  const result = Math.floor(Math.random() * 37);
  const evaluation = evaluateRouletteBet(extra.pick, extra.detail, result);

  if (!evaluation.valid || evaluation.coveredNumbers.length === 0) {
    const err = new Error('룰렛 베팅 정보가 올바르지 않습니다.');
    err.status = 400;
    throw err;
  }

  if (evaluation.win) {
    player[color] += bet * evaluation.payout;
  } else {
    player[color] -= bet;
  }

  const nextExtra = {
    ...extra,
    result,
    detail: extra.detail || evaluation.coveredNumbers.join(', ')
  };

  return {
    win: evaluation.win,
    payout: evaluation.payout,
    result,
    extra: nextExtra
  };
}

function drawHighLowCard() {
  const cards = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
  const index = Math.floor(Math.random() * cards.length);
  return {
    card: cards[index],
    value: index + 1
  };
}

function playHighLow(player, color, bet, extra) {
  const choice = String(extra.choice || extra.pick || '').toUpperCase();

  if (!['HIGH', 'LOW'].includes(choice)) {
    const err = new Error('하이로우 선택은 HIGH 또는 LOW여야 합니다.');
    err.status = 400;
    throw err;
  }

  const result = drawHighLowCard();
  const outcome = result.value > 7 ? 'HIGH' : result.value < 7 ? 'LOW' : 'PUSH';
  const win = choice === outcome;

  if (outcome === 'PUSH') {
    return {
      win: false,
      push: true,
      extra: {
        ...extra,
        choice,
        card: `${result.card} (${result.value})`
      }
    };
  }

  if (win) {
    player[color] += bet;
  } else {
    player[color] -= bet;
  }

  return {
    win,
    push: false,
    extra: {
      ...extra,
      choice,
      card: `${result.card} (${result.value})`
    }
  };
}

function createBaccaratShoe(deckCount = 8) {
  const suits = ['S', 'H', 'D', 'C'];
  const ranks = ['A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
  const shoe = [];

  for (let deck = 1; deck <= deckCount; deck += 1) {
    suits.forEach((suit) => {
      ranks.forEach((rank) => {
        shoe.push(`${deck}${suit}-${rank}`);
      });
    });
  }

  return shoe;
}

function drawBaccaratCard(shoe) {
  const index = Math.floor(Math.random() * shoe.length);
  const [card] = shoe.splice(index, 1);
  return card;
}

function getBaccaratCardValue(card) {
  const rank = getBlackjackCardRank(card);

  if (rank === 'A') {
    return 1;
  }

  if (['10', 'J', 'Q', 'K'].includes(rank)) {
    return 0;
  }

  return Number(rank);
}

function getBaccaratTotal(cards) {
  return cards.reduce((sum, card) => sum + getBaccaratCardValue(card), 0) % 10;
}

function shouldBankerDraw(bankerTotal, playerThirdValue) {
  if (bankerTotal <= 2) return true;
  if (bankerTotal === 3) return playerThirdValue !== 8;
  if (bankerTotal === 4) return playerThirdValue >= 2 && playerThirdValue <= 7;
  if (bankerTotal === 5) return playerThirdValue >= 4 && playerThirdValue <= 7;
  if (bankerTotal === 6) return playerThirdValue === 6 || playerThirdValue === 7;
  return false;
}

function dealBaccaratRound(shoe = createBaccaratShoe()) {
  const playerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
  const bankerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
  let playerTotal = getBaccaratTotal(playerCards);
  let bankerTotal = getBaccaratTotal(bankerCards);
  let playerThirdCard = '';
  let bankerThirdCard = '';

  if (playerTotal < 8 && bankerTotal < 8) {
    if (playerTotal <= 5) {
      playerThirdCard = drawBaccaratCard(shoe);
      playerCards.push(playerThirdCard);
      playerTotal = getBaccaratTotal(playerCards);
    }

    if (playerThirdCard) {
      const playerThirdValue = getBaccaratCardValue(playerThirdCard);

      if (shouldBankerDraw(bankerTotal, playerThirdValue)) {
        bankerThirdCard = drawBaccaratCard(shoe);
        bankerCards.push(bankerThirdCard);
        bankerTotal = getBaccaratTotal(bankerCards);
      }
    } else if (bankerTotal <= 5) {
      bankerThirdCard = drawBaccaratCard(shoe);
      bankerCards.push(bankerThirdCard);
      bankerTotal = getBaccaratTotal(bankerCards);
    }
  }

  const outcome = playerTotal > bankerTotal
    ? 'PLAYER'
    : bankerTotal > playerTotal
      ? 'BANKER'
      : 'TIE';

  return {
    playerCards,
    bankerCards,
    playerThirdCard,
    bankerThirdCard,
    playerTotal,
    bankerTotal,
    outcome,
    remainingCards: shoe.length
  };
}

function playBaccarat(player, color, bet, extra) {
  const side = String(extra.side || extra.pick || '').toUpperCase();

  if (!['PLAYER', 'BANKER', 'TIE'].includes(side)) {
    const err = new Error('바카라 선택은 PLAYER, BANKER, TIE 중 하나여야 합니다.');
    err.status = 400;
    throw err;
  }

  const shoe = createBaccaratShoe();
  const playerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
  const bankerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
  let playerTotal = getBaccaratTotal(playerCards);
  let bankerTotal = getBaccaratTotal(bankerCards);
  let playerThirdCard = null;
  let bankerThirdCard = null;

  if (playerTotal < 8 && bankerTotal < 8) {
    if (playerTotal <= 5) {
      playerThirdCard = drawBaccaratCard(shoe);
      playerCards.push(playerThirdCard);
      playerTotal = getBaccaratTotal(playerCards);
    }

    if (playerThirdCard) {
      const playerThirdValue = getBaccaratCardValue(playerThirdCard);
      if (shouldBankerDraw(bankerTotal, playerThirdValue)) {
        bankerThirdCard = drawBaccaratCard(shoe);
        bankerCards.push(bankerThirdCard);
        bankerTotal = getBaccaratTotal(bankerCards);
      }
    } else if (bankerTotal <= 5) {
      bankerThirdCard = drawBaccaratCard(shoe);
      bankerCards.push(bankerThirdCard);
      bankerTotal = getBaccaratTotal(bankerCards);
    }
  }

  const outcome = playerTotal > bankerTotal
    ? 'PLAYER'
    : bankerTotal > playerTotal
      ? 'BANKER'
      : 'TIE';
  const push = outcome === 'TIE' && side !== 'TIE';
  const win = side === outcome;
  const payout = side === 'TIE' ? 8 : side === 'BANKER' ? 0.95 : 1;

  if (win) {
    player[color] += toChipAmount(bet * payout);
  } else if (!push) {
    player[color] -= bet;
  }

  return {
    win,
    push,
    payout,
    outcome,
    extra: {
      ...extra,
      side,
      result: `${outcome}
PLAYER: ${playerCards.join(', ')} (${playerTotal})
BANKER: ${bankerCards.join(', ')} (${bankerTotal})
Player third: ${playerThirdCard || '-'}
Banker third: ${bankerThirdCard || '-'}`,
      turn: 'standard'
    }
  };
}

function isValidBaccaratSide(side) {
  return ['PLAYER', 'BANKER', 'TIE'].includes(side);
}

function getBaccaratPayout(side) {
  return side === 'TIE' ? 8 : side === 'BANKER' ? 0.95 : 1;
}

function isBaccaratNatural(session) {
  return getBaccaratTotal(session.playerCards) >= 8 || getBaccaratTotal(session.bankerCards) >= 8;
}

function getBaccaratState(session) {
  const playerTotal = getBaccaratTotal(session.playerCards);
  const bankerTotal = getBaccaratTotal(session.bankerCards);
  const push = session.outcome === 'TIE' && session.side !== 'TIE';
  const win = session.side === session.outcome;

  return {
    playerId: session.playerId,
    side: session.side,
    color: session.color,
    bet: session.bet,
    playerCards: session.playerCards,
    bankerCards: session.bankerCards,
    playerTotal,
    bankerTotal,
    playerThirdCard: session.playerThirdCard || '',
    bankerThirdCard: session.bankerThirdCard || '',
    playerAction: session.playerAction || '',
    remainingCards: session.shoe ? session.shoe.length : 0,
    done: session.done || false,
    outcome: session.outcome || '',
    push,
    win,
    resultChipDelta: Number.isFinite(session.resultChipDelta) ? session.resultChipDelta : null,
    resultChipBalance: Number.isFinite(session.resultChipBalance) ? session.resultChipBalance : null,
    resultChipText: session.resultChipText || ''
  };
}

function settleBaccaratBet(player, session) {
  const push = session.outcome === 'TIE' && session.side !== 'TIE';
  const win = session.side === session.outcome;
  const beforeBalance = player[session.color] || 0;

  if (win) {
    player[session.color] += toChipAmount(session.bet * getBaccaratPayout(session.side));
  } else if (!push) {
    player[session.color] -= session.bet;
  }

  setSessionChipResult(player, session, (player[session.color] || 0) - beforeBalance);

  return { win, push };
}

function finishBaccaratSession(player, session) {
  if (!isBaccaratNatural(session)) {
    const bankerTotal = getBaccaratTotal(session.bankerCards);

    if (session.playerThirdCard) {
      const playerThirdValue = getBaccaratCardValue(session.playerThirdCard);

      if (shouldBankerDraw(bankerTotal, playerThirdValue)) {
        session.bankerThirdCard = drawBaccaratCard(session.shoe);
        session.bankerCards.push(session.bankerThirdCard);
      }
    } else if (bankerTotal <= 5) {
      session.bankerThirdCard = drawBaccaratCard(session.shoe);
      session.bankerCards.push(session.bankerThirdCard);
    }
  }

  const playerTotal = getBaccaratTotal(session.playerCards);
  const bankerTotal = getBaccaratTotal(session.bankerCards);

  session.outcome = playerTotal > bankerTotal
    ? 'PLAYER'
    : bankerTotal > playerTotal
      ? 'BANKER'
      : 'TIE';
  session.done = true;

  const { win, push } = settleBaccaratBet(player, session);
  return push ? '무승부(push)' : win ? '승리' : '패배';
}

function baccaratStateText(session) {
  const state = getBaccaratState(session);

  return `PLAYER: ${state.playerCards.join(', ')} (${state.playerTotal})
BANKER: ${state.bankerCards.join(', ')} (${state.bankerTotal})
플레이어 선택: ${state.playerAction || '-'}
플레이어 추가 카드: ${state.playerThirdCard || '-'}
뱅커 추가 카드: ${state.bankerThirdCard || '-'}
남은 카드: ${state.remainingCards}`;
}

function makeBaccaratProgressLog(player, session, action, resultText) {
  return `바카라 진행
플레이어: ${player.name}
베팅: ${session.side} / ${session.color} ${session.bet}
행동: ${action}
${baccaratStateText(session)}
결과: ${session.outcome || resultText}
판정: ${resultText}
현재: ${chipStr(player)}${session.done ? chipResultLine(session.resultChipText) : ''}`;
}

function makeBaccaratPublicLog(player, session, action, resultText) {
  const state = getBaccaratState(session);

  if (!session.done) {
    return `베팅: ${session.side} / ${session.color} ${session.bet}
PLAYER: ${state.playerCards.join(', ')} (${state.playerTotal})
BANKER: ${state.bankerCards.join(', ')} (${state.bankerTotal})
선택 가능: Hit 또는 Stand`;
  }

  return `플레이어: ${player.name}
베팅: ${session.side} / ${session.color} ${session.bet}
행동: ${action}
PLAYER: ${state.playerCards.join(', ')} (${state.playerTotal})
BANKER: ${state.bankerCards.join(', ')} (${state.bankerTotal})
결과: ${state.outcome}
판정: ${resultText}${chipResultLine(session.resultChipText)}`;
}

function settlePvpBet(playerOne, playerTwo, color, bet, winnerKey) {
  if (winnerKey === 'playerOne') {
    playerOne[color] += bet;
    playerTwo[color] -= bet;
  } else if (winnerKey === 'playerTwo') {
    playerOne[color] -= bet;
    playerTwo[color] += bet;
  }
}

function getPvpWinnerName(playerOne, playerTwo, winnerKey) {
  if (winnerKey === 'playerOne') {
    return playerOne.name;
  }

  if (winnerKey === 'playerTwo') {
    return playerTwo.name;
  }

  return '무승부';
}

function getPvpPlayerName(playerId) {
  return getPlayerById(playerId)?.name || `#${playerId}`;
}

function getPvpSessionState(session) {
  const playerOneName = getPvpPlayerName(session.playerOneId);
  const playerTwoName = getPvpPlayerName(session.playerTwoId);
  const activePlayerName = session.turn === 'playerTwo' ? playerTwoName : playerOneName;
  const state = {
    id: session.id,
    gameType: session.gameType,
    playerOne: { id: session.playerOneId, name: playerOneName },
    playerTwo: { id: session.playerTwoId, name: playerTwoName },
    color: session.color,
    bet: session.bet,
    playerOneCards: session.playerOneCards || [],
    playerTwoCards: session.playerTwoCards || [],
    playerOneAction: session.playerOneAction || '',
    playerTwoAction: session.playerTwoAction || '',
    turn: session.turn || 'playerOne',
    activePlayerName,
    done: session.done || false,
    winnerKey: session.winnerKey || '',
    winnerName: session.winnerKey ? getPvpWinnerName({ name: playerOneName }, { name: playerTwoName }, session.winnerKey) : '',
    result: session.result || '',
    playerOneResultChipText: session.playerOneResultChipText || '',
    playerTwoResultChipText: session.playerTwoResultChipText || '',
    resultChipTexts: session.resultChipTexts || [],
    remainingCards: session.gameType === 'pvpbaccarat'
      ? (session.shoe || []).length
      : (session.deck || []).length
  };

  if (session.gameType === 'pvpbaccarat') {
    state.playerOneTotal = getBaccaratTotal(state.playerOneCards);
    state.playerTwoTotal = getBaccaratTotal(state.playerTwoCards);
    state.outcome = session.done
      ? state.playerOneTotal > state.playerTwoTotal
        ? 'PLAYER'
        : state.playerTwoTotal > state.playerOneTotal
          ? 'BANKER'
          : 'TIE'
      : '';
  } else {
    state.playerOneValue = getBlackjackHandValue(state.playerOneCards);
    state.playerTwoValue = getBlackjackHandValue(state.playerTwoCards);
  }

  return state;
}

function getPvpWinnerKey(session) {
  if (session.gameType === 'pvpbaccarat') {
    const playerOneTotal = getBaccaratTotal(session.playerOneCards);
    const playerTwoTotal = getBaccaratTotal(session.playerTwoCards);

    if (playerOneTotal > playerTwoTotal) {
      return 'playerOne';
    }

    if (playerTwoTotal > playerOneTotal) {
      return 'playerTwo';
    }

    return 'tie';
  }

  return getPvpBlackjackWinnerKey(
    getBlackjackHandValue(session.playerOneCards),
    getBlackjackHandValue(session.playerTwoCards)
  );
}

function finishPvpSession(playerOne, playerTwo, session) {
  const winnerKey = getPvpWinnerKey(session);
  const playerOneBefore = playerOne[session.color] || 0;
  const playerTwoBefore = playerTwo[session.color] || 0;
  settlePvpBet(playerOne, playerTwo, session.color, session.bet, winnerKey);
  session.done = true;
  session.winnerKey = winnerKey;
  session.result = winnerKey === 'tie'
    ? '무승부'
    : `${getPvpWinnerName(playerOne, playerTwo, winnerKey)} 승리`;
  session.playerOneResultChipText = makePlayerChipResultFromBefore(playerOne, session.color, playerOneBefore);
  session.playerTwoResultChipText = makePlayerChipResultFromBefore(playerTwo, session.color, playerTwoBefore);
  session.resultChipTexts = [
    `${playerOne.name}: ${session.playerOneResultChipText}`,
    `${playerTwo.name}: ${session.playerTwoResultChipText}`
  ];
  return session.result;
}

function isPvpPlayerDone(session, key) {
  return ['stand', 'bust', 'hit', 'natural'].includes(session[`${key}Action`]);
}

function advancePvpTurnOrFinish(playerOne, playerTwo, session) {
  if (session.gameType === 'pvpbaccarat') {
    if (!isPvpPlayerDone(session, 'playerOne')) {
      session.turn = 'playerOne';
      return '진행 중';
    }

    if (!isPvpPlayerDone(session, 'playerTwo')) {
      session.turn = 'playerTwo';
      return '진행 중';
    }

    return finishPvpSession(playerOne, playerTwo, session);
  }

  if (!isPvpPlayerDone(session, 'playerOne')) {
    session.turn = 'playerOne';
    return '진행 중';
  }

  if (!isPvpPlayerDone(session, 'playerTwo')) {
    session.turn = 'playerTwo';
    return '진행 중';
  }

  return finishPvpSession(playerOne, playerTwo, session);
}

function applyPvpAction(playerOne, playerTwo, session, action) {
  const nextAction = String(action || '').toLowerCase();

  if (!['hit', 'stand'].includes(nextAction)) {
    throw createHttpError(400, '행동은 hit 또는 stand여야 합니다.');
  }

  if (session.done) {
    throw createHttpError(400, '이미 종료된 PVP 게임입니다.');
  }

  const key = session.turn === 'playerTwo' ? 'playerTwo' : 'playerOne';
  const cards = key === 'playerTwo' ? session.playerTwoCards : session.playerOneCards;

  if (isPvpPlayerDone(session, key)) {
    throw createHttpError(400, '이미 행동을 마친 플레이어입니다.');
  }

  if (nextAction === 'hit') {
    if (session.gameType === 'pvpbaccarat') {
      cards.push(drawBaccaratCard(session.shoe));
      session[`${key}Action`] = 'hit';
    } else {
      cards.push(drawBlackjackCard(session.deck));
      session[`${key}Action`] = getBlackjackHandValue(cards) > 21 ? 'bust' : '';
    }
  } else {
    session[`${key}Action`] = 'stand';
  }

  return advancePvpTurnOrFinish(playerOne, playerTwo, session);
}

function makePvpLog(playerOne, playerTwo, session, resultText) {
  const state = getPvpSessionState(session);
  const title = session.gameType === 'pvpbaccarat' ? 'PVP 바카라' : 'PVP 블랙잭';
  const valueLabel = session.gameType === 'pvpbaccarat' ? 'Total' : 'Value';
  const playerOneValue = session.gameType === 'pvpbaccarat' ? state.playerOneTotal : state.playerOneValue;
  const playerTwoValue = session.gameType === 'pvpbaccarat' ? state.playerTwoTotal : state.playerTwoValue;

  return `${title}
플레이어 1: ${playerOne.name}
플레이어 2: ${playerTwo.name}
베팅: ${session.color} ${session.bet}
턴: ${state.done ? '-' : state.activePlayerName}
${playerOne.name} 행동: ${state.playerOneAction || '-'}
${playerTwo.name} 행동: ${state.playerTwoAction || '-'}
${playerOne.name} 패: ${state.playerOneCards.join(', ')} (${valueLabel}: ${playerOneValue})
${playerTwo.name} 패: ${state.playerTwoCards.join(', ')} (${valueLabel}: ${playerTwoValue})
결과: ${resultText}
승자: ${state.winnerName || '-'}
${state.done && state.resultChipTexts.length ? `결과 칩:\n${state.resultChipTexts.join('\n')}` : ''}
${playerOne.name}: ${chipStr(playerOne)}
${playerTwo.name}: ${chipStr(playerTwo)}`;
}

function formatPvpCardsForViewer(cards, valueLabel, value, revealed) {
  if (!revealed) {
    return `비공개 (${valueLabel}: -)`;
  }

  return `${cards.join(', ')} (${valueLabel}: ${value})`;
}

function formatPvpOpponentCardsForViewer(cards, valueLabel, value, revealed, gameType) {
  if (revealed) {
    return formatPvpCardsForViewer(cards, valueLabel, value, true);
  }

  if (gameType === 'pvpblackjack') {
    const visibleCards = cards.length
      ? [cards[0], ...cards.slice(1).map(() => 'HIDDEN')]
      : ['HIDDEN'];

    return `${visibleCards.join(', ')} (${valueLabel}: -)`;
  }

  return formatPvpCardsForViewer(cards, valueLabel, value, false);
}

function makePvpPlayerPublicSection(viewer, state, session, viewerKey) {
  const valueLabel = session.gameType === 'pvpbaccarat' ? 'Total' : 'Value';
  const opponentKey = viewerKey === 'playerOne' ? 'playerTwo' : 'playerOne';
  const viewerCards = state[`${viewerKey}Cards`] || [];
  const opponentCards = state[`${opponentKey}Cards`] || [];
  const viewerValue = session.gameType === 'pvpbaccarat'
    ? state[`${viewerKey}Total`]
    : state[`${viewerKey}Value`];
  const opponentValue = session.gameType === 'pvpbaccarat'
    ? state[`${opponentKey}Total`]
    : state[`${opponentKey}Value`];
  const revealOpponent = Boolean(state.done);

  return `${viewer.name}
내 패: ${formatPvpCardsForViewer(viewerCards, valueLabel, viewerValue, true)}
상대 패: ${formatPvpOpponentCardsForViewer(opponentCards, valueLabel, opponentValue, revealOpponent, session.gameType)}${state.done ? chipResultLine(state[`${viewerKey}ResultChipText`]) : ''}`;
}

function makePvpPublicLog(playerOne, playerTwo, session, resultText) {
  const state = getPvpSessionState(session);
  const title = session.gameType === 'pvpbaccarat' ? 'PVP 바카라' : 'PVP 블랙잭';

  return `${title}
[공통 공개]
플레이어 1: ${playerOne.name}
플레이어 2: ${playerTwo.name}
베팅: ${session.color} ${session.bet}
결과: ${resultText}
승자: ${state.winnerName || '-'}${state.done && state.resultChipTexts.length ? `\n결과 칩:\n${state.resultChipTexts.join('\n')}` : ''}

[플레이어 1 공개]
${makePvpPlayerPublicSection(playerOne, state, session, 'playerOne')}

[플레이어 2 공개]
${makePvpPlayerPublicSection(playerTwo, state, session, 'playerTwo')}`;
}

function playRedBlack(player, color, bet, extra) {
  const pick = String(extra.pick || '').toUpperCase();

  if (!['RED', 'BLACK'].includes(pick)) {
    const err = new Error('레드 앤 블랙 선택은 RED 또는 BLACK이어야 합니다.');
    err.status = 400;
    throw err;
  }

  const result = Math.random() < 0.5 ? 'RED' : 'BLACK';
  const win = pick === result;

  if (win) {
    player[color] += bet;
  } else {
    player[color] -= bet;
  }

  return {
    win,
    result,
    extra: {
      ...extra,
      pick,
      result
    }
  };
}

function getRussianRoulettePlayerName(playerId) {
  return getPlayerById(playerId)?.name || `#${playerId}`;
}

function getRussianRouletteState(session) {
  return {
    id: session.id,
    participants: session.participantIds.map((id) => ({
      id,
      name: getRussianRoulettePlayerName(id)
    })),
    activePlayers: session.activeIds.map((id) => ({
      id,
      name: getRussianRoulettePlayerName(id)
    })),
    eliminatedPlayers: session.eliminated.map((entry) => ({
      id: entry.id,
      name: getRussianRoulettePlayerName(entry.id),
      round: entry.round
    })),
    winnerId: session.winnerId || null,
    winnerName: session.winnerId ? getRussianRoulettePlayerName(session.winnerId) : '',
    color: session.color,
    bet: session.bet,
    pot: session.pot,
    round: session.round,
    done: session.done || false,
    result: session.result || '',
    lastAction: session.lastAction || '',
    resultChipTexts: session.resultChipTexts || []
  };
}

function makeRussianRouletteLog(session) {
  const state = getRussianRouletteState(session);
  const active = state.activePlayers.map((p) => p.name).join(', ') || '-';
  const eliminated = state.eliminatedPlayers.map((p) => `${p.round}R ${p.name}`).join(', ') || '-';

  return `러시안룰렛
참가자: ${state.participants.map((p) => p.name).join(', ')}
생존자: ${active}
탈락자: ${eliminated}
라운드: ${state.round}
베팅: ${state.color} ${state.bet}
팟: ${state.color} ${state.pot}
결과: ${state.lastAction}${state.done ? `\n최종 승자: ${state.winnerName}` : ''}${state.done && state.resultChipTexts.length ? `\n결과 칩:\n${state.resultChipTexts.join('\n')}` : ''}`;
}

function makeRussianRoulettePublicLog(session) {
  const state = getRussianRouletteState(session);
  const active = state.activePlayers.map((p) => p.name).join(', ') || '-';
  const eliminated = state.eliminatedPlayers.map((p) => `${p.round}R ${p.name}`).join(', ') || '-';
  const resultText = state.done
    ? `${state.color} ${state.pot} 지급`
    : state.lastAction;

  return `생존자: ${active}
탈락자: ${eliminated}
라운드: ${state.round}
결과: ${resultText}${state.done ? `\n최종 승자: ${state.winnerName}` : ''}${state.done && state.resultChipTexts.length ? `\n결과 칩:\n${state.resultChipTexts.join('\n')}` : ''}`;
}

// =========================
// 🎮 게임별 로그
// =========================
function logRoulette(p, color, bet, pick, detail, result, multiplier, win, resultChipText = '') {
  return `🎡 룰렛
${p.name} | 칩: ${color} ${bet}
베팅 종류: ${pick}
세부: ${detail || '-'}
결과: ${result}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
}

function logHighLow(p, color, bet, choice, card, multiplier, win, resultChipText = '') {
  return `📈 하이로우
${p.name} | 칩: ${color} ${bet}
선택: ${choice}
오픈 카드: ${card || '-'}
배율: ${multiplier}배
판정: ${multiplier === 'push' ? '무승부' : win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
}

function logBaccarat(p, color, bet, side, result, multiplier, turn, win, resultChipText = '') {
  return `🃏 바카라
${p.name} | 칩: ${color} ${bet}
턴: ${turn}
선택: ${side}
결과: ${result}
배율: ${multiplier}배
판정: ${multiplier === 'push' ? '무승부' : win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
}

function logBlackjack(p, color, bet, action, playerSum, dealerSum, multiplier, turn, win, resultChipText = '') {
  return `♠ 블랙잭
${p.name} | 칩: ${color} ${bet}
턴: ${turn}
행동: ${action}
플레이어: ${playerSum}
딜러: ${dealerSum}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
}

function logRedBlack(p, color, bet, pick, result, multiplier, win, resultChipText = '') {
  return `🟥⬛ 레드 앤 블랙
${p.name} | 칩: ${color} ${bet}
선택: ${pick}
결과: ${result}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
}

function getLogJudgement(multiplier, win) {
  return multiplier === 'push' ? '무승부' : win ? '승리' : '패배';
}

function makeGameLog(gameType, p, color, amount, extra, multiplier, win, resultChipText = '') {
  switch (gameType) {
    case 'roulette':
      return logRoulette(
        p,
        color,
        amount,
        extra.pick,
        extra.detail,
        extra.result,
        multiplier,
        win,
        resultChipText
      );
    case 'highlow':
      return logHighLow(
        p,
        color,
        amount,
        extra.choice,
        extra.card,
        multiplier,
        win,
        resultChipText
      );
    case 'baccarat':
      return logBaccarat(
        p,
        color,
        amount,
        extra.side,
        extra.result,
        multiplier,
        extra.turn,
        win,
        resultChipText
      );
    case 'blackjack':
      return logBlackjack(
        p,
        color,
        amount,
        extra.action,
        extra.playerSum,
        extra.dealerSum,
        multiplier,
        extra.turn,
        win,
        resultChipText
      );
    case 'redblack':
      return logRedBlack(
        p,
        color,
        amount,
        extra.pick,
        extra.result,
        multiplier,
        win,
        resultChipText
      );
    default:
      return `🎮 게임 결과
${p.name} | 칩: ${color} ${amount}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}${chipResultLine(resultChipText)}
현재: ${chipStr(p)}`;
  }
}

function makeGamePublicLog(gameType, p, color, amount, extra, multiplier, win, resultChipText = '') {
  const judgement = getLogJudgement(multiplier, win);
  const chipLine = chipResultLine(resultChipText);

  switch (gameType) {
    case 'roulette':
      return `플레이어: ${p.name}
베팅: ${color} ${amount}
베팅 종류: ${extra.pick}
결과: ${extra.result}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
    case 'highlow':
      return `플레이어: ${p.name}
베팅: ${color} ${amount}
선택: ${extra.choice}
오픈 카드: ${extra.card || '-'}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
    case 'baccarat':
      return `플레이어: ${p.name}
베팅: ${extra.side} / ${color} ${amount}
결과: ${extra.result}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
    case 'blackjack':
      return `플레이어: ${p.name}
베팅: ${color} ${amount}
턴: ${extra.turn || '-'}
행동: ${extra.action || '-'}
플레이어: ${extra.playerSum || '-'}
딜러: ${extra.dealerSum || '-'}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
    case 'redblack':
      return `플레이어: ${p.name}
베팅: ${color} ${amount}
선택: ${extra.pick}
결과: ${extra.result}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
    default:
      return `플레이어: ${p.name}
베팅: ${color} ${amount}
배율: ${multiplier}배
판정: ${judgement}${chipLine}`;
  }
}

function makeTransferLog(from, to, color, amount) {
  return `📦 칩 이동
보내는 사람: ${from.name}
받는 사람: ${to.name}
이동 칩: ${color} ${amount}

${from.name}: ${chipStr(from)}
${to.name}: ${chipStr(to)}`;
}

function makeTransferPublicLog(from, to, color, amount) {
  return `받는 사람: ${to.name}
보내는 사람: ${from.name}
이동 칩: ${color} ${amount}`;
}

function makeConvertLog(p, fromColor, toColor, amount, result) {
  return `💱 환전
플레이어: ${p.name}
변환: ${fromColor} ${amount} → ${toColor} ${result}
기준 비율: ${COLORS.map((color) => `${CHIP_LABELS[color]}=${rates[color]}`).join(', ')}
현재: ${chipStr(p)}`;
}

function makeConvertPublicLog(p, fromColor, toColor, amount, result) {
  return `플레이어: ${p.name}
변환: ${fromColor} ${amount} → ${toColor} ${result}
기준 비율: ${COLORS.map((color) => `${CHIP_LABELS[color]}=${rates[color]}`).join(', ')}`;
}

function makeChipAdjustLog(p, color, amount, action) {
  const label = action === 'add' ? '추가' : '감소';

  return `관리자 칩 ${label}
플레이어: ${p.name}
변경 칩: ${color} ${amount}
현재: ${chipStr(p)}`;
}

function makeChipAdjustPublicLog(p, color, amount, action) {
  const label = action === 'add' ? '추가' : '감소';

  return `플레이어: ${p.name}
변경: ${color} ${amount} ${label}`;
}

function makeBalanceLog(p) {
  return `현재 잔고
ID: ${p.id}
플레이어: ${p.name}
잔고: ${chipStr(p)}
총 가치: ${getPlayerChipValue(p)}`;
}

function makeBalancePublicLog(p) {
  return `플레이어: ${p.name}
잔고: ${chipStr(p)}
총 가치: ${getPlayerChipValue(p)}`;
}

const FORMULA_CARD_SUITS = [
  { key: 'gold', label: '금', rank: 4 },
  { key: 'silver', label: '은', rank: 3 },
  { key: 'bronze', label: '동', rank: 2 },
  { key: 'black', label: '흑', rank: 1 }
];
const FORMULA_CARD_SUIT_RANK = Object.fromEntries(FORMULA_CARD_SUITS.map((suit) => [suit.key, suit.rank]));
const FORMULA_CARD_SUIT_LABEL = Object.fromEntries(FORMULA_CARD_SUITS.map((suit) => [suit.key, suit.label]));
const FORMULA_OPERATOR_LABELS = {
  '+': '+',
  '-': '-',
  '*': '×',
  '/': '÷'
};
const FORMULA_CHOICES = ['HIGH', 'LOW', 'SWING'];
const FORMULA_EPSILON = 1e-9;

function createFormulaHighLowDeck() {
  const numberCards = FORMULA_CARD_SUITS.flatMap((suit) =>
    Array.from({ length: 11 }, (_, value) => ({
      kind: 'number',
      suit: suit.key,
      value,
      id: `${suit.key}-${value}`
    }))
  );
  const rootCards = Array.from({ length: 4 }, (_, index) => ({
    kind: 'root',
    id: `root-${index + 1}`
  }));
  const multiplyCards = Array.from({ length: 4 }, (_, index) => ({
    kind: 'multiply',
    id: `multiply-${index + 1}`
  }));

  return [...numberCards, ...rootCards, ...multiplyCards];
}

function drawFormulaCard(deck) {
  if (!deck.length) {
    throw createHttpError(400, '수식 하이 로우 덱에 남은 카드가 없습니다.');
  }

  const index = Math.floor(Math.random() * deck.length);
  const [card] = deck.splice(index, 1);
  return card;
}

function isFormulaNumberCard(card) {
  return card?.kind === 'number';
}

function drawFormulaNumberCard(deck, discardedCards) {
  while (deck.length) {
    const card = drawFormulaCard(deck);

    if (isFormulaNumberCard(card)) {
      return card;
    }

    discardedCards.push(card);
  }

  throw createHttpError(400, '추가 드로우에서 숫자 카드를 찾지 못했습니다.');
}

function drawFormulaHiddenCard(deck, discardedCards) {
  return drawFormulaNumberCard(deck, discardedCards);
}

function drawFormulaOpenSlot(deck, discardedCards) {
  const card = drawFormulaCard(deck);

  if (isFormulaNumberCard(card)) {
    return { symbol: null, number: card };
  }

  if (card.kind === 'root' || card.kind === 'multiply') {
    return {
      symbol: card,
      number: drawFormulaNumberCard(deck, discardedCards)
    };
  }

  discardedCards.push(card);
  return drawFormulaOpenSlot(deck, discardedCards);
}

function createFormulaPlayerState(deck, discardedCards) {
  return {
    hidden: drawFormulaHiddenCard(deck, discardedCards),
    openSlots: [
      drawFormulaOpenSlot(deck, discardedCards),
      drawFormulaOpenSlot(deck, discardedCards)
    ]
  };
}

function dealFormulaFinalOpenSlot(playerState, deck, discardedCards) {
  if ((playerState.openSlots || []).length >= 3) {
    throw createHttpError(400, '이미 최종 오픈 카드를 받았습니다.');
  }

  playerState.openSlots = [...(playerState.openSlots || []), drawFormulaOpenSlot(deck, discardedCards)];
}

function formatFormulaCard(card) {
  if (!card) {
    return '-';
  }

  if (card.kind === 'number') {
    return `${FORMULA_CARD_SUIT_LABEL[card.suit] || card.suit}${card.value}`;
  }

  if (card.kind === 'root') {
    return '√';
  }

  if (card.kind === 'multiply') {
    return '×';
  }

  return card.id || String(card);
}

function formatFormulaOpenSlot(slot) {
  if (!slot) {
    return '-';
  }

  if (slot.symbol) {
    return `${formatFormulaCard(slot.symbol)} ${formatFormulaCard(slot.number)}`;
  }

  return formatFormulaCard(slot.number);
}

function getFormulaNumbers(playerState) {
  return [
    playerState.hidden,
    ...(playerState.openSlots || []).map((slot) => slot.number)
  ].filter(isFormulaNumberCard);
}

function getFormulaRootCount(playerState) {
  return (playerState.openSlots || []).filter((slot) => slot.symbol?.kind === 'root').length;
}

function getFormulaMultiplyCount(playerState) {
  return (playerState.openSlots || []).filter((slot) => slot.symbol?.kind === 'multiply').length;
}

function uniqueFormulaPermutations(items) {
  const results = [];
  const used = Array(items.length).fill(false);
  const seen = new Set();

  function visit(path) {
    if (path.length === items.length) {
      const key = path.join('|');

      if (!seen.has(key)) {
        seen.add(key);
        results.push(path.map((index) => items[index]));
      }

      return;
    }

    for (let index = 0; index < items.length; index += 1) {
      if (!used[index]) {
        used[index] = true;
        path.push(index);
        visit(path);
        path.pop();
        used[index] = false;
      }
    }
  }

  visit([]);
  return results;
}

function getFormulaRootAssignments(numberCount, rootCount) {
  if (rootCount <= 0) {
    return [[]];
  }

  const assignments = [];

  function visit(startIndex, selected) {
    if (selected.length === rootCount) {
      assignments.push([...selected]);
      return;
    }

    for (let index = startIndex; index < numberCount; index += 1) {
      selected.push(index);
      visit(index + 1, selected);
      selected.pop();
    }
  }

  visit(0, []);
  return assignments;
}

function uniqueFormulaOperatorSets(sets) {
  const seen = new Set();

  return sets.filter((set) => {
    const key = [...set].sort().join('');

    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function getFormulaOperatorSets(multiplyCount) {
  let sets = [['+', '-', '/']];

  for (let count = 0; count < multiplyCount; count += 1) {
    const nextSets = [];

    sets.forEach((set) => {
      const expanded = [...set, '*'];

      expanded.forEach((operator, index) => {
        if (['+', '-', '*'].includes(operator)) {
          nextSets.push(expanded.filter((_, candidateIndex) => candidateIndex !== index));
        }
      });
    });

    sets = uniqueFormulaOperatorSets(nextSets);
  }

  return sets;
}

function applyFormulaOperator(left, operator, right) {
  if (operator === '+') {
    return left + right;
  }

  if (operator === '-') {
    return left - right;
  }

  if (operator === '*') {
    return left * right;
  }

  if (operator === '/') {
    return Math.abs(right) < FORMULA_EPSILON ? null : left / right;
  }

  return null;
}

function formatFormulaValue(value) {
  if (!Number.isFinite(value)) {
    return '-';
  }

  const rounded = Math.round(value * 10000) / 10000;
  return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
}

function makeFormulaOperand(card, useRoot) {
  const label = formatFormulaCard(card);

  return {
    value: useRoot ? Math.sqrt(card.value) : card.value,
    expression: useRoot ? `√(${label})` : label,
    card
  };
}

function makeFormulaNode(value, expression) {
  return Number.isFinite(value) ? { value, expression } : null;
}

function combineFormulaNodes(left, operator, right) {
  const value = applyFormulaOperator(left.value, operator, right.value);

  if (value === null || !Number.isFinite(value)) {
    return null;
  }

  return makeFormulaNode(value, `(${left.expression} ${FORMULA_OPERATOR_LABELS[operator]} ${right.expression})`);
}

function evaluateFormulaShapes(operands, operators) {
  const [a, b, c, d] = operands;
  const [op1, op2, op3] = operators;
  const results = [];
  const ab = combineFormulaNodes(a, op1, b);
  const bc = combineFormulaNodes(b, op2, c);
  const cd = combineFormulaNodes(c, op3, d);

  if (ab) {
    const abc = combineFormulaNodes(ab, op2, c);

    if (abc) {
      const abcd = combineFormulaNodes(abc, op3, d);

      if (abcd) {
        results.push(abcd);
      }
    }
  }

  if (bc) {
    const abc = combineFormulaNodes(a, op1, bc);

    if (abc) {
      const abcd = combineFormulaNodes(abc, op3, d);

      if (abcd) {
        results.push(abcd);
      }
    }
  }

  if (bc) {
    const bcd = combineFormulaNodes(bc, op3, d);

    if (bcd) {
      const abcd = combineFormulaNodes(a, op1, bcd);

      if (abcd) {
        results.push(abcd);
      }
    }
  }

  if (cd) {
    const bcd = combineFormulaNodes(b, op2, cd);

    if (bcd) {
      const abcd = combineFormulaNodes(a, op1, bcd);

      if (abcd) {
        results.push(abcd);
      }
    }
  }

  if (ab && cd) {
    const abcd = combineFormulaNodes(ab, op2, cd);

    if (abcd) {
      results.push(abcd);
    }
  }

  return results;
}

function getFormulaTieCard(playerState, side) {
  const cards = getFormulaNumbers(playerState);

  return cards.reduce((best, card) => {
    if (!best) {
      return card;
    }

    if (side === 'HIGH') {
      if (card.value !== best.value) {
        return card.value > best.value ? card : best;
      }

      return FORMULA_CARD_SUIT_RANK[card.suit] > FORMULA_CARD_SUIT_RANK[best.suit] ? card : best;
    }

    if (card.value !== best.value) {
      return card.value < best.value ? card : best;
    }

    return FORMULA_CARD_SUIT_RANK[card.suit] < FORMULA_CARD_SUIT_RANK[best.suit] ? card : best;
  }, null);
}

function findBestFormulaForTarget(playerState, target) {
  const numbers = getFormulaNumbers(playerState);
  const rootCount = getFormulaRootCount(playerState);
  const operatorSets = getFormulaOperatorSets(getFormulaMultiplyCount(playerState));
  let best = null;

  if (numbers.length !== 4) {
    return null;
  }

  getFormulaRootAssignments(numbers.length, rootCount).forEach((rootIndexes) => {
    const rootIndexSet = new Set(rootIndexes);
    const operands = numbers.map((card, index) => makeFormulaOperand(card, rootIndexSet.has(index)));

    uniqueFormulaPermutations(operands).forEach((operandPermutation) => {
      operatorSets.forEach((operatorSet) => {
        uniqueFormulaPermutations(operatorSet).forEach((operatorPermutation) => {
          evaluateFormulaShapes(operandPermutation, operatorPermutation).forEach((candidate) => {
            const distance = Math.abs(candidate.value - target);

            if (
              !best ||
              distance < best.distance - FORMULA_EPSILON ||
              (Math.abs(distance - best.distance) <= FORMULA_EPSILON && candidate.expression.length < best.expression.length)
            ) {
              best = {
                ...candidate,
                target,
                distance,
                valueText: formatFormulaValue(candidate.value),
                distanceText: formatFormulaValue(distance)
              };
            }
          });
        });
      });
    });
  });

  return best;
}

function evaluateFormulaPlayerState(playerState) {
  return {
    high: findBestFormulaForTarget(playerState, 20),
    low: findBestFormulaForTarget(playerState, 1),
    highTieCard: getFormulaTieCard(playerState, 'HIGH'),
    lowTieCard: getFormulaTieCard(playerState, 'LOW')
  };
}

function compareFormulaSide(playerOneEval, playerTwoEval, side) {
  const key = side === 'HIGH' ? 'high' : 'low';
  const first = playerOneEval[key];
  const second = playerTwoEval[key];

  if (!first || !second) {
    return first ? 'playerOne' : second ? 'playerTwo' : 'tie';
  }

  if (first.distance < second.distance - FORMULA_EPSILON) {
    return 'playerOne';
  }

  if (second.distance < first.distance - FORMULA_EPSILON) {
    return 'playerTwo';
  }

  const firstTieCard = side === 'HIGH' ? playerOneEval.highTieCard : playerOneEval.lowTieCard;
  const secondTieCard = side === 'HIGH' ? playerTwoEval.highTieCard : playerTwoEval.lowTieCard;

  if (!firstTieCard || !secondTieCard) {
    return 'tie';
  }

  if (side === 'HIGH') {
    if (firstTieCard.value !== secondTieCard.value) {
      return firstTieCard.value > secondTieCard.value ? 'playerOne' : 'playerTwo';
    }

    if (FORMULA_CARD_SUIT_RANK[firstTieCard.suit] !== FORMULA_CARD_SUIT_RANK[secondTieCard.suit]) {
      return FORMULA_CARD_SUIT_RANK[firstTieCard.suit] > FORMULA_CARD_SUIT_RANK[secondTieCard.suit] ? 'playerOne' : 'playerTwo';
    }
  } else {
    if (firstTieCard.value !== secondTieCard.value) {
      return firstTieCard.value < secondTieCard.value ? 'playerOne' : 'playerTwo';
    }

    if (FORMULA_CARD_SUIT_RANK[firstTieCard.suit] !== FORMULA_CARD_SUIT_RANK[secondTieCard.suit]) {
      return FORMULA_CARD_SUIT_RANK[firstTieCard.suit] < FORMULA_CARD_SUIT_RANK[secondTieCard.suit] ? 'playerOne' : 'playerTwo';
    }
  }

  return 'tie';
}

function normalizeFormulaChoice(choice) {
  const normalized = String(choice || '').trim().toUpperCase();

  if (!FORMULA_CHOICES.includes(normalized)) {
    throw createHttpError(400, '수식 하이 로우 선택은 HIGH, LOW, SWING 중 하나여야 합니다.');
  }

  return normalized;
}

function getFormulaChoiceSides(choice) {
  const normalized = normalizeFormulaChoice(choice);

  if (normalized === 'SWING') {
    return ['HIGH', 'LOW'];
  }

  return [normalized];
}

function getFormulaRequiredSides(choices) {
  return {
    playerOne: getFormulaChoiceSides(choices.playerOne),
    playerTwo: getFormulaChoiceSides(choices.playerTwo)
  };
}

function getFormulaEligibleSideWinner(evaluation, choices, side) {
  const eligible = ['playerOne', 'playerTwo'].filter((key) => (
    choices[key] === side || choices[key] === 'SWING'
  ));

  if (eligible.length === 1) {
    return eligible[0];
  }

  if (eligible.length === 2) {
    return compareFormulaSide(evaluation.playerOne, evaluation.playerTwo, side);
  }

  return 'tie';
}

function getFormulaOverallWinner(evaluation, choices) {
  const highWinner = getFormulaEligibleSideWinner(evaluation, choices, 'HIGH');
  const lowWinner = getFormulaEligibleSideWinner(evaluation, choices, 'LOW');
  const playerOneSwing = choices.playerOne === 'SWING';
  const playerTwoSwing = choices.playerTwo === 'SWING';

  if (playerOneSwing && playerTwoSwing) {
    if (highWinner === 'playerOne' && lowWinner === 'playerOne') {
      return { winnerKey: 'playerOne', highWinner, lowWinner };
    }

    if (highWinner === 'playerTwo' && lowWinner === 'playerTwo') {
      return { winnerKey: 'playerTwo', highWinner, lowWinner };
    }

    return { winnerKey: 'tie', highWinner, lowWinner };
  }

  if (playerOneSwing) {
    return {
      winnerKey: highWinner === 'playerOne' && lowWinner === 'playerOne' ? 'playerOne' : 'playerTwo',
      highWinner,
      lowWinner
    };
  }

  if (playerTwoSwing) {
    return {
      winnerKey: highWinner === 'playerTwo' && lowWinner === 'playerTwo' ? 'playerTwo' : 'playerOne',
      highWinner,
      lowWinner
    };
  }

  if (choices.playerOne !== choices.playerTwo) {
    return { winnerKey: 'tie', highWinner, lowWinner };
  }

  return {
    winnerKey: choices.playerOne === 'HIGH' ? highWinner : lowWinner,
    highWinner,
    lowWinner
  };
}

function getFormulaWinnerName(playerOne, playerTwo, winnerKey) {
  if (winnerKey === 'playerOne') {
    return playerOne.name;
  }

  if (winnerKey === 'playerTwo') {
    return playerTwo.name;
  }

  return '무승부 / 팟 분배';
}

function getFormulaPlayerStateSummary(playerState, revealHidden = false) {
  return {
    hidden: revealHidden ? formatFormulaCard(playerState.hidden) : 'hidden',
    openSlots: (playerState.openSlots || []).map(formatFormulaOpenSlot),
    numbers: getFormulaNumbers(playerState).map(formatFormulaCard),
    rootCount: getFormulaRootCount(playerState),
    multiplyCount: getFormulaMultiplyCount(playerState)
  };
}

function getFormulaEvaluationSummary(playerEval) {
  if (!playerEval) {
    return null;
  }

  return {
    high: playerEval.high ? {
      expression: playerEval.high.expression,
      value: playerEval.high.valueText,
      distance: playerEval.high.distanceText
    } : null,
    low: playerEval.low ? {
      expression: playerEval.low.expression,
      value: playerEval.low.valueText,
      distance: playerEval.low.distanceText
    } : null,
    highTieCard: formatFormulaCard(playerEval.highTieCard),
    lowTieCard: formatFormulaCard(playerEval.lowTieCard)
  };
}

function getFormulaHighLowSessionState(session, options = {}) {
  const revealHidden = Boolean(options.revealHidden);
  const evaluation = options.evaluation || null;
  const playerOneName = getPvpPlayerName(session.playerOneId);
  const playerTwoName = getPvpPlayerName(session.playerTwoId);

  return {
    id: session.id,
    playerOne: { id: session.playerOneId, name: playerOneName },
    playerTwo: { id: session.playerTwoId, name: playerTwoName },
    color: session.color,
    bet: session.bet,
    pot: session.bet * 2,
    roundNumber: session.roundNumber,
    stage: session.stage,
    winnerKey: session.winnerKey || '',
    winnerName: session.winnerKey ? getFormulaWinnerName({ name: playerOneName }, { name: playerTwoName }, session.winnerKey) : '',
    result: session.result || '',
    discardedCards: (session.discardedCards || []).map(formatFormulaCard),
    remainingCards: (session.deck || []).length,
    playerOneState: getFormulaPlayerStateSummary(session.playerOneState || {}, revealHidden),
    playerTwoState: getFormulaPlayerStateSummary(session.playerTwoState || {}, revealHidden),
    evaluation: evaluation ? {
      playerOne: getFormulaEvaluationSummary(evaluation.playerOne),
      playerTwo: getFormulaEvaluationSummary(evaluation.playerTwo),
      choices: evaluation.choices || {},
      highWinner: evaluation.highWinner || '',
      lowWinner: evaluation.lowWinner || '',
      winnerKey: evaluation.winnerKey || ''
    } : null
  };
}

function evaluateFormulaHighLowSession(session, choices) {
  const playerOne = evaluateFormulaPlayerState(session.playerOneState);
  const playerTwo = evaluateFormulaPlayerState(session.playerTwoState);
  const normalizedChoices = {
    playerOne: normalizeFormulaChoice(choices.playerOne),
    playerTwo: normalizeFormulaChoice(choices.playerTwo)
  };
  const outcome = getFormulaOverallWinner({ playerOne, playerTwo }, normalizedChoices);

  return {
    playerOne,
    playerTwo,
    choices: normalizedChoices,
    ...outcome
  };
}

function tokenizeFormulaExpression(expression) {
  const tokens = [];
  const text = String(expression || '');
  let index = 0;

  while (index < text.length) {
    const char = text[index];

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    const rest = text.slice(index);
    const sqrtMatch = rest.match(/^(sqrt|root|루트)/i);

    if (sqrtMatch) {
      tokens.push({ type: 'sqrt', value: sqrtMatch[0] });
      index += sqrtMatch[0].length;
      continue;
    }

    if (char === '√') {
      tokens.push({ type: 'sqrt', value: char });
      index += 1;
      continue;
    }

    if (/[0-9]/.test(char)) {
      const match = rest.match(/^\d+(?:\.\d+)?/);
      tokens.push({ type: 'number', value: Number(match[0]), raw: match[0] });
      index += match[0].length;
      continue;
    }

    if (['+', '-', '*', '/', '(', ')'].includes(char)) {
      tokens.push({ type: char, value: char });
      index += 1;
      continue;
    }

    if (char === '×' || char === 'x' || char === 'X') {
      tokens.push({ type: '*', value: '*' });
      index += 1;
      continue;
    }

    if (char === '÷') {
      tokens.push({ type: '/', value: '/' });
      index += 1;
      continue;
    }

    throw createHttpError(400, `수식에 사용할 수 없는 문자가 있습니다: ${char}`);
  }

  return tokens;
}

function mergeFormulaNode(left, operator, right) {
  const value = applyFormulaOperator(left.value, operator, right.value);

  if (value === null || !Number.isFinite(value)) {
    throw createHttpError(400, '0으로 나눌 수 없습니다.');
  }

  return {
    value,
    numbers: [...left.numbers, ...right.numbers],
    operators: [...left.operators, operator, ...right.operators],
    roots: left.roots + right.roots,
    expression: `(${left.expression} ${FORMULA_OPERATOR_LABELS[operator]} ${right.expression})`
  };
}

function parseSubmittedFormulaExpression(expression) {
  const tokens = tokenizeFormulaExpression(expression);
  let cursor = 0;

  function peek(type) {
    return tokens[cursor]?.type === type;
  }

  function consume(type) {
    if (!peek(type)) {
      throw createHttpError(400, `수식 형식이 올바르지 않습니다. ${type} 토큰이 필요합니다.`);
    }

    cursor += 1;
    return tokens[cursor - 1];
  }

  function parseRootNumber() {
    if (peek('(')) {
      consume('(');
      const token = consume('number');
      consume(')');
      return token;
    }

    return consume('number');
  }

  function parsePrimary() {
    if (peek('number')) {
      const token = consume('number');
      return {
        value: token.value,
        numbers: [token.value],
        operators: [],
        roots: 0,
        expression: token.raw
      };
    }

    if (peek('sqrt')) {
      consume('sqrt');
      const token = parseRootNumber();

      if (token.value < 0) {
        throw createHttpError(400, '음수에는 루트를 적용할 수 없습니다.');
      }

      return {
        value: Math.sqrt(token.value),
        numbers: [token.value],
        operators: [],
        roots: 1,
        expression: `√(${token.raw})`
      };
    }

    if (peek('(')) {
      consume('(');
      const node = parseExpression();
      consume(')');
      return node;
    }

    throw createHttpError(400, '수식 형식이 올바르지 않습니다.');
  }

  function parseTerm() {
    let node = parsePrimary();

    while (peek('*') || peek('/')) {
      const operator = tokens[cursor].type;
      cursor += 1;
      node = mergeFormulaNode(node, operator, parsePrimary());
    }

    return node;
  }

  function parseExpression() {
    let node = parseTerm();

    while (peek('+') || peek('-')) {
      const operator = tokens[cursor].type;
      cursor += 1;
      node = mergeFormulaNode(node, operator, parseTerm());
    }

    return node;
  }

  if (!tokens.length) {
    throw createHttpError(400, '수식을 입력해야 합니다.');
  }

  const result = parseExpression();

  if (cursor !== tokens.length) {
    throw createHttpError(400, '수식 끝에 해석할 수 없는 내용이 있습니다.');
  }

  return result;
}

function numberCounts(values) {
  return values.reduce((counts, value) => {
    const key = String(value);
    counts[key] = (counts[key] || 0) + 1;
    return counts;
  }, {});
}

function countsMatch(first, second) {
  const keys = new Set([...Object.keys(first), ...Object.keys(second)]);

  return [...keys].every((key) => first[key] === second[key]);
}

function validateSubmittedFormulaNode(playerState, node) {
  const availableNumbers = getFormulaNumbers(playerState).map((card) => card.value);
  const submittedNumberCounts = numberCounts(node.numbers);
  const availableNumberCounts = numberCounts(availableNumbers);
  const operatorKey = [...node.operators].sort().join('');
  const validOperatorKeys = new Set(getFormulaOperatorSets(getFormulaMultiplyCount(playerState))
    .map((operators) => [...operators].sort().join('')));

  if (!countsMatch(submittedNumberCounts, availableNumberCounts)) {
    throw createHttpError(400, `수식은 보유 숫자 ${availableNumbers.join(', ')}를 정확히 한 번씩 사용해야 합니다.`);
  }

  if (node.roots !== getFormulaRootCount(playerState)) {
    throw createHttpError(400, `루트는 ${getFormulaRootCount(playerState)}번 사용해야 합니다.`);
  }

  if (!validOperatorKeys.has(operatorKey)) {
    throw createHttpError(400, '사용한 기호 조합이 보유 기호 카드와 맞지 않습니다.');
  }
}

function evaluateSubmittedFormula(playerState, expression, target) {
  const node = parseSubmittedFormulaExpression(expression);
  validateSubmittedFormulaNode(playerState, node);

  const distance = Math.abs(node.value - target);

  return {
    value: node.value,
    expression: node.expression,
    target,
    distance,
    valueText: formatFormulaValue(node.value),
    distanceText: formatFormulaValue(distance)
  };
}

function evaluateSubmittedFormulaForSide(playerState, expression, target, label) {
  try {
    return evaluateSubmittedFormula(playerState, expression, target);
  } catch (err) {
    if (err.status) {
      throw createHttpError(err.status, `${label}: ${err.message}`);
    }

    throw err;
  }
}

function evaluateSubmittedPlayerFormulas(playerState, formulas, requiredSides, playerLabel) {
  const highFormula = String(formulas.high || '').trim();
  const lowFormula = String(formulas.low || '').trim();
  const needsHigh = requiredSides.includes('HIGH');
  const needsLow = requiredSides.includes('LOW');

  if (needsHigh && !highFormula) {
    throw createHttpError(400, `${playerLabel} HIGH 수식을 입력해야 합니다.`);
  }

  if (needsLow && !lowFormula) {
    throw createHttpError(400, `${playerLabel} LOW 수식을 입력해야 합니다.`);
  }

  return {
    high: needsHigh ? evaluateSubmittedFormulaForSide(playerState, highFormula, 20, `${playerLabel} HIGH`) : null,
    low: needsLow ? evaluateSubmittedFormulaForSide(playerState, lowFormula, 1, `${playerLabel} LOW`) : null,
    highTieCard: getFormulaTieCard(playerState, 'HIGH'),
    lowTieCard: getFormulaTieCard(playerState, 'LOW')
  };
}

function evaluateFormulaHighLowSubmission(session, choices, formulas) {
  const normalizedChoices = {
    playerOne: normalizeFormulaChoice(choices.playerOne),
    playerTwo: normalizeFormulaChoice(choices.playerTwo)
  };
  const requiredSides = getFormulaRequiredSides(normalizedChoices);
  const playerOne = evaluateSubmittedPlayerFormulas(
    session.playerOneState,
    formulas.playerOne || {},
    requiredSides.playerOne,
    'Player 1'
  );
  const playerTwo = evaluateSubmittedPlayerFormulas(
    session.playerTwoState,
    formulas.playerTwo || {},
    requiredSides.playerTwo,
    'Player 2'
  );
  const outcome = getFormulaOverallWinner({ playerOne, playerTwo }, normalizedChoices);

  return {
    playerOne,
    playerTwo,
    choices: normalizedChoices,
    ...outcome
  };
}

function getFormulaSuggestedChoice(key, highWinner, lowWinner, playerEval) {
  if (highWinner === key && lowWinner === key) {
    return 'SWING';
  }

  if (highWinner === key) {
    return 'HIGH';
  }

  if (lowWinner === key) {
    return 'LOW';
  }

  const highDistance = playerEval?.high?.distance ?? Number.POSITIVE_INFINITY;
  const lowDistance = playerEval?.low?.distance ?? Number.POSITIVE_INFINITY;
  return highDistance <= lowDistance ? 'HIGH' : 'LOW';
}

function previewFormulaHighLowSession(session) {
  const playerOne = evaluateFormulaPlayerState(session.playerOneState);
  const playerTwo = evaluateFormulaPlayerState(session.playerTwoState);
  const sideEvaluation = { playerOne, playerTwo };
  const highWinner = compareFormulaSide(playerOne, playerTwo, 'HIGH');
  const lowWinner = compareFormulaSide(playerOne, playerTwo, 'LOW');
  const choices = {
    playerOne: getFormulaSuggestedChoice('playerOne', highWinner, lowWinner, playerOne),
    playerTwo: getFormulaSuggestedChoice('playerTwo', highWinner, lowWinner, playerTwo)
  };
  const outcome = getFormulaOverallWinner(sideEvaluation, choices);

  return {
    playerOne,
    playerTwo,
    choices,
    suggestedChoices: choices,
    suggestedWinnerKey: outcome.winnerKey,
    highWinner,
    lowWinner,
    winnerKey: outcome.winnerKey
  };
}

function formatFormulaHighLowRound(roundNumber) {
  return Number.isFinite(roundNumber) && roundNumber > 0
    ? `${roundNumber}라운드`
    : '라운드 미입력';
}

function formatFormulaHighLowRevealRule(roundNumber) {
  return Number.isFinite(roundNumber) && roundNumber >= 25
    ? '하이/로우 선택 동시 공개'
    : '하이/로우 선택 일반 공개';
}

function getFormulaHighLowWinnerLabel(settlement) {
  if (settlement.winnerKey === 'playerOne') {
    return settlement.playerOne.name;
  }

  if (settlement.winnerKey === 'playerTwo') {
    return settlement.playerTwo.name;
  }

  return '무승부 / 팟 분배';
}

function getFormulaHighLowPayoutText(settlement) {
  if (settlement.winnerKey === 'tie') {
    return `${settlement.playerOne.name} ${settlement.bet}개, ${settlement.playerTwo.name} ${settlement.bet}개`;
  }

  return `${getFormulaHighLowWinnerLabel(settlement)} ${settlement.pot}개`;
}

function makeFormulaHighLowCardsText(name, playerState, playerEval, choice) {
  const hidden = formatFormulaCard(playerState.hidden);
  const open = (playerState.openSlots || []).map(formatFormulaOpenSlot).join(', ');
  const high = playerEval?.high;
  const low = playerEval?.low;

  return `${name}
선택: ${choice || '-'}
히든: ${hidden}
오픈: ${open || '-'}
하이 수식: ${high ? `${high.expression} = ${high.valueText} (20까지 ${high.distanceText})` : '-'}
로우 수식: ${low ? `${low.expression} = ${low.valueText} (1까지 ${low.distanceText})` : '-'}
하이 동점 카드: ${formatFormulaCard(playerEval?.highTieCard)}
로우 동점 카드: ${formatFormulaCard(playerEval?.lowTieCard)}`;
}

function makeFormulaHighLowProgressLog(playerOne, playerTwo, session, action) {
  return `수식 하이 로우 진행
행동: ${action}
라운드: ${formatFormulaHighLowRound(session.roundNumber)}
공개 규칙: ${formatFormulaHighLowRevealRule(session.roundNumber)}
플레이어 1: ${playerOne.name}
플레이어 2: ${playerTwo.name}
각자 베팅: ${session.color} ${session.bet}개
단계: ${session.stage}

${playerOne.name}
히든: ${formatFormulaCard(session.playerOneState.hidden)}
오픈: ${(session.playerOneState.openSlots || []).map(formatFormulaOpenSlot).join(', ') || '-'}

${playerTwo.name}
히든: ${formatFormulaCard(session.playerTwoState.hidden)}
오픈: ${(session.playerTwoState.openSlots || []).map(formatFormulaOpenSlot).join(', ') || '-'}

버린 기호: ${(session.discardedCards || []).map(formatFormulaCard).join(', ') || '-'}
남은 카드: ${(session.deck || []).length}`;
}

function makeFormulaHighLowProgressPublicLog(playerOne, playerTwo, session, action) {
  return `수식 하이 로우 진행
행동: ${action}
라운드: ${formatFormulaHighLowRound(session.roundNumber)}
공개 규칙: ${formatFormulaHighLowRevealRule(session.roundNumber)}
플레이어 1: ${playerOne.name}
플레이어 2: ${playerTwo.name}
각자 베팅: ${session.color} ${session.bet}개
단계: ${session.stage}

${playerOne.name} 오픈: ${(session.playerOneState.openSlots || []).map(formatFormulaOpenSlot).join(', ') || '-'}
${playerTwo.name} 오픈: ${(session.playerTwoState.openSlots || []).map(formatFormulaOpenSlot).join(', ') || '-'}
버린 기호: ${(session.discardedCards || []).map(formatFormulaCard).join(', ') || '-'}`;
}

function makeFormulaHighLowPlayerInputLog(player, playerState, session, playerIndex) {
  const numberCards = getFormulaNumbers(playerState);
  const numbers = numberCards.map((card) => card.value).join(', ');
  const numberDetails = numberCards.map(formatFormulaCard).join(', ');
  const rootCount = getFormulaRootCount(playerState);
  const multiplyCount = getFormulaMultiplyCount(playerState);
  const operatorSets = getFormulaOperatorSets(multiplyCount)
    .map((operators) => operators.map((operator) => FORMULA_OPERATOR_LABELS[operator]).join(' '))
    .join(' / ');

  return `로그${playerIndex}
수식 하이 로우 입력 안내
플레이어${playerIndex}: ${player.name}
라운드: ${formatFormulaHighLowRound(session.roundNumber)}
히든: ${formatFormulaCard(playerState.hidden)}
오픈: ${(playerState.openSlots || []).map(formatFormulaOpenSlot).join(', ') || '-'}
숫자: ${numbers}
숫자 카드: ${numberDetails}
루트 사용 횟수: ${rootCount}
가능 기호 조합: ${operatorSets}

플레이어 입력
선택: HIGH / LOW / SWING
HIGH 수식:
LOW 수식:

참고: 수식에는 사용 숫자를 정확히 한 번씩 모두 넣어야 합니다. 루트가 있으면 sqrt(숫자), root(숫자), 루트(숫자), √숫자 형식으로 입력하세요.`;
}

function makeFormulaHighLowLog(settlement) {
  return `수식 하이 로우 정산
라운드: ${formatFormulaHighLowRound(settlement.roundNumber)}
공개 규칙: ${formatFormulaHighLowRevealRule(settlement.roundNumber)}
칩 색: ${settlement.color}
플레이어 1: ${settlement.playerOne.name}
플레이어 2: ${settlement.playerTwo.name}
각자 베팅: ${settlement.bet}개
총 팟: ${settlement.pot}개
하이 승자: ${getFormulaWinnerName(settlement.playerOne, settlement.playerTwo, settlement.evaluation.highWinner)}
로우 승자: ${getFormulaWinnerName(settlement.playerOne, settlement.playerTwo, settlement.evaluation.lowWinner)}
승자: ${getFormulaHighLowWinnerLabel(settlement)}
지급: ${getFormulaHighLowPayoutText(settlement)}

${makeFormulaHighLowCardsText(settlement.playerOne.name, settlement.session.playerOneState, settlement.evaluation.playerOne, settlement.evaluation.choices.playerOne)}

${makeFormulaHighLowCardsText(settlement.playerTwo.name, settlement.session.playerTwoState, settlement.evaluation.playerTwo, settlement.evaluation.choices.playerTwo)}

현재
${settlement.playerOne.name}: ${chipStr(settlement.playerOne)}
${settlement.playerTwo.name}: ${chipStr(settlement.playerTwo)}`;
}

function makeFormulaHighLowPublicLog(settlement) {
  return `수식 하이 로우
라운드: ${formatFormulaHighLowRound(settlement.roundNumber)}
공개 규칙: ${formatFormulaHighLowRevealRule(settlement.roundNumber)}
플레이어 1: ${settlement.playerOne.name}
플레이어 2: ${settlement.playerTwo.name}
각자 베팅: ${settlement.color} ${settlement.bet}개
총 팟: ${settlement.color} ${settlement.pot}개
하이 승자: ${getFormulaWinnerName(settlement.playerOne, settlement.playerTwo, settlement.evaluation.highWinner)}
로우 승자: ${getFormulaWinnerName(settlement.playerOne, settlement.playerTwo, settlement.evaluation.lowWinner)}
승자: ${getFormulaHighLowWinnerLabel(settlement)}
지급: ${getFormulaHighLowPayoutText(settlement)}

${makeFormulaHighLowCardsText(settlement.playerOne.name, settlement.session.playerOneState, settlement.evaluation.playerOne, settlement.evaluation.choices.playerOne)}

${makeFormulaHighLowCardsText(settlement.playerTwo.name, settlement.session.playerTwoState, settlement.evaluation.playerTwo, settlement.evaluation.choices.playerTwo)}`;
}

const COMBINATION_OPERATOR_CARDS = ['+', '-', '*', '/', '^'];
const COMBINATION_OPERATOR_LABELS = {
  '+': '+',
  '-': '-',
  '*': '×',
  '/': '÷',
  '^': '^'
};
const COMBINATION_OPERATOR_ALIASES = {
  '+': '+',
  '＋': '+',
  '-': '-',
  '－': '-',
  '*': '*',
  '×': '*',
  'x': '*',
  'X': '*',
  '/': '/',
  '÷': '/',
  '^': '^',
  '＾': '^'
};
const COMBINATION_TARGET_MIN = 0;
const COMBINATION_TARGET_MAX = 100;

function drawCombinationNumberCard() {
  return Math.floor(Math.random() * 11);
}

function createCombinationParticipantState(playerId) {
  return {
    playerId: Number(playerId),
    numbers: Array.from({ length: 6 }, drawCombinationNumberCard),
    submission: '',
    evaluation: null
  };
}

function drawCombinationTarget() {
  return COMBINATION_TARGET_MIN + Math.floor(Math.random() * (COMBINATION_TARGET_MAX - COMBINATION_TARGET_MIN + 1));
}

function normalizeCombinationParticipantIds({ participantIds = [], playerId, opponentId } = {}) {
  const sourceIds = Array.isArray(participantIds) && participantIds.length
    ? participantIds
    : [playerId, opponentId];

  return [...new Set(sourceIds.map(Number))]
    .filter((id) => Number.isFinite(id));
}

function normalizeCombinationOperator(operator) {
  const normalized = COMBINATION_OPERATOR_ALIASES[String(operator || '')];

  if (!normalized || !COMBINATION_OPERATOR_CARDS.includes(normalized)) {
    throw createHttpError(400, `사용할 수 없는 기호입니다: ${operator}`);
  }

  return normalized;
}

function tokenizeCombinationExpression(expression) {
  const tokens = [];
  const text = String(expression || '');
  let index = 0;

  while (index < text.length) {
    const char = text[index];

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    if (/[0-9]/.test(char)) {
      const match = text.slice(index).match(/^\d+/);
      tokens.push({ type: 'number', value: Number(match[0]), raw: match[0] });
      index += match[0].length;
      continue;
    }

    if (COMBINATION_OPERATOR_ALIASES[char]) {
      tokens.push({ type: 'operator', value: normalizeCombinationOperator(char), raw: char });
      index += 1;
      continue;
    }

    throw createHttpError(400, `수식에 사용할 수 없는 문자가 있습니다: ${char}`);
  }

  return tokens;
}

function parseCombinationExpression(expression) {
  const tokens = tokenizeCombinationExpression(expression);

  if (
    tokens.length !== 5 ||
    tokens[0]?.type !== 'number' ||
    tokens[1]?.type !== 'operator' ||
    tokens[2]?.type !== 'number' ||
    tokens[3]?.type !== 'operator' ||
    tokens[4]?.type !== 'number'
  ) {
    throw createHttpError(400, '수식은 숫자, 기호, 숫자, 기호, 숫자 형식이어야 합니다.');
  }

  return {
    numbers: [tokens[0].value, tokens[2].value, tokens[4].value],
    operators: [tokens[1].value, tokens[3].value],
    tokens,
    expression: tokens
      .map((token) => token.type === 'operator' ? COMBINATION_OPERATOR_LABELS[token.value] : token.raw)
      .join(' ')
  };
}

function hasEnoughCombinationNumbers(availableNumbers, submittedNumbers) {
  const available = numberCounts(availableNumbers);
  const submitted = numberCounts(submittedNumbers);

  return Object.keys(submitted).every((key) => (available[key] || 0) >= submitted[key]);
}

function applyCombinationOperator(left, operator, right) {
  if (operator === '+') {
    return left + right;
  }

  if (operator === '-') {
    return left - right;
  }

  if (operator === '*') {
    return left * right;
  }

  if (operator === '/') {
    if (Math.abs(right) < FORMULA_EPSILON) {
      throw createHttpError(400, '0으로 나눌 수 없습니다.');
    }

    return left / right;
  }

  if (operator === '^') {
    return Math.pow(left, right);
  }

  throw createHttpError(400, `사용할 수 없는 기호입니다: ${operator}`);
}

function evaluateCombinationParsedExpression(parsed) {
  const precedence = { '+': 1, '-': 1, '*': 2, '/': 2, '^': 3 };
  const values = [];
  const operators = [];

  function applyTopOperator() {
    const operator = operators.pop();
    const right = values.pop();
    const left = values.pop();

    if (left === undefined || right === undefined) {
      throw createHttpError(400, '수식 형식이 올바르지 않습니다.');
    }

    const value = applyCombinationOperator(left, operator, right);

    if (!Number.isFinite(value)) {
      throw createHttpError(400, '수식 결과가 너무 크거나 계산할 수 없습니다.');
    }

    values.push(value);
  }

  parsed.tokens.forEach((token) => {
    if (token.type === 'number') {
      values.push(token.value);
      return;
    }

    while (
      operators.length &&
      (
        precedence[operators[operators.length - 1]] > precedence[token.value] ||
        (
          precedence[operators[operators.length - 1]] === precedence[token.value] &&
          token.value !== '^'
        )
      )
    ) {
      applyTopOperator();
    }

    operators.push(token.value);
  });

  while (operators.length) {
    applyTopOperator();
  }

  if (values.length !== 1 || !Number.isFinite(values[0])) {
    throw createHttpError(400, '수식 결과를 계산할 수 없습니다.');
  }

  return values[0];
}

function evaluateCombinationFormula(playerState, expression, target, playerLabel) {
  try {
    const parsed = parseCombinationExpression(expression);
    const invalidNumbers = parsed.numbers.filter((value) => !Number.isInteger(value) || value < 0 || value > 10);

    if (invalidNumbers.length) {
      throw createHttpError(400, '숫자 카드는 0부터 10까지만 사용할 수 있습니다.');
    }

    if (!hasEnoughCombinationNumbers(playerState.numbers || [], parsed.numbers)) {
      throw createHttpError(400, `보유 숫자 ${formatCombinationNumbers(playerState.numbers)} 중 3개만 사용할 수 있습니다.`);
    }

    const value = evaluateCombinationParsedExpression(parsed);
    const distance = Math.abs(value - target);

    return {
      expression: parsed.expression,
      value,
      valueText: formatFormulaValue(value),
      distance,
      distanceText: formatFormulaValue(distance),
      target
    };
  } catch (err) {
    if (err.status) {
      throw createHttpError(err.status, `${playerLabel}: ${err.message}`);
    }

    throw err;
  }
}

function formatCombinationNumbers(numbers) {
  return Array.isArray(numbers) && numbers.length ? numbers.join(', ') : '-';
}

function formatCombinationOperatorsWithInputAliases() {
  return COMBINATION_OPERATOR_CARDS
    .map((operator) => {
      if (operator === '*') {
        return '×(*)';
      }

      if (operator === '/') {
        return '÷(/)';
      }

      return COMBINATION_OPERATOR_LABELS[operator];
    })
    .join(', ');
}

function getCombinationPlayerName(playerId) {
  return getPlayerById(playerId)?.name || `#${playerId}`;
}

function getCombinationParticipantState(session, playerId) {
  return (session.participantStates || [])
    .find((state) => Number(state.playerId) === Number(playerId));
}

function getCombinationWinnerIds(participantStates) {
  const evaluatedStates = (participantStates || []).filter((state) => state.evaluation);
  const bestDistance = evaluatedStates.reduce((best, state) => (
    state.evaluation.distance < best ? state.evaluation.distance : best
  ), Number.POSITIVE_INFINITY);

  return evaluatedStates
    .filter((state) => Math.abs(state.evaluation.distance - bestDistance) <= FORMULA_EPSILON)
    .map((state) => Number(state.playerId));
}

function getCombinationWinnerNames(winnerIds) {
  return (winnerIds || []).map(getCombinationPlayerName).join(', ') || '-';
}

function getCombinationBettingSessionState(session) {
  const resultChipTextByPlayerId = session.resultChipTextByPlayerId || {};
  const participants = (session.participantStates || []).map((state) => ({
    id: Number(state.playerId),
    name: getCombinationPlayerName(state.playerId),
    numbers: state.numbers || [],
    operators: COMBINATION_OPERATOR_CARDS.map((operator) => COMBINATION_OPERATOR_LABELS[operator]),
    submission: state.submission || '',
    resultChipText: resultChipTextByPlayerId[Number(state.playerId)] || '',
    evaluation: state.evaluation ? {
      expression: state.evaluation.expression,
      value: state.evaluation.valueText,
      distance: state.evaluation.distanceText
    } : null
  }));

  return {
    id: session.id,
    gameType: 'combinationbetting',
    participantCount: participants.length,
    participants,
    color: session.color,
    bet: session.bet,
    pot: session.pot || session.bet * participants.length,
    target: session.target,
    operatorCards: COMBINATION_OPERATOR_CARDS.map((operator) => COMBINATION_OPERATOR_LABELS[operator]),
    roundNumber: session.roundNumber,
    stage: session.stage,
    winnerIds: session.winnerIds || [],
    winnerNames: getCombinationWinnerNames(session.winnerIds || []),
    result: session.result || '',
    resultChipTexts: session.resultChipTexts || []
  };
}

function makeCombinationBettingStartLog(participants, session) {
  const hands = (session.participantStates || [])
    .map((state, index) => {
      const player = participants.find((candidate) => Number(candidate.id) === Number(state.playerId));
      return `${index + 1}. ${player?.name || getCombinationPlayerName(state.playerId)}: ${formatCombinationNumbers(state.numbers)}`;
    })
    .join('\n');

  return `콤비네이션 베팅 시작
목표 숫자: ${formatFormulaValue(session.target)}
참가자: ${participants.map((player) => player.name).join(', ')}
각자 베팅: ${session.color} ${session.bet}개
총 팟: ${session.color} ${session.pot}개
사용 가능 기호: ${formatCombinationOperatorsWithInputAliases()}
규칙: 보유 숫자 6개 중 3개와 기호 2개를 골라 숫자, 기호, 숫자, 기호, 숫자 형식의 수식을 제출합니다.

숫자 카드
${hands}`;
}

function makeCombinationBettingStartPublicLog(participants, session) {
  const playerSections = (session.participantStates || [])
    .map((state, index) => {
      return `[플레이어 ${index + 1} 제출 정보]
숫자: ${formatCombinationNumbers(state.numbers)}
타겟: ${formatFormulaValue(session.target)}
연산자: ${formatCombinationOperatorsWithInputAliases()}
제출 형식: 숫자, 연산자, 숫자, 연산자, 숫자`;
    })
    .join('\n\n');

  return `콤비네이션 베팅 시작
[공통 공개]
목표 숫자: ${formatFormulaValue(session.target)}
참가자: ${participants.map((player) => player.name).join(', ')}
각자 베팅: ${session.color} ${session.bet}개
총 팟: ${session.color} ${session.pot}개
사용 가능 기호: ${formatCombinationOperatorsWithInputAliases()}

${playerSections}`;
}

function makeCombinationBettingSettlementLog(participants, session, payouts) {
  const results = (session.participantStates || [])
    .map((state) => {
      const player = participants.find((candidate) => Number(candidate.id) === Number(state.playerId));
      const evaluation = state.evaluation;
      const payout = payouts.get(Number(state.playerId)) || 0;

      return `${player?.name || getCombinationPlayerName(state.playerId)}
숫자: ${formatCombinationNumbers(state.numbers)}
수식: ${evaluation?.expression || state.submission || '-'}
결과값: ${evaluation?.valueText || '-'}
거리: ${evaluation?.distanceText || '-'}
지급: ${session.color} ${payout}개`;
    })
    .join('\n\n');

  const balances = participants
    .map((player) => `${player.name}: ${chipStr(player)}`)
    .join('\n');
  const totalPaid = [...payouts.values()].reduce((sum, amount) => sum + amount, 0);
  const discarded = Math.max(0, session.pot - totalPaid);

  return `콤비네이션 베팅 정산
목표 숫자: ${formatFormulaValue(session.target)}
각자 베팅: ${session.color} ${session.bet}개
총 팟: ${session.color} ${session.pot}개
지급 합계: ${session.color} ${totalPaid}개
버림: ${session.color} ${discarded}개
승자: ${getCombinationWinnerNames(session.winnerIds)}
결과: ${session.result}
${session.resultChipTexts?.length ? `결과 칩:\n${session.resultChipTexts.join('\n')}` : ''}

${results}

현재
${balances}`;
}

function makeCombinationBettingSettlementPublicLog(participants, session, payouts) {
  const totalPaid = [...payouts.values()].reduce((sum, amount) => sum + amount, 0);
  const discarded = Math.max(0, session.pot - totalPaid);

  return `콤비네이션 베팅 정산
목표 숫자: ${formatFormulaValue(session.target)}
각자 베팅: ${session.color} ${session.bet}개
총 팟: ${session.color} ${session.pot}개
지급 합계: ${session.color} ${totalPaid}개
버림: ${session.color} ${discarded}개
승자: ${getCombinationWinnerNames(session.winnerIds)}${session.resultChipTexts?.length ? `\n결과 칩:\n${session.resultChipTexts.join('\n')}` : ''}`;
}

// =========================
// 🌐 페이지 라우팅
// =========================
app.get('/', (req, res) => {
  res.redirect('/dashboard.html');
});

// =========================
// API
// =========================

// 플레이어 조회
app.get('/players', async (req, res) => {
  try {
    await Promise.all([loadSheetCached(), loadRatesCached()]);
    res.json(serializePlayers(players));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '플레이어 데이터를 불러오지 못했습니다.' });
  }
});

app.get('/teams', async (req, res) => {
  try {
    await Promise.all([loadSheetCached(), loadRatesCached()]);
    res.json({
      rates,
      teams: buildTeamTotals(players),
      players: serializePlayers(players)
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '팀 데이터를 불러오지 못했습니다.' });
  }
});

// 마지막 로그 조회
app.get('/logtext', async (req, res) => {
  try {
    const logs = await loadLogHistory(1);
    res.send(logs[0]?.text || lastLogText);
  } catch (err) {
    console.error(err);
    res.status(500).send('로그를 불러오지 못했습니다.');
  }
});

app.get('/logs', async (req, res) => {
  try {
    const logs = await loadLogHistory(req.query.limit);
    res.json(logs);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '로그를 불러오지 못했습니다.', detail: getErrorDetail(err) });
  }
});

app.get('/debug/env', (req, res) => {
  res.json(getEnvDebugInfo());
});

app.get('/reset-game-limits', async (req, res) => {
  try {
    const reset = await ensureDailyGameLimitReset();
    const today = getKoreaDateKey();

    res.json({
      ok: true,
      reset,
      date: today,
      range: `${PLAYER_SHEET_NAME}!${PLAYER_GAME_LIMIT_RESET_RANGE_START_COLUMN}2:${PLAYER_GAME_LIMIT_RESET_RANGE_END_COLUMN}x`,
      value: DEFAULT_DAILY_GAME_LIMIT
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '게임 횟수 리셋 중 오류가 발생했습니다.' });
  }
});

app.post('/reset-game-limits', async (req, res) => {
  try {
    const reset = await ensureDailyGameLimitReset();
    const today = getKoreaDateKey();

    res.json({
      ok: true,
      reset,
      date: today,
      range: `${PLAYER_SHEET_NAME}!${PLAYER_GAME_LIMIT_RESET_RANGE_START_COLUMN}2:${PLAYER_GAME_LIMIT_RESET_RANGE_END_COLUMN}x`,
      value: DEFAULT_DAILY_GAME_LIMIT
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '게임 횟수 리셋 중 오류가 발생했습니다.' });
  }
});

app.post('/logs/test', async (req, res) => {
  try {
    const log = await addLog('test', `로그 저장 테스트\ncreatedAt: ${new Date().toISOString()}`);
    res.json({ ok: true, log, logs: [log] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '로그 저장 테스트 실패', detail: getErrorDetail(err) });
  }
});

async function handleBalanceLog(req, res, playerId) {
  try {
    await Promise.all([
      loadPlayersByIds([playerId]),
      loadRatesCached()
    ]);

    const player = getPlayerById(playerId);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    const log = await addLog('balance', makeBalanceLog(player), makeBalancePublicLog(player));

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '잔고 로그를 생성하지 못했습니다.', detail: getErrorDetail(err) });
  }
}

app.post('/balance-log', async (req, res) => {
  return handleBalanceLog(req, res, req.body.playerId);
});

app.post('/players/:playerId/balance-log', async (req, res) => {
  return handleBalanceLog(req, res, req.params.playerId);
});

// 관리자 칩 조정
app.post('/chips/adjust', async (req, res) => {
  try {
    const { playerId, color, amount, action } = req.body;
    await loadPlayersByIds([playerId]);

    const player = getPlayerById(playerId);
    const adjustAmount = toChipAmount(amount);
    const normalizedAction = String(action || '').toLowerCase();

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!['add', 'subtract'].includes(normalizedAction)) {
      return res.status(400).json({ error: '조정 방식이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(adjustAmount) || adjustAmount <= 0) {
      return res.status(400).json({ error: '조정 수량이 올바르지 않습니다.' });
    }

    if (normalizedAction === 'subtract' && player[color] < adjustAmount) {
      return res.status(400).json({ error: '보유 칩이 부족합니다.' });
    }

    if (normalizedAction === 'add') {
      player[color] += adjustAmount;
    } else {
      player[color] -= adjustAmount;
    }

    const [log] = await Promise.all([
      addLog(
        'chip-adjust',
        makeChipAdjustLog(player, color, adjustAmount, normalizedAction),
        makeChipAdjustPublicLog(player, color, adjustAmount, normalizedAction)
      ),
      savePlayers([player])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '칩 조정 중 오류가 발생했습니다.' });
  }
});

// 환율 조회
app.get('/rates', async (req, res) => {
  try {
    const nextRates = await loadRatesCached();
    res.json(nextRates);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '칩 가치를 불러오지 못했습니다.' });
  }
});

// 환율 설정
app.post('/rates', async (req, res) => {
  try {
    const nextRates = Object.fromEntries(
      COLORS.map((color) => [color, Number(req.body[color])])
    );

    if (COLORS.some((color) => !Number.isFinite(nextRates[color]) || nextRates[color] <= 0)) {
      return res.status(400).json({ error: '비율은 1 이상의 숫자여야 합니다.' });
    }

    const savedRates = await saveRates(nextRates);

    res.json({ ok: true, rates: savedRates });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '칩 가치를 저장하지 못했습니다.' });
  }
});

app.post('/formula-high-low/start', async (req, res) => {
  try {
    const { color, amount, roundNumber } = req.body;
    const bet = toChipAmount(amount);
    const parsedRoundNumber = toChipAmount(roundNumber);
    const participantIds = normalizeCombinationParticipantIds(req.body);

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (participantIds.length < 2) {
      return res.status(400).json({ error: '콤비네이션 베팅은 2명 이상 참여해야 합니다.' });
    }

    await loadPlayersByIds(participantIds);

    const participants = participantIds.map(getPlayerById);

    if (participants.some((player) => !player)) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (participants.some((player) => player[color] < bet)) {
      return res.status(400).json({ error: '모든 참가자가 베팅할 칩을 보유해야 합니다.' });
    }

    participants.forEach((player) => {
      player[color] -= bet;
    });

    const session = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      participantIds,
      color,
      bet,
      target: drawCombinationTarget(),
      operatorCards: [...COMBINATION_OPERATOR_CARDS],
      participantStates: participantIds.map(createCombinationParticipantState),
      pot: bet * participantIds.length,
      roundNumber: Number.isFinite(parsedRoundNumber) && parsedRoundNumber > 0 ? parsedRoundNumber : null,
      stage: 'input',
      winnerIds: [],
      result: '',
      memo: ''
    };

    const [log] = await Promise.all([
      addLog(
        'formulahighlow',
        makeCombinationBettingStartLog(participants, session),
        makeCombinationBettingStartPublicLog(participants, session)
      ),
      saveFormulaHighLowSession(session),
      savePlayers(participants)
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({
      ok: true,
      state: getCombinationBettingSessionState(session),
      log: log.text,
      logs: [log],
      players: serializePlayers(participants)
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '콤비네이션 베팅 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/formula-high-low/final-card', async (req, res) => {
  try {
    const { sessionId } = req.body;
    const session = await getFormulaHighLowSession(sessionId);

    if (!session) {
      return res.status(400).json({ error: '진행 중인 콤비네이션 베팅이 없습니다.' });
    }

    res.json({
      ok: true,
      state: getCombinationBettingSessionState(session),
      log: '',
      logs: [],
      players: []
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '콤비네이션 베팅 상태 조회 중 오류가 발생했습니다.' });
  }
});

async function handleFormulaHighLowResolve(req, res) {
  try {
    const {
      sessionId,
      submissions = []
    } = req.body;
    const session = await getFormulaHighLowSession(sessionId);

    if (!session) {
      return res.status(400).json({ error: '진행 중인 콤비네이션 베팅이 없습니다.' });
    }

    if (session.stage !== 'input') {
      return res.status(400).json({ error: '정산할 수 없는 단계입니다.' });
    }

    await loadPlayersByIds(session.participantIds);

    const participants = session.participantIds.map(getPlayerById);

    if (participants.some((player) => !player)) {
      await deleteFormulaHighLowSession(session.id);
      return res.status(404).json({ error: '플레이어를 찾을 수 없어 세션을 종료했습니다.' });
    }

    const submissionByPlayerId = new Map(
      (Array.isArray(submissions) ? submissions : [])
        .map((entry) => [Number(entry?.playerId), String(entry?.formula || '').trim()])
        .filter(([playerId]) => Number.isFinite(playerId))
    );
    const missingPlayers = session.participantIds
      .filter((playerId) => !submissionByPlayerId.get(Number(playerId)))
      .map(getCombinationPlayerName);

    if (missingPlayers.length) {
      return res.status(400).json({ error: `수식을 입력하지 않은 참가자가 있습니다: ${missingPlayers.join(', ')}` });
    }

    session.participantStates = (session.participantStates || []).map((state) => {
      const player = participants.find((candidate) => Number(candidate.id) === Number(state.playerId));
      const formula = submissionByPlayerId.get(Number(state.playerId));
      const evaluation = evaluateCombinationFormula(
        state,
        formula,
        session.target,
        player?.name || getCombinationPlayerName(state.playerId)
      );

      return {
        ...state,
        submission: formula,
        evaluation
      };
    });

    const winnerIds = getCombinationWinnerIds(session.participantStates);

    if (!winnerIds.length) {
      throw createHttpError(400, '승자를 계산할 수 없습니다.');
    }

    const payouts = new Map();
    const baseShare = Math.floor(session.pot / winnerIds.length);

    winnerIds.forEach((winnerId) => {
      const player = getPlayerById(winnerId);
      const payout = baseShare;

      player[session.color] += payout;
      payouts.set(Number(winnerId), payout);
    });

    session.stage = 'done';
    session.winnerIds = winnerIds;
    session.memo = '';
    session.result = winnerIds.length === 1
      ? `${getCombinationPlayerName(winnerIds[0])} 승리`
      : `동점 / 팟 분배: ${getCombinationWinnerNames(winnerIds)}`;
    session.resultChipTextByPlayerId = Object.fromEntries(
      participants.map((player) => {
        const payout = payouts.get(Number(player.id)) || 0;
        return [Number(player.id), makeChipResultText(session.color, payout - session.bet, player[session.color] || 0)];
      })
    );
    session.resultChipTexts = participants.map((player) => (
      `${player.name}: ${session.resultChipTextByPlayerId[Number(player.id)]}`
    ));

    const [log] = await Promise.all([
      addLog(
        'formulahighlow',
        makeCombinationBettingSettlementLog(participants, session, payouts),
        makeCombinationBettingSettlementPublicLog(participants, session, payouts)
      ),
      savePlayers(participants),
      deleteFormulaHighLowSession(session.id)
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({
      ok: true,
      state: getCombinationBettingSessionState(session),
      pot: session.pot,
      payout: Object.fromEntries(payouts),
      winnerIds,
      log: log.text,
      logs: [log],
      players: serializePlayers(participants)
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '콤비네이션 베팅 정산 중 오류가 발생했습니다.' });
  }
}

app.post('/formula-high-low/resolve', handleFormulaHighLowResolve);
app.post('/formula-high-low/settle', handleFormulaHighLowResolve);

app.post('/pvp-baccarat/start', async (req, res) => {
  try {
    const { playerId, opponentId, color, amount } = req.body;
    await loadPlayersByIds([playerId, opponentId]);
    const playerOne = getPlayerById(playerId);
    const playerTwo = getPlayerById(opponentId);
    const bet = toChipAmount(amount);

    if (!playerOne || !playerTwo) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (playerOne.id === playerTwo.id) {
      return res.status(400).json({ error: '서로 다른 플레이어를 선택해야 합니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (playerOne[color] < bet || playerTwo[color] < bet) {
      return res.status(400).json({ error: '두 플레이어 모두 베팅할 칩을 보유해야 합니다.' });
    }

    const shoe = createBaccaratShoe();
    const session = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      gameType: 'pvpbaccarat',
      playerOneId: playerOne.id,
      playerTwoId: playerTwo.id,
      color,
      bet,
      shoe,
      deck: [],
      playerOneCards: [drawBaccaratCard(shoe), drawBaccaratCard(shoe)],
      playerTwoCards: [drawBaccaratCard(shoe), drawBaccaratCard(shoe)],
      playerOneAction: '',
      playerTwoAction: '',
      turn: 'playerOne',
      done: false,
      winnerKey: '',
      result: ''
    };

    const [log] = await Promise.all([
      addLog(
        'pvpbaccarat',
        makePvpLog(playerOne, playerTwo, session, '진행 중'),
        makePvpPublicLog(playerOne, playerTwo, session, '진행 중')
      ),
      savePvpSession(session),
      savePlayers([playerOne, playerTwo])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getPvpSessionState(session), log: log.text, logs: [log], players: serializePlayers([playerOne, playerTwo]) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || 'PVP 바카라 처리 중 오류가 발생했습니다.' });
  }
});

app.post('/pvp-blackjack/start', async (req, res) => {
  try {
    const { playerId, opponentId, color, amount } = req.body;
    await loadPlayersByIds([playerId, opponentId]);
    const playerOne = getPlayerById(playerId);
    const playerTwo = getPlayerById(opponentId);
    const bet = toChipAmount(amount);

    if (!playerOne || !playerTwo) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (playerOne.id === playerTwo.id) {
      return res.status(400).json({ error: '서로 다른 플레이어를 선택해야 합니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (playerOne[color] < bet || playerTwo[color] < bet) {
      return res.status(400).json({ error: '두 플레이어 모두 베팅할 칩을 보유해야 합니다.' });
    }

    const deck = createBlackjackDeck();
    const session = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      gameType: 'pvpblackjack',
      playerOneId: playerOne.id,
      playerTwoId: playerTwo.id,
      color,
      bet,
      deck,
      shoe: [],
      playerOneCards: [drawBlackjackCard(deck), drawBlackjackCard(deck)],
      playerTwoCards: [drawBlackjackCard(deck), drawBlackjackCard(deck)],
      playerOneAction: '',
      playerTwoAction: '',
      turn: 'playerOne',
      done: false,
      winnerKey: '',
      result: ''
    };

    const [log] = await Promise.all([
      addLog(
        'pvpblackjack',
        makePvpLog(playerOne, playerTwo, session, '진행 중'),
        makePvpPublicLog(playerOne, playerTwo, session, '진행 중')
      ),
      savePvpSession(session),
      savePlayers([playerOne, playerTwo])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getPvpSessionState(session), log: log.text, logs: [log], players: serializePlayers([playerOne, playerTwo]) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || 'PVP 블랙잭 처리 중 오류가 발생했습니다.' });
  }
});

app.post('/pvp/action', async (req, res) => {
  try {
    const { sessionId, action } = req.body;
    const session = await getPvpSession(sessionId);

    if (!session) {
      return res.status(400).json({ error: '진행 중인 PVP 게임이 없습니다.' });
    }

    await loadPlayersByIds([session.playerOneId, session.playerTwoId]);
    const playerOne = getPlayerById(session.playerOneId);
    const playerTwo = getPlayerById(session.playerTwoId);

    if (!playerOne || !playerTwo) {
      await deletePvpSession(session.id);
      return res.status(404).json({ error: '플레이어를 찾을 수 없어 세션을 종료했습니다.' });
    }

    const resultText = applyPvpAction(playerOne, playerTwo, session, action);
    const persistence = session.done
      ? deletePvpSession(session.id)
      : savePvpSession(session);
    const savePlayerPromise = session.done
      ? savePlayers([playerOne, playerTwo])
      : Promise.resolve();

    const [log] = await Promise.all([
      addLog(
        session.gameType,
        makePvpLog(playerOne, playerTwo, session, resultText),
        makePvpPublicLog(playerOne, playerTwo, session, resultText)
      ),
      persistence,
      savePlayerPromise
    ]);

    if (session.done) {
      emitRealtime('update', players);
    }
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getPvpSessionState(session), log: log.text, logs: [log], players: serializePlayers([playerOne, playerTwo]) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || 'PVP 진행 중 오류가 발생했습니다.' });
  }
});

app.post('/blackjack/start', async (req, res) => {
  try {
    const { playerId, color, amount } = req.body;
    await loadPlayersByIds([playerId]);
    const player = getPlayerById(playerId);
    const bet = toChipAmount(amount);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (player[color] < bet) {
      return res.status(400).json({ error: '보유 칩이 부족합니다.' });
    }

    assertGameLimitAvailable(player, 'blackjack');
    consumeGameLimit(player, 'blackjack');

    const deck = createBlackjackDeck();
    const playerCards = [drawBlackjackCard(deck), drawBlackjackCard(deck)];
    const dealerCards = [drawBlackjackCard(deck), drawBlackjackCard(deck)];

    const session = {
      playerId: player.id,
      color,
      bet,
      deck,
      playerCards,
      dealerCards,
      logGroupId: createLogGroupId('blackjack', player.id),
      lastDraw: `시작 카드: 플레이어 ${playerCards.join(', ')} / 딜러 ${dealerCards.join(', ')}`
    };

    const [log] = await Promise.all([
      addLog(
        'blackjack',
        makeBlackjackProgressLog(player, session, 'start', '진행 중'),
        makeBlackjackPublicLog(player, session, 'start', '진행 중'),
        { groupId: session.logGroupId }
      ),
      saveBlackjackSession(session),
      savePlayers([player])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '블랙잭 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/blackjack/action', async (req, res) => {
  try {
    const { playerId, action } = req.body;
    const [session] = await Promise.all([
      getBlackjackSession(playerId),
      loadPlayersByIds([playerId])
    ]);
    const player = getPlayerById(playerId);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!session) {
      return res.status(400).json({ error: '진행 중인 블랙잭이 없습니다.' });
    }

    if (action === 'hit') {
      const card = drawBlackjackCard(session.deck);
      session.playerCards.push(card);
      session.lastDraw = `플레이어 hit: ${card}`;

      let resultText = '진행 중';
      let log;

      if (getBlackjackHandValue(session.playerCards) > 21) {
        resultText = finishBlackjackSession(player, session);
        [log] = await Promise.all([
        addLog(
          'blackjack',
          makeBlackjackProgressLog(player, session, 'hit', resultText),
          makeBlackjackPublicLog(player, session, 'hit', resultText),
          { groupId: ensureLogGroupId(session, 'blackjack', player.id) }
        ),
          savePlayers([player]),
          deleteBlackjackSession(playerId)
        ]);
        emitRealtime('update', players);
      } else {
        [log] = await Promise.all([
          addLog(
            'blackjack',
            makeBlackjackProgressLog(player, session, 'hit', resultText),
            makeBlackjackPublicLog(player, session, 'hit', resultText),
            { groupId: ensureLogGroupId(session, 'blackjack', player.id) }
          ),
          saveBlackjackSession(session)
        ]);
      }

      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: [log], players: serializePlayers([player]) });
    }

    if (action === 'stand') {
      session.lastDraw = '플레이어 stand';
      const resultText = finishBlackjackSession(player, session);
      const [log] = await Promise.all([
        addLog(
          'blackjack',
          makeBlackjackProgressLog(player, session, 'stand', resultText),
          makeBlackjackPublicLog(player, session, 'stand', resultText),
          { groupId: ensureLogGroupId(session, 'blackjack', player.id) }
        ),
        savePlayers([player]),
        deleteBlackjackSession(playerId)
      ]);

      emitRealtime('update', players);
      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: [log], players: serializePlayers([player]) });
    }

    return res.status(400).json({ error: '블랙잭 행동이 올바르지 않습니다.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '블랙잭 진행 중 오류가 발생했습니다.' });
  }
});

app.post('/baccarat/start', async (req, res) => {
  try {
    const { playerId, color, amount, side } = req.body;
    await loadPlayersByIds([playerId]);
    const player = getPlayerById(playerId);
    const bet = toChipAmount(amount);
    const nextSide = String(side || '').toUpperCase();

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (!isValidBaccaratSide(nextSide)) {
      return res.status(400).json({ error: '바카라 선택은 PLAYER, BANKER, TIE 중 하나여야 합니다.' });
    }

    if (player[color] < bet) {
      return res.status(400).json({ error: '보유 칩이 부족합니다.' });
    }

    assertGameLimitAvailable(player, 'baccarat');
    consumeGameLimit(player, 'baccarat');

    const shoe = createBaccaratShoe();
    const playerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
    const bankerCards = [drawBaccaratCard(shoe), drawBaccaratCard(shoe)];
    const session = {
      playerId: player.id,
      side: nextSide,
      color,
      bet,
      shoe,
      playerCards,
      bankerCards,
      playerThirdCard: '',
      bankerThirdCard: '',
      playerAction: '',
      done: false,
      outcome: '',
      logGroupId: createLogGroupId('baccarat', player.id)
    };

    let resultText = '플레이어 선택 대기';
    let log;

    if (isBaccaratNatural(session)) {
      session.playerAction = 'natural';
      resultText = finishBaccaratSession(player, session);
      [log] = await Promise.all([
        addLog(
          'baccarat',
          makeBaccaratProgressLog(player, session, 'start', resultText),
          makeBaccaratPublicLog(player, session, 'start', resultText),
          { groupId: session.logGroupId }
        ),
        savePlayers([player]),
        deleteBaccaratSession(player.id)
      ]);
    } else {
      [log] = await Promise.all([
        addLog(
          'baccarat',
          makeBaccaratProgressLog(player, session, 'start', resultText),
          makeBaccaratPublicLog(player, session, 'start', resultText),
          { groupId: session.logGroupId }
        ),
        saveBaccaratSession(session),
        savePlayers([player])
      ]);
    }

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getBaccaratState(session), log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '바카라 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/baccarat/action', async (req, res) => {
  try {
    const { playerId, action } = req.body;
    const [session] = await Promise.all([
      getBaccaratSession(playerId),
      loadPlayersByIds([playerId])
    ]);
    const player = getPlayerById(playerId);
    const nextAction = String(action || '').toLowerCase();

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!session) {
      return res.status(400).json({ error: '진행 중인 바카라가 없습니다.' });
    }

    if (!['hit', 'stand'].includes(nextAction)) {
      return res.status(400).json({ error: '바카라 행동은 hit 또는 stand여야 합니다.' });
    }

    if (nextAction === 'hit') {
      const card = drawBaccaratCard(session.shoe);
      session.playerCards.push(card);
      session.playerThirdCard = card;
      session.playerAction = 'hit';
    } else {
      session.playerAction = 'stand';
    }

    const resultText = finishBaccaratSession(player, session);
    const [log] = await Promise.all([
      addLog(
        'baccarat',
        makeBaccaratProgressLog(player, session, nextAction, resultText),
        makeBaccaratPublicLog(player, session, nextAction, resultText),
        { groupId: ensureLogGroupId(session, 'baccarat', player.id) }
      ),
      savePlayers([player]),
      deleteBaccaratSession(playerId)
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    return res.json({ ok: true, state: getBaccaratState(session), log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '바카라 진행 중 오류가 발생했습니다.' });
  }
});

app.post('/russian-roulette/start', async (req, res) => {
  try {
    const { participantCount, participantIds = [], color, amount } = req.body;
    const expectedParticipantCount = toChipAmount(participantCount);
    const uniqueParticipantIds = [...new Set(participantIds.map(Number))]
      .filter((id) => Number.isFinite(id));
    await loadPlayersByIds(uniqueParticipantIds);
    const participants = uniqueParticipantIds.map(getPlayerById);
    const bet = toChipAmount(amount);

    if (uniqueParticipantIds.length < 2) {
      return res.status(400).json({ error: '러시안룰렛은 2명 이상 참여해야 합니다.' });
    }

    if (
      Number.isFinite(expectedParticipantCount) &&
      expectedParticipantCount >= 2 &&
      uniqueParticipantIds.length !== expectedParticipantCount
    ) {
      return res.status(400).json({ error: `참가자 ${expectedParticipantCount}명을 정확히 선택해야 합니다.` });
    }

    if (participants.some((player) => !player)) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (participants.some((player) => player[color] < bet)) {
      return res.status(400).json({ error: '모든 참가자가 베팅할 칩을 보유해야 합니다.' });
    }

    participants.forEach((player) => {
      player[color] -= bet;
    });

    const session = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      participantIds: uniqueParticipantIds,
      activeIds: [...uniqueParticipantIds],
      eliminated: [],
      color,
      bet,
      pot: bet * uniqueParticipantIds.length,
      round: 1,
      lastAction: `게임 시작. 참가자 ${uniqueParticipantIds.length}명, 라운드마다 생존자 중 1명 탈락.`
    };

    const [log] = await Promise.all([
      addLog(
        'russianroulette',
        makeRussianRouletteLog(session),
        makeRussianRoulettePublicLog(session),
        { groupId: session.id }
      ),
      saveRussianRouletteSession(session),
      savePlayers(participants)
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getRussianRouletteState(session), log: log.text, logs: [log], players: serializePlayers(participants) });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '러시안룰렛 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/russian-roulette/trigger', async (req, res) => {
  try {
    const { sessionId } = req.body;
    const session = await getRussianRouletteSession(sessionId);

    if (!session) {
      return res.status(400).json({ error: '진행 중인 러시안룰렛이 없습니다.' });
    }

    if (session.done || session.activeIds.length <= 1) {
      await deleteRussianRouletteSession(sessionId);
      return res.status(400).json({ error: '이미 종료된 러시안룰렛입니다.' });
    }

    await loadPlayersByIds(session.participantIds);

    const eliminatedIndex = Math.floor(Math.random() * session.activeIds.length);
    const [eliminatedId] = session.activeIds.splice(eliminatedIndex, 1);
    const eliminatedPlayer = getPlayerById(eliminatedId);
    session.eliminated.push({ id: eliminatedId, round: session.round });
    session.lastAction = `${session.round}라운드 탈락: ${eliminatedPlayer?.name || `#${eliminatedId}`}`;

    if (session.activeIds.length === 1) {
      const winner = getPlayerById(session.activeIds[0]);
      if (!winner) {
        await deleteRussianRouletteSession(sessionId);
        return res.status(404).json({ error: '승자를 찾을 수 없어 세션을 종료했습니다.' });
      }

      winner[session.color] += session.pot;
      session.done = true;
      session.winnerId = winner.id;
      session.result = `${winner.name} 승리`;
      session.lastAction += `. 최종 승자: ${winner.name}. 팟 ${session.color} ${session.pot} 지급.`;
      session.resultChipTexts = participants.map((player) => {
        const delta = Number(player.id) === Number(winner.id)
          ? session.pot - session.bet
          : -session.bet;
        return makeNamedChipResultText(player, session.color, delta);
      });

      const [log] = await Promise.all([
        addLog(
          'russianroulette',
          makeRussianRouletteLog(session),
          makeRussianRoulettePublicLog(session),
          { groupId: session.id }
        ),
        savePlayers([winner]),
        deleteRussianRouletteSession(session.id)
      ]);

      emitRealtime('update', players);
      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, fired: true, state: getRussianRouletteState(session), log: log.text, logs: [log], players: serializePlayers(participants) });
    }

    session.round += 1;

    const [log] = await Promise.all([
      addLog(
        'russianroulette',
        makeRussianRouletteLog(session),
        makeRussianRoulettePublicLog(session),
        { groupId: session.id }
      ),
      saveRussianRouletteSession(session)
    ]);

    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, fired: false, state: getRussianRouletteState(session), log: log.text, logs: [log], players: serializePlayers(players) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '러시안룰렛 진행 중 오류가 발생했습니다.' });
  }
});

// 게임 처리
app.post('/game', async (req, res) => {
  try {
    const { gameType, playerId, color, amount, multiplier = 2, extra = {} } = req.body;
    await loadPlayersByIds([playerId]);
    const player = getPlayerById(playerId);
    const bet = toChipAmount(amount);
    const mul = Number(multiplier);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(bet) || bet <= 0) {
      return res.status(400).json({ error: '베팅 수량이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(mul) || mul < 1) {
      return res.status(400).json({ error: '배율은 1 이상 숫자여야 합니다.' });
    }

    if (player[color] < bet) {
      return res.status(400).json({ error: '보유 칩이 부족합니다.' });
    }

    assertGameLimitAvailable(player, gameType);

    let win;
    let logMultiplier = mul;
    let logExtra = extra;
    let result = undefined;
    const beforeBalance = player[color] || 0;

    if (gameType === 'roulette') {
      const roulette = playRoulette(player, color, bet, extra);
      win = roulette.win;
      logMultiplier = `${roulette.payout}:1`;
      logExtra = roulette.extra;
      result = roulette.result;
    } else if (gameType === 'highlow') {
      const highlow = playHighLow(player, color, bet, extra);
      win = highlow.win;
      logMultiplier = highlow.push ? 'push' : '1:1';
      logExtra = highlow.extra;
      result = highlow.extra.card;
    } else if (gameType === 'baccarat') {
      const baccarat = playBaccarat(player, color, bet, extra);
      win = baccarat.win;
      logMultiplier = baccarat.push ? 'push' : `${baccarat.payout}:1`;
      logExtra = baccarat.extra;
      result = baccarat.outcome;
    } else if (gameType === 'redblack') {
      const redblack = playRedBlack(player, color, bet, extra);
      win = redblack.win;
      logMultiplier = '1:1';
      logExtra = redblack.extra;
      result = redblack.result;
    } else {
      win = Math.random() > 0.5;

      // 2배면 +1배, 3배면 +2배, 지면 건 만큼 감소
      if (win) {
        player[color] += toChipAmount(bet * (mul - 1));
      } else {
        player[color] -= bet;
      }
    }

    consumeGameLimit(player, gameType);
    const resultChipDelta = (player[color] || 0) - beforeBalance;
    const resultChipBalance = player[color] || 0;
    const resultChipText = makeChipResultText(color, resultChipDelta, resultChipBalance);

    const [log] = await Promise.all([
      addLog(
        gameType,
        makeGameLog(gameType, player, color, bet, logExtra, logMultiplier, win, resultChipText),
        makeGamePublicLog(gameType, player, color, bet, logExtra, logMultiplier, win, resultChipText)
      ),
      savePlayers([player])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({
      ok: true,
      win,
      log: log.text,
      logs: [log],
      result,
      resultChipDelta,
      resultChipBalance,
      resultChipText,
      players: serializePlayers([player])
    });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '게임 처리 중 오류가 발생했습니다.' });
  }
});

// 칩 이동
app.post('/exchange', async (req, res) => {
  try {
    const { fromId, toId, color, amount } = req.body;
    await loadPlayersByIds([fromId, toId]);
    const from = getPlayerById(fromId);
    const to = getPlayerById(toId);
    const moveAmount = toChipAmount(amount);

    if (!from || !to) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(color)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (!Number.isFinite(moveAmount) || moveAmount <= 0) {
      return res.status(400).json({ error: '이동 수량이 올바르지 않습니다.' });
    }

    if (from.id === to.id) {
      return res.status(400).json({ error: '같은 플레이어끼리는 이동할 수 없습니다.' });
    }

    if (from[color] < moveAmount) {
      return res.status(400).json({ error: '보내는 쪽의 칩이 부족합니다.' });
    }

    from[color] -= moveAmount;
    to[color] += moveAmount;

    const [log] = await Promise.all([
      addLog(
        'exchange',
        makeTransferLog(from, to, color, moveAmount),
        makeTransferPublicLog(from, to, color, moveAmount)
      ),
      savePlayers([from, to])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, log: log.text, logs: [log], players: serializePlayers([from, to]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '칩 이동 중 오류가 발생했습니다.' });
  }
});

// 색 환전
app.post('/convert', async (req, res) => {
  try {
    const { playerId, fromColor, toColor, amount } = req.body;
    await Promise.all([
      loadPlayersByIds([playerId]),
      loadRates()
    ]);
    const player = getPlayerById(playerId);
    const convertAmount = toChipAmount(amount);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    if (!validateColor(fromColor) || !validateColor(toColor)) {
      return res.status(400).json({ error: '칩 색이 올바르지 않습니다.' });
    }

    if (fromColor === toColor) {
      return res.status(400).json({ error: '같은 색끼리는 환전할 수 없습니다.' });
    }

    if (!Number.isFinite(convertAmount) || convertAmount <= 0) {
      return res.status(400).json({ error: '환전 수량이 올바르지 않습니다.' });
    }

    if (player[fromColor] < convertAmount) {
      return res.status(400).json({ error: '보유 칩이 부족합니다.' });
    }

    const totalValue = convertAmount * rates[fromColor];
    const result = Math.floor(totalValue / rates[toColor]);

    if (result <= 0) {
      return res.status(400).json({ error: '환전 결과가 0입니다.' });
    }

    player[fromColor] -= convertAmount;
    player[toColor] += result;

    const [log] = await Promise.all([
      addLog(
        'convert',
        makeConvertLog(player, fromColor, toColor, convertAmount, result),
        makeConvertPublicLog(player, fromColor, toColor, convertAmount, result)
      ),
      savePlayers([player])
    ]);

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, result, log: log.text, logs: [log], players: serializePlayers([player]) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '환전 처리 중 오류가 발생했습니다.' });
  }
});



  if (io) {
    io.on('connection', (socket) => {
      socket.emit('update', players);
      socket.emit('log', lastLogText);
      socket.emit('logs', logHistory);
    });
  }

  return { app, loadSheet, loadRates };
}

const { app } = createApiApp();

module.exports = app;
module.exports.createApiApp = createApiApp;

