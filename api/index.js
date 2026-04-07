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
let logSheetReady = false;
let blackjackSessions = new Map();
let russianRouletteSessions = new Map();
const COLORS = ['red', 'blue', 'green', 'yellow', 'white'];
const CHIP_LABELS = {
  red: 'Red',
  blue: 'Blue',
  green: 'Green',
  yellow: 'Yellow',
  white: 'White'
};

let rates = { red: 1, blue: 3, green: 5, yellow: 10, white: 15 };

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

const PLAYER_RANGE = '플레이어!A:G';
const PLAYER_WRITE_RANGE = '플레이어!A1';
const LOG_SHEET_NAME = '로그';
const LOG_SHEET_REF = `'${LOG_SHEET_NAME}'`;
const LOG_HEADER_RANGE = `${LOG_SHEET_REF}!A1:E1`;
const LOG_RANGE = `${LOG_SHEET_REF}!A:E`;
const LOG_APPEND_RANGE = `${LOG_SHEET_REF}!A:E`;
const LOG_HEADER = ['id', 'type', 'text', 'publicText', 'createdAt'];

// =========================
// 📥 시트 → 서버
// =========================
async function loadSheet() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_RANGE
  });

  const rows = res.data.values || [];

  if (rows.length <= 1) {
    players = [];
    return;
  }

  players = rows.slice(1).map((r) => ({
    id: Number(r[0] || 0),
    name: r[1] || '',
    red: readChipValue(r[2]),
    blue: readChipValue(r[3]),
    green: readChipValue(r[4]),
    yellow: readChipValue(r[5]),
    white: readChipValue(r[6])
  }));
}

// =========================
// 📤 서버 → 시트
// =========================
async function saveSheet() {
  const values = [
    ['id', 'name', ...COLORS],
    ...players.map((p) => [p.id, p.name, ...COLORS.map((color) => p[color] || 0)])
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: PLAYER_WRITE_RANGE,
    valueInputOption: 'RAW',
    requestBody: { values }
  });
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

function parseLogRow(row) {
  return {
    id: Number(row[0]) || Date.parse(row[4]) || 0,
    type: row[1] || '',
    text: row[2] || '',
    publicText: row[3] || row[2] || '',
    createdAt: row[4] || new Date(Number(row[0]) || Date.now()).toISOString()
  };
}

async function loadLogHistory() {
  await ensureLogSheetHeader();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: LOG_RANGE
  });

  const rows = res.data.values || [];
  logHistory = rows
    .slice(1)
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
      values: [[log.id, log.type, log.text, log.publicText, log.createdAt]]
    }
  });
}

function getErrorDetail(err) {
  return err.response?.data?.error?.message || err.message || String(err);
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

function getPlayerById(playerId) {
  return players.find((p) => p.id === Number(playerId));
}

async function addLog(type, text, publicText = text) {
  lastLogText = text;

  const log = {
    id: Date.now(),
    type,
    text,
    publicText,
    createdAt: new Date().toISOString()
  };

  logHistory = [log, ...logHistory];
  await appendLogToSheet(log);
  await loadLogHistory();
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
    result: session.result || ''
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

function makeBlackjackProgressLog(player, session, action, resultText) {
  return `블랙잭 진행
플레이어: ${player.name}
칩: ${session.color} ${session.bet}
행동: ${action}
${blackjackStateText(session)}
결과: ${resultText}
현재: ${chipStr(player)}`;
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
    return `블랙잭
플레이어 패: ${state.playerCards.join(', ')} (${state.playerValue})
딜러 패: ${getPublicDealerCards(session)}`;
  }

  const dealerValue = session.done ? ` (${state.dealerValue})` : '';
  const dealerDraws = session.done && state.dealerDraws.length
    ? `\n딜러 추가 카드: ${state.dealerDraws.join(', ')}`
    : '';

  return `블랙잭
플레이어: ${player.name}
베팅: ${session.color} ${session.bet}
행동: ${action}
플레이어 패: ${state.playerCards.join(', ')} (${state.playerValue})
딜러 패: ${getPublicDealerCards(session)}${dealerValue}
이번 결과: ${state.lastDraw || '-'}${dealerDraws}
결과: ${resultText}${session.done ? `\n플레이어 잔고: ${chipStr(player)}` : ''}`;
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
  } else if (dealerValue > 21 || playerValue > dealerValue) {
    player[session.color] += session.bet;
    session.result = '승리';
  } else if (playerValue < dealerValue) {
    player[session.color] -= session.bet;
    session.result = '패배';
  } else {
    session.result = '무승부';
  }

  session.done = true;
  blackjackSessions.delete(String(session.playerId));
  return session.result;
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
    lastAction: session.lastAction || ''
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
결과: ${state.lastAction}${state.done ? `\n최종 승자: ${state.winnerName}` : ''}`;
}

function makeRussianRoulettePublicLog(session) {
  const state = getRussianRouletteState(session);
  const active = state.activePlayers.map((p) => p.name).join(', ') || '-';
  const eliminated = state.eliminatedPlayers.map((p) => `${p.round}R ${p.name}`).join(', ') || '-';

  return `러시안룰렛
생존자: ${active}
탈락자: ${eliminated}
라운드: ${state.round}
결과: ${state.lastAction}${state.done ? `\n최종 승자: ${state.winnerName}` : ''}`;
}

// =========================
// 🎮 게임별 로그
// =========================
function logRoulette(p, color, bet, pick, detail, result, multiplier, win) {
  return `🎡 룰렛
${p.name} | 칩: ${color} ${bet}
베팅 종류: ${pick}
세부: ${detail || '-'}
결과: ${result}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
}

function logHighLow(p, color, bet, choice, card, multiplier, win) {
  return `📈 하이로우
${p.name} | 칩: ${color} ${bet}
선택: ${choice}
오픈 카드: ${card || '-'}
배율: ${multiplier}배
판정: ${multiplier === 'push' ? '무승부' : win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
}

function logBaccarat(p, color, bet, side, result, multiplier, turn, win) {
  return `🃏 바카라
${p.name} | 칩: ${color} ${bet}
턴: ${turn}
선택: ${side}
결과: ${result}
배율: ${multiplier}배
판정: ${multiplier === 'push' ? '무승부' : win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
}

function logBlackjack(p, color, bet, action, playerSum, dealerSum, multiplier, turn, win) {
  return `♠ 블랙잭
${p.name} | 칩: ${color} ${bet}
턴: ${turn}
행동: ${action}
플레이어: ${playerSum}
딜러: ${dealerSum}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
}

function logRedBlack(p, color, bet, pick, result, multiplier, win) {
  return `🟥⬛ 레드 앤 블랙
${p.name} | 칩: ${color} ${bet}
선택: ${pick}
결과: ${result}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
}

function makeGameLog(gameType, p, color, amount, extra, multiplier, win) {
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
        win
      );
    case 'highlow':
      return logHighLow(
        p,
        color,
        amount,
        extra.choice,
        extra.card,
        multiplier,
        win
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
        win
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
        win
      );
    case 'redblack':
      return logRedBlack(
        p,
        color,
        amount,
        extra.pick,
        extra.result,
        multiplier,
        win
      );
    default:
      return `🎮 게임 결과
${p.name} | 칩: ${color} ${amount}
배율: ${multiplier}배
판정: ${win ? '승리' : '패배'}
현재: ${chipStr(p)}`;
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

function makeConvertLog(p, fromColor, toColor, amount, result) {
  return `💱 환전
플레이어: ${p.name}
변환: ${fromColor} ${amount} → ${toColor} ${result}
기준 비율: ${COLORS.map((color) => `${CHIP_LABELS[color]}=${rates[color]}`).join(', ')}
현재: ${chipStr(p)}`;
}

function makeBalanceLog(p) {
  return `현재 잔고
ID: ${p.id}
플레이어: ${p.name}
잔고: ${chipStr(p)}`;
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
    await loadSheet();
    res.json(players);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '플레이어 데이터를 불러오지 못했습니다.' });
  }
});

// 마지막 로그 조회
app.get('/logtext', async (req, res) => {
  try {
    const logs = await loadLogHistory();
    res.send(logs[0]?.text || lastLogText);
  } catch (err) {
    console.error(err);
    res.status(500).send('로그를 불러오지 못했습니다.');
  }
});

app.get('/logs', async (req, res) => {
  try {
    const logs = await loadLogHistory();
    res.json(logs);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '로그를 불러오지 못했습니다.', detail: getErrorDetail(err) });
  }
});

app.post('/logs/test', async (req, res) => {
  try {
    const log = await addLog('test', `로그 저장 테스트\ncreatedAt: ${new Date().toISOString()}`);
    res.json({ ok: true, log, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '로그 저장 테스트 실패', detail: getErrorDetail(err) });
  }
});

app.post('/players/:playerId/balance-log', async (req, res) => {
  try {
    await loadSheet();

    const player = getPlayerById(req.params.playerId);

    if (!player) {
      return res.status(404).json({ error: '플레이어를 찾을 수 없습니다.' });
    }

    const log = await addLog('balance', makeBalanceLog(player));

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, log: log.text, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '잔고 로그를 생성하지 못했습니다.', detail: getErrorDetail(err) });
  }
});

// 환율 조회
app.get('/rates', (req, res) => {
  res.json(rates);
});

// 환율 설정
app.post('/rates', (req, res) => {
  const nextRates = Object.fromEntries(
    COLORS.map((color) => [color, Number(req.body[color])])
  );

  if (COLORS.some((color) => !Number.isFinite(nextRates[color]) || nextRates[color] <= 0)) {
    return res.status(400).json({ error: '비율은 1 이상의 숫자여야 합니다.' });
  }

  rates = nextRates;

  res.json({ ok: true, rates });
});

app.post('/blackjack/start', async (req, res) => {
  try {
    await loadSheet();

    const { playerId, color, amount } = req.body;
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
      lastDraw: `시작 카드: 플레이어 ${playerCards.join(', ')} / 딜러 ${dealerCards.join(', ')}`
    };

    blackjackSessions.set(String(player.id), session);

    const log = await addLog(
      'blackjack',
      makeBlackjackProgressLog(player, session, 'start', '진행 중'),
      makeBlackjackPublicLog(player, session, 'start', '진행 중')
    );

    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '블랙잭 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/blackjack/action', async (req, res) => {
  try {
    await loadSheet();

    const { playerId, action } = req.body;
    const player = getPlayerById(playerId);
    const session = blackjackSessions.get(String(playerId));

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
      if (getBlackjackHandValue(session.playerCards) > 21) {
        resultText = finishBlackjackSession(player, session);
        await saveSheet();
        emitRealtime('update', players);
      }

      const log = await addLog(
        'blackjack',
        makeBlackjackProgressLog(player, session, 'hit', resultText),
        makeBlackjackPublicLog(player, session, 'hit', resultText)
      );
      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: logHistory });
    }

    if (action === 'stand') {
      session.lastDraw = '플레이어 stand';
      const resultText = finishBlackjackSession(player, session);
      await saveSheet();

      const log = await addLog(
        'blackjack',
        makeBlackjackProgressLog(player, session, 'stand', resultText),
        makeBlackjackPublicLog(player, session, 'stand', resultText)
      );

      emitRealtime('update', players);
      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, state: getBlackjackState(session), log: log.text, logs: logHistory });
    }

    return res.status(400).json({ error: '블랙잭 행동이 올바르지 않습니다.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '블랙잭 진행 중 오류가 발생했습니다.' });
  }
});

app.post('/russian-roulette/start', async (req, res) => {
  try {
    await loadSheet();

    const { participantCount, participantIds = [], color, amount } = req.body;
    const expectedParticipantCount = toChipAmount(participantCount);
    const uniqueParticipantIds = [...new Set(participantIds.map(Number))]
      .filter((id) => Number.isFinite(id));
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

    russianRouletteSessions.set(session.id, session);
    await saveSheet();

    const log = await addLog(
      'russianroulette',
      makeRussianRouletteLog(session),
      makeRussianRoulettePublicLog(session)
    );

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, state: getRussianRouletteState(session), log: log.text, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '러시안룰렛 시작 중 오류가 발생했습니다.' });
  }
});

app.post('/russian-roulette/trigger', async (req, res) => {
  try {
    await loadSheet();

    const { sessionId } = req.body;
    const session = russianRouletteSessions.get(String(sessionId));

    if (!session) {
      return res.status(400).json({ error: '진행 중인 러시안룰렛이 없습니다.' });
    }

    if (session.done || session.activeIds.length <= 1) {
      russianRouletteSessions.delete(String(sessionId));
      return res.status(400).json({ error: '이미 종료된 러시안룰렛입니다.' });
    }

    const eliminatedIndex = Math.floor(Math.random() * session.activeIds.length);
    const [eliminatedId] = session.activeIds.splice(eliminatedIndex, 1);
    const eliminatedPlayer = getPlayerById(eliminatedId);
    session.eliminated.push({ id: eliminatedId, round: session.round });
    session.lastAction = `${session.round}라운드 탈락: ${eliminatedPlayer?.name || `#${eliminatedId}`}`;

    if (session.activeIds.length === 1) {
      const winner = getPlayerById(session.activeIds[0]);
      if (!winner) {
        russianRouletteSessions.delete(String(sessionId));
        return res.status(404).json({ error: '승자를 찾을 수 없어 세션을 종료했습니다.' });
      }

      winner[session.color] += session.pot;
      session.done = true;
      session.winnerId = winner.id;
      session.result = `${winner.name} 승리`;
      session.lastAction += `. 최종 승자: ${winner.name}. 팟 ${session.color} ${session.pot} 지급.`;

      await saveSheet();
      russianRouletteSessions.delete(String(sessionId));

      const log = await addLog(
        'russianroulette',
        makeRussianRouletteLog(session),
        makeRussianRoulettePublicLog(session)
      );

      emitRealtime('update', players);
      emitRealtime('log', lastLogText);
      emitRealtime('logs', logHistory);

      return res.json({ ok: true, fired: true, state: getRussianRouletteState(session), log: log.text, logs: logHistory });
    }

    session.round += 1;

    const log = await addLog(
      'russianroulette',
      makeRussianRouletteLog(session),
      makeRussianRoulettePublicLog(session)
    );

    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, fired: false, state: getRussianRouletteState(session), log: log.text, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '러시안룰렛 진행 중 오류가 발생했습니다.' });
  }
});

// 게임 처리
app.post('/game', async (req, res) => {
  try {
    await loadSheet();

    const { gameType, playerId, color, amount, multiplier = 2, extra = {} } = req.body;
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

    let win;
    let logMultiplier = mul;
    let logExtra = extra;
    let result = undefined;

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

    const log = await addLog(gameType, makeGameLog(gameType, player, color, bet, logExtra, logMultiplier, win));

    await saveSheet();

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, win, log: log.text, logs: logHistory, result });
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ error: err.message || '게임 처리 중 오류가 발생했습니다.' });
  }
});

// 칩 이동
app.post('/exchange', async (req, res) => {
  try {
    await loadSheet();

    const { fromId, toId, color, amount } = req.body;
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

    const log = await addLog('exchange', makeTransferLog(from, to, color, moveAmount));

    await saveSheet();

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, log: log.text, logs: logHistory });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: '칩 이동 중 오류가 발생했습니다.' });
  }
});

// 색 환전
app.post('/convert', async (req, res) => {
  try {
    await loadSheet();

    const { playerId, fromColor, toColor, amount } = req.body;
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

    const log = await addLog('convert', makeConvertLog(player, fromColor, toColor, convertAmount, result));

    await saveSheet();

    emitRealtime('update', players);
    emitRealtime('log', lastLogText);
    emitRealtime('logs', logHistory);

    res.json({ ok: true, result, log: log.text, logs: logHistory });
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

  return { app, loadSheet };
}

const { app } = createApiApp();

module.exports = app;
module.exports.createApiApp = createApiApp;

