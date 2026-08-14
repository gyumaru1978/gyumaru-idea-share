/**
 * 新商品アイデア共有アプリ バックエンド（Google Apps Script）
 *
 * 新商品のアイデアを店舗をまたいで社内で出し合い、レシピ（材料・分量・作り方・原価）まで
 * 一緒に貯めておくためのアプリ。
 * 「何月に採用した案」「今度のイベントの採用案」を切り替えて見られることが中心機能。
 *
 * ▼デプロイ手順
 * 1. Googleドライブで新しいスプレッドシートを作る（名前の例：新商品アイデア共有DB）
 *    URL の https://docs.google.com/spreadsheets/d/【ここがID】/edit の部分をコピーする
 * 2. 下の SPREADSHEET_ID に 1. のIDを貼り付ける
 *    ※シート（アイデア／材料／店舗／イベント／カテゴリ）は初回アクセス時に自動生成されるので、
 *      手動でシートを作る必要はない
 * 3. https://script.google.com/ で「新しいプロジェクト」を作成する
 *    ※固定資産管理アプリの Apps Script プロジェクトを使い回さないこと。
 *      使い回すと固定資産管理アプリが上書きされて壊れる
 * 4. このファイルの内容を「Code.gs」に、index.html の内容を「index.html」に貼り付ける
 *    （ファイル > 新規作成 > HTML、ファイル名は必ず index にする）
 *    ※clasp を使う場合は、このフォルダで `clasp push` するだけでよい
 * 5. 合い言葉（パスコード）を設定する
 *    「プロジェクトの設定（歯車アイコン）> スクリプト プロパティ」を開き、
 *    プロパティ名 = APP_PASSCODE、値 = 任意の合い言葉 を1つ追加して保存する。
 *    これを設定しないとアプリは全操作を拒否する（＝ロックされたまま）。
 *    ※合い言葉は clasp からは設定できない。必ずブラウザ上で設定する
 * 6. 「デプロイ > 新しいデプロイ」→ 種類の歯車アイコンで「ウェブアプリ」を選択
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員（社内のみで使うなら「Googleアカウントを持つ全員」でも可）
 *        ※アプリ起動時に APP_PASSCODE（合い言葉）の入力を求めるため、
 *          URLを知っていても合い言葉を知らない人は閲覧・登録・削除できない。
 *          合い言葉は社内だけで共有し、外部に漏らさないこと。
 * 7. 発行された URL(.../exec) が、そのままスマホ・PCからアクセスするアプリのURLになる
 *
 * ▼修正したときの再デプロイ（重要な落とし穴）
 * - ブラウザで編集した場合：
 *   「デプロイを管理」→ 編集（鉛筆アイコン）→ バージョンを「新バージョン」にして「デプロイ」。
 *   これをやらないと URL 上の内容が更新されない。
 * - clasp を使う場合（clasp 3.x で動作確認済み）：
 *   `clasp push` の後に `clasp deploy` を素で叩くと【毎回新しいURLが発行され】、
 *   社員がブックマークした旧URLは古いコードのまま動き続ける。同じURLを保つには
 *   既存のデプロイIDを指定して「更新」する：
 *       clasp push
 *       clasp deployments                        … 既存のデプロイIDを確認する
 *       clasp redeploy <既存ID> -d "v2 ○○を修正"  … 同じURLのまま中身だけ差し替わる
 *   （`clasp redeploy` は `clasp update-deployment` の別名。
 *     `clasp deploy -i <既存ID>` でも同じことができる）
 *
 * ▼写真について
 * 写真は Google ドライブ内の専用フォルダにアップロードし、そのURLだけをシートに保存する
 * （シート1セルの文字数上限（5万字）を base64 画像が超えてしまうため）。
 */

// ▼▼▼ ここにスプレッドシートIDを設定する ▼▼▼
// 「新商品アイデア共有」スプレッドシート
// https://docs.google.com/spreadsheets/d/14aEoP1eeF4LrClVLxgpiUiqQDEfA2cWde5aRYFnnFnI/edit
const SPREADSHEET_ID = '14aEoP1eeF4LrClVLxgpiUiqQDEfA2cWde5aRYFnnFnI';

const IDEA_SHEET_NAME     = 'アイデア';
const MATERIAL_SHEET_NAME = '材料';
const STORE_SHEET_NAME    = '店舗';
const EVENT_SHEET_NAME    = 'イベント';
const CATEGORY_SHEET_NAME = 'カテゴリ';
const PHOTO_FOLDER_NAME   = '新商品アイデア_写真';

// ============================================================
// アイデアシートの列定義
// ============================================================
// 列の位置は必ずこの IDEA_COL 経由で参照する（1始まり＝シートの列番号）。
// getRange(row, 26) のような数字の直書きをすると、列を1本足しただけで
// 写真・更新日時などが静かにズレて壊れるため。
const IDEA_COL = {
  id: 1, name: 2, catch: 3, author: 4, store: 5, date: 6,
  status: 7, month: 8, event: 9, categories: 10, tags: 11,
  yieldQty: 12, yieldUnit: 13, cost: 14, costPer: 15, price: 16, costRate: 17,
  steps: 18, allergens: 19,
  kcal: 20, protein: 21, fat: 22, carb: 23, salt: 24,
  storage: 25, bestBefore: 26, memo: 27,
  photo1: 28, photo2: 29, updatedAt: 30, updatedBy: 31
};

// 分類の軸は4つ。それぞれ役割が違う。
//   店舗     … どの店舗から出た案か。マスタで管理（店舗間で共有するため必須）
//   カテゴリ … 商品の種類。マスタで管理し、勝手に増やせない（惣菜／デザート／麺 など）
//   タグ     … 誰でも自由に付けられる横断ラベル（SNS映え／春の定番／低カロリー など）。マスタ無し
//   イベント … 採用先の催事。マスタで管理
const IDEA_HEADERS = [
  'アイデアID', '商品名', 'キャッチコピー', '提案者', '提案店舗', '提案日',
  'ステータス', '採用月', 'イベント', 'カテゴリ', 'タグ',
  '出来上がり数', '出来上がり単位', '原価合計', '1食あたり原価', '想定売価', '原価率(%)',
  '作り方', 'アレルギー',
  'エネルギー(kcal)', 'たんぱく質(g)', '脂質(g)', '炭水化物(g)', '食塩相当量(g)',
  '保存方法', '賞味期限目安', 'メモ',
  '写真URL1', '写真URL2', '更新日時', '更新者'
];

// 数式インジェクション対策として書式を強制的に「テキスト」にする列。
// 数値として集計・グラフ化したい列（出来上がり数・原価・売価・原価率・栄養）だけを除外する。
// 提案日・採用月もテキスト書式にして、シート側で日付型に自動変換されない
// （'2026/09' が '2026/09/01' の日付に化けてフィルタが全滅しない）ようにする。
const IDEA_NUM_COLS = [
  IDEA_COL.yieldQty, IDEA_COL.cost, IDEA_COL.costPer, IDEA_COL.price, IDEA_COL.costRate,
  IDEA_COL.kcal, IDEA_COL.protein, IDEA_COL.fat, IDEA_COL.carb, IDEA_COL.salt
];
const IDEA_TEXT_COLS = IDEA_HEADERS
  .map((_, i) => i + 1)
  .filter(col => IDEA_NUM_COLS.indexOf(col) < 0);

const MATERIAL_HEADERS   = ['アイデアID', '表示順', '材料名', '数量', '単位', '単価', '金額', '備考'];
const MATERIAL_TEXT_COLS = [1, 3, 5, 8];

const EVENT_HEADERS   = ['イベント名', '開催時期', '表示順', '備考'];
const EVENT_TEXT_COLS = [1, 2, 4];

const STORE_HEADERS   = ['店舗名', '表示順'];
const STORE_TEXT_COLS = [1];

const CATEGORY_HEADER = 'カテゴリ名';

const STATUS_OPTIONS = ['提案中', '検討中', '採用', '見送り'];
const UNIT_OPTIONS   = ['g', 'kg', 'ml', 'L', '個', '枚', '本', '袋', '缶', '大さじ', '小さじ', '適量'];
const YIELD_UNIT_OPTIONS = ['食', '個', '本', '枚', 'パック', 'kg'];
// 特定原材料8品目＋準ずるもの20品目（消費者庁の表示区分に対応）
const ALLERGEN_OPTIONS = [
  'えび', 'かに', 'くるみ', '小麦', 'そば', '卵', '乳', '落花生',
  'アーモンド', 'あわび', 'いか', 'いくら', 'オレンジ', 'カシューナッツ', 'キウイフルーツ',
  '牛肉', 'ごま', 'さけ', 'さば', '大豆', '鶏肉', 'バナナ', '豚肉', 'まつたけ',
  'もも', 'やまいも', 'りんご', 'ゼラチン'
];
const STORAGE_OPTIONS = ['常温', '冷蔵', '冷凍'];

// タグは自由入力だが、まだ1件も登録が無いと何を入れてよいか分からないので候補を用意しておく。
// 実際に使われたタグは自動で候補に追加されていく（collectTags_）。
const TAG_SUGGESTIONS = [
  'SNS映え', '春の定番', '夏の定番', '秋の定番', '冬の定番',
  '低カロリー', '高単価', 'テイクアウト向き', '子供向け', '土産向き',
  '地元食材', '時短', '既存設備で作れる', '試作済み'
];

// 店舗の初期値。実際の店舗名は「マスタ」画面から追加・削除できる。
const STORE_SEED    = [['嬉野本店', 1], ['武雄店', 2]];
const EVENT_SEED = [
  ['秋の物産展', '10月', 1, ''],
  ['春の物産展', '4月', 2, ''],
  ['夏のギフト', '7月', 3, ''],
  ['定番化検討', '通年', 9, 'イベントに紐づかない定番候補']
];
const CATEGORY_SEED = ['惣菜', 'デザート', '弁当・丼', 'パン', '麺', '季節限定', '新食材', 'リニューアル'];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('新商品アイデア共有')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 公開APIの入口（合い言葉チェック → 各処理へ振り分け）
// ============================================================
// フロントは google.script.run.apiCall(...) だけを呼ぶ。
// 実処理の関数は末尾「_」で非公開にしてあり、外部から直接は呼べない
// （google.script.run はアンダースコア終わりの関数を実行できない）ため、
// 合い言葉チェックを迂回して addIdea_ 等を直接叩かれる穴が塞がれる。
function apiCall(passcode, action, args) {
  checkPasscode_(passcode);
  args = args || [];
  switch (action) {
    case 'getInitialData':   return getInitialData_();
    case 'getIdeaDetail':    return getIdeaDetail_(args[0]);
    case 'addIdea':          return addIdea_(args[0]);
    case 'updateIdea':       return updateIdea_(args[0], args[1]);
    case 'updateIdeaStatus': return updateIdeaStatus_(args[0], args[1]);
    case 'deleteIdea':       return deleteIdea_(args[0]);
    case 'addStore':         return addStore_(args[0]);
    case 'updateStore':      return updateStore_(args[0], args[1]);
    case 'deleteStore':      return deleteStore_(args[0]);
    case 'addEvent':         return addEvent_(args[0]);
    case 'updateEvent':      return updateEvent_(args[0], args[1]);
    case 'deleteEvent':      return deleteEvent_(args[0]);
    case 'addCategory':      return addCategory_(args[0]);
    case 'deleteCategory':   return deleteCategory_(args[0]);
    default: throw new Error('不正な操作です: ' + action);
  }
}

// 合い言葉はコードに直書きせず、GASの「プロジェクトの設定 ＞ スクリプト プロパティ」に
// キー APP_PASSCODE で保存する。未設定なら全操作を拒否する。
function checkPasscode_(passcode) {
  const stored = PropertiesService.getScriptProperties().getProperty('APP_PASSCODE');
  if (!stored) {
    throw new Error('合い言葉(APP_PASSCODE)が未設定です。GASのプロジェクト設定＞スクリプト プロパティで設定してください。');
  }
  if (String(passcode) !== String(stored)) {
    throw new Error('合い言葉が違います');
  }
}

// ============================================================
// シート取得（無ければヘッダー付きで新規作成）
// ============================================================
function getSpreadsheet_() {
  if (!SPREADSHEET_ID) {
    throw new Error('スプレッドシートIDが未設定です。Code.gs の SPREADSHEET_ID を設定してください。');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 新規シート作成時の共通処理。ヘッダーを書き、1行目を固定し、テキスト列の書式を先に当てる。
function initSheet_(sh, headers, textCols, preformatRows) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  textCols.forEach(col => sh.getRange(2, col, preformatRows || 1000, 1).setNumberFormat('@'));
}

function getIdeaSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(IDEA_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(IDEA_SHEET_NAME);
    initSheet_(sh, IDEA_HEADERS, IDEA_TEXT_COLS, 1000);
  }
  return sh;
}

function getMaterialSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(MATERIAL_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MATERIAL_SHEET_NAME);
    initSheet_(sh, MATERIAL_HEADERS, MATERIAL_TEXT_COLS, 5000);
  }
  return sh;
}

function getStoreSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(STORE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(STORE_SHEET_NAME);
    initSheet_(sh, STORE_HEADERS, STORE_TEXT_COLS, 500);
    sh.getRange(2, 1, STORE_SEED.length, STORE_HEADERS.length).setValues(STORE_SEED);
  }
  return sh;
}

function getEventSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(EVENT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(EVENT_SHEET_NAME);
    initSheet_(sh, EVENT_HEADERS, EVENT_TEXT_COLS, 500);
    sh.getRange(2, 1, EVENT_SEED.length, EVENT_HEADERS.length).setValues(EVENT_SEED);
  }
  return sh;
}

function getCategorySheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CATEGORY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CATEGORY_SHEET_NAME);
    initSheet_(sh, [CATEGORY_HEADER], [1], 500);
    sh.getRange(2, 1, CATEGORY_SEED.length, 1).setValues(CATEGORY_SEED.map(c => [c]));
  }
  return sh;
}

// ============================================================
// 初期データ取得
// ============================================================
// 一覧画面に必要なものを一括で返す。材料は含めない（詳細を開いたときに getIdeaDetail で取る）。
// 読み取り系では LockService を取らない。朝に全社員が同時に一覧を開いたとき、
// 読み取りがロック待ちで直列化すると体感が壊れるため。
//
// ※アイデアが数百件を超えて初回表示が重くなってきたら、ここを一覧用の軽い列だけに絞り、
//   詳細は getIdeaDetail に寄せる（写真URLと原価だけあれば一覧は描ける）。
function getInitialData_() {
  const ideas = getIdeas_();
  return {
    ideas: ideas,
    stores: getStoreList_(),
    events: getEvents_(),
    categories: getCategoryList_(),
    tags: collectTags_(ideas),
    statusOptions: STATUS_OPTIONS,
    unitOptions: UNIT_OPTIONS,
    yieldUnitOptions: YIELD_UNIT_OPTIONS,
    allergenOptions: ALLERGEN_OPTIONS,
    storageOptions: STORAGE_OPTIONS,
    monthOptions: buildMonthOptions_(ideas),
    today: todayStr_(),
    thisMonth: thisMonthStr_()
  };
}

function getIdeas_() {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  const ideas = [];
  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, IDEA_HEADERS.length).getValues();
    values.forEach((row, i) => {
      if (row.every(c => c === '' || c === null || c === undefined)) return; // 空行はスキップ
      ideas.push(rowToIdea_(row, i + 2));
    });
  }
  return ideas;
}

function getIdeaDetail_(id) {
  const sh = getIdeaSheet_();
  const rowNum = findIdeaRow_(sh, id);
  if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + id);
  const row = sh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];
  return {
    idea: rowToIdea_(row, rowNum),
    materials: getMaterialsFor_(id)
  };
}

// A列（アイデアID）だけを読んで行番号を返す。
// 全31列を読んでから探すのは無駄が大きいのでこの形にしている。
function findIdeaRow_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const idx = ids.findIndex(r => String(r[0]) === String(id));
  return idx < 0 ? -1 : idx + 2;
}

// ============================================================
// 行 <-> オブジェクト変換
// ============================================================
// row は0始まりの配列、IDEA_COL は1始まりの列番号なので必ず -1 して引く
function cell_(row, col) {
  const v = row[col - 1];
  return v === null || v === undefined ? '' : v;
}

function rowToIdea_(row, rowNum) {
  return {
    _row: rowNum,
    id: cell_(row, IDEA_COL.id),
    name: cell_(row, IDEA_COL.name),
    catch: cell_(row, IDEA_COL.catch),
    author: cell_(row, IDEA_COL.author),
    store: cell_(row, IDEA_COL.store),
    date: cell_(row, IDEA_COL.date),
    status: cell_(row, IDEA_COL.status) || STATUS_OPTIONS[0],
    month: cell_(row, IDEA_COL.month),          // 採用月 'yyyy/MM'
    event: cell_(row, IDEA_COL.event),
    categories: splitList_(cell_(row, IDEA_COL.categories)),
    tags: splitList_(cell_(row, IDEA_COL.tags)),
    yieldQty: numOrZero_(cell_(row, IDEA_COL.yieldQty)),
    yieldUnit: cell_(row, IDEA_COL.yieldUnit),
    cost: numOrZero_(cell_(row, IDEA_COL.cost)),
    costPer: numOrZero_(cell_(row, IDEA_COL.costPer)),
    price: numOrZero_(cell_(row, IDEA_COL.price)),
    costRate: numOrZero_(cell_(row, IDEA_COL.costRate)),
    steps: cell_(row, IDEA_COL.steps),
    allergens: splitList_(cell_(row, IDEA_COL.allergens)),
    kcal: numOrBlank_(cell_(row, IDEA_COL.kcal)),
    protein: numOrBlank_(cell_(row, IDEA_COL.protein)),
    fat: numOrBlank_(cell_(row, IDEA_COL.fat)),
    carb: numOrBlank_(cell_(row, IDEA_COL.carb)),
    salt: numOrBlank_(cell_(row, IDEA_COL.salt)),
    storage: cell_(row, IDEA_COL.storage),
    bestBefore: cell_(row, IDEA_COL.bestBefore),
    memo: cell_(row, IDEA_COL.memo),
    photo1: cell_(row, IDEA_COL.photo1),
    photo2: cell_(row, IDEA_COL.photo2),
    updatedAt: cell_(row, IDEA_COL.updatedAt),
    updatedBy: cell_(row, IDEA_COL.updatedBy)
  };
}

function ideaToRow_(id, data, costs, photo1, photo2, updatedAt) {
  const row = new Array(IDEA_HEADERS.length).fill('');
  const set = (col, v) => { row[col - 1] = v; };

  set(IDEA_COL.id, id);
  set(IDEA_COL.name, String(data.name || '').trim());
  set(IDEA_COL.catch, data.catch || '');
  set(IDEA_COL.author, String(data.author || '').trim());
  set(IDEA_COL.store, String(data.store || '').trim());
  set(IDEA_COL.date, data.date || todayStr_());
  set(IDEA_COL.status, data.status || STATUS_OPTIONS[0]);
  set(IDEA_COL.month, data.month || '');
  set(IDEA_COL.event, data.event || '');
  set(IDEA_COL.categories, joinList_(data.categories));
  set(IDEA_COL.tags, joinList_(data.tags));
  set(IDEA_COL.yieldQty, numOrZero_(data.yieldQty));
  set(IDEA_COL.yieldUnit, data.yieldUnit || '');
  set(IDEA_COL.cost, costs.cost);
  set(IDEA_COL.costPer, costs.costPer);
  set(IDEA_COL.price, numOrZero_(data.price));
  set(IDEA_COL.costRate, costs.costRate);
  set(IDEA_COL.steps, data.steps || '');
  set(IDEA_COL.allergens, joinList_(data.allergens));
  // 栄養値は未入力なら空セルのままにする（0 と「測っていない」は意味が違うため）
  set(IDEA_COL.kcal, numOrBlank_(data.kcal));
  set(IDEA_COL.protein, numOrBlank_(data.protein));
  set(IDEA_COL.fat, numOrBlank_(data.fat));
  set(IDEA_COL.carb, numOrBlank_(data.carb));
  set(IDEA_COL.salt, numOrBlank_(data.salt));
  set(IDEA_COL.storage, data.storage || '');
  set(IDEA_COL.bestBefore, data.bestBefore || '');
  set(IDEA_COL.memo, data.memo || '');
  set(IDEA_COL.photo1, photo1 || '');
  set(IDEA_COL.photo2, photo2 || '');
  set(IDEA_COL.updatedAt, updatedAt);
  set(IDEA_COL.updatedBy, String(data.updatedBy || data.author || '').trim());

  return row;
}

function splitList_(v) {
  return String(v == null ? '' : v)
    .split(',')
    .map(s => s.trim())
    .filter(s => s !== '');
}

// カンマ区切りで1セルに詰めるので、要素側のカンマ・改行は取り除いておく
// （タグは自由入力なので、ここを素通しにすると分解が壊れる）。重複も落とす。
function joinList_(arr) {
  if (!arr) return '';
  const list = Array.isArray(arr) ? arr : String(arr).split(',');
  const seen = {};
  return list
    .map(s => String(s).replace(/[,\r\n]/g, ' ').trim())
    .filter(s => {
      if (s === '' || seen[s]) return false;
      seen[s] = true;
      return true;
    })
    .join(',');
}

// 既に使われているタグを集めて、入力候補として返す（使用頻度の高い順）
function collectTags_(ideas) {
  const count = {};
  (ideas || []).forEach(idea => {
    (idea.tags || []).forEach(t => { count[t] = (count[t] || 0) + 1; });
  });
  const used = Object.keys(count).sort((a, b) => count[b] - count[a]);
  const seen = {};
  return used.concat(TAG_SUGGESTIONS).filter(t => {
    if (seen[t]) return false;
    seen[t] = true;
    return true;
  });
}

function numOrZero_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

// 栄養値は「未入力」と「0」を区別したいので、空欄はそのまま空文字で返す
function numOrBlank_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isFinite(n) ? n : '';
}

function applyIdeaTextFormat_(sh, rowNum) {
  IDEA_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
}

function thisMonthStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM');
}

// 採用月の選択肢。実際に使われている月と、今月の前後（-3ヶ月〜+9ヶ月）を混ぜて降順で返す。
// 未来月を先に出しておくことで「2026年11月の枠に今から案を入れておく」運用ができる。
function buildMonthOptions_(ideas) {
  const set = {};
  (ideas || []).forEach(idea => {
    if (idea.month) set[idea.month] = true;
  });
  const base = new Date();
  for (let i = -3; i <= 9; i++) {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    set[Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM')] = true;
  }
  return Object.keys(set).sort().reverse();
}

// ============================================================
// 原価計算
// ============================================================
// ★この計算式は index.html の calcCost() と完全に同一。
//   片方だけ直すと画面表示とシート保存値がズレるので、必ず両方セットで直すこと。
//   （GASにはビルド工程が無く共通モジュールを作れないため、二重実装は避けられない）
//
// 丸めは「行ごとに丸めてから合計」に固定する。丸めずに合計してから丸めると
// クライアントの表示値とサーバーの保存値が1円ずれ、「画面は30.1%なのにシートは30.0%」になる。
function calcCost_(materials, yieldQty, price) {
  let cost = 0;
  (materials || []).forEach(m => {
    cost += Math.round(numOrZero_(m.qty) * numOrZero_(m.unitPrice));
  });
  const y = numOrZero_(yieldQty);
  const costPer = y > 0 ? Math.round(cost / y) : 0;
  const p = numOrZero_(price);
  // 0除算で Infinity / NaN がシートに書かれると以後の集計が全部壊れるので明示的に0を返す
  const costRate = (p > 0 && costPer > 0) ? Math.round(costPer / p * 1000) / 10 : 0;
  return { cost: cost, costPer: costPer, costRate: costRate };
}

// ============================================================
// 材料（子シート）
// ============================================================
function getMaterialsFor_(ideaId) {
  const sh = getMaterialSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, MATERIAL_HEADERS.length).getValues();
  return values
    .filter(r => String(r[0]) === String(ideaId))
    .sort((a, b) => numOrZero_(a[1]) - numOrZero_(b[1]))
    .map(r => ({
      name: r[2] || '',
      qty: numOrZero_(r[3]),
      unit: r[4] || '',
      unitPrice: numOrZero_(r[5]),
      amount: numOrZero_(r[6]),
      memo: r[7] || ''
    }));
}

// 対象アイデアの既存材料行を消してから、新しい材料をブロックで末尾に追記する。
// 差分更新より単純で崩れず、削除は「連続する行のまとまり」単位でまとめて行うため
// 材料が数千行に増えても API 呼び出し回数が行数に比例しない。
function replaceMaterialBlock_(ideaId, materials) {
  const sh = getMaterialSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const targetRows = [];
    ids.forEach((r, i) => {
      if (String(r[0]) === String(ideaId)) targetRows.push(i + 2);
    });
    // 連続する行をまとめて、下から削除する（上から消すと行番号がずれる）
    for (let i = targetRows.length - 1; i >= 0; i--) {
      const end = targetRows[i];
      let start = end;
      while (i > 0 && targetRows[i - 1] === start - 1) {
        i--;
        start = targetRows[i];
      }
      sh.deleteRows(start, end - start + 1);
    }
  }

  const rows = (materials || [])
    .filter(m => String(m.name || '').trim() !== '')
    .map((m, i) => [
      ideaId,
      i + 1,
      String(m.name).trim(),
      numOrZero_(m.qty),
      m.unit || '',
      numOrZero_(m.unitPrice),
      Math.round(numOrZero_(m.qty) * numOrZero_(m.unitPrice)),
      m.memo || ''
    ]);
  if (rows.length) {
    const startRow = sh.getLastRow() + 1;
    // 書式は必ず setValues の前に当てる。逆順だと1回目の書き込みが数式として評価される
    MATERIAL_TEXT_COLS.forEach(col => sh.getRange(startRow, col, rows.length, 1).setNumberFormat('@'));
    sh.getRange(startRow, 1, rows.length, MATERIAL_HEADERS.length).setValues(rows);
  }
}

function deleteMaterialsFor_(ideaId) {
  replaceMaterialBlock_(ideaId, []);
}

// ============================================================
// アイデア 新規登録
// ============================================================
function addIdea_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    validateIdeaInput_(data);
    const sh = getIdeaSheet_();
    const id = generateNextIdeaId_(sh);
    const materials = data.materials || [];
    // クライアントから送られてきた原価は一切信用せず、材料行から必ず再計算する
    const costs = calcCost_(materials, data.yieldQty, data.price);
    const photo1 = resolvePhoto_(data.photo1, '');
    const photo2 = resolvePhoto_(data.photo2, '');
    const updatedAt = nowStr_();
    const row = ideaToRow_(id, data, costs, photo1, photo2, updatedAt);

    const startRow = sh.getLastRow() + 1;
    applyIdeaTextFormat_(sh, startRow);
    sh.getRange(startRow, 1, 1, row.length).setValues([row]);
    replaceMaterialBlock_(id, materials);

    return { idea: rowToIdea_(row, startRow), materials: getMaterialsFor_(id) };
  } finally {
    lock.releaseLock();
  }
}

// 既存のアイデアID（A0001 形式）の最大値+1を採番する
function generateNextIdeaId_(sh) {
  const lastRow = sh.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const m = String(r[0]).match(/^A(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return 'A' + String(max + 1).padStart(4, '0');
}

// ============================================================
// アイデア 更新
// ============================================================
function updateIdea_(id, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    validateIdeaInput_(data);
    const sh = getIdeaSheet_();
    const rowNum = findIdeaRow_(sh, id);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + id);
    const oldRow = sh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];

    // 楽観ロック。画面を開いている間に他の人が保存していたら上書きせずに弾く。
    // （data.updatedAt を送ってこない呼び出しではチェックしない）
    const oldUpdatedAt = cell_(oldRow, IDEA_COL.updatedAt);
    if (data.updatedAt && String(oldUpdatedAt) !== String(data.updatedAt)) {
      throw new Error('他の人が先に更新しました。画面を再読み込みしてから編集してください（最終更新: '
        + oldUpdatedAt + ' / ' + (cell_(oldRow, IDEA_COL.updatedBy) || '不明') + '）');
    }

    const materials = data.materials || [];
    const costs = calcCost_(materials, data.yieldQty, data.price);
    const photo1 = resolvePhoto_(data.photo1, cell_(oldRow, IDEA_COL.photo1));
    const photo2 = resolvePhoto_(data.photo2, cell_(oldRow, IDEA_COL.photo2));
    const updatedAt = nowStr_();
    // 提案日は登録時のものを引き継ぐ（編集で今日に書き換わってしまわないように）
    const merged = Object.assign({}, data, { date: data.date || cell_(oldRow, IDEA_COL.date) });
    const row = ideaToRow_(id, merged, costs, photo1, photo2, updatedAt);

    applyIdeaTextFormat_(sh, rowNum);
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    replaceMaterialBlock_(id, materials);

    return { idea: rowToIdea_(row, rowNum), materials: getMaterialsFor_(id) };
  } finally {
    lock.releaseLock();
  }
}

// 一覧・詳細からのステータス変更専用。フォーム全体を送り直さないので
// 写真の再アップロードも材料の再書き込みも起きない（会議中に連続で「採用」を押す場面用）。
// patch = { status, month, event, updatedBy }
function updateIdeaStatus_(id, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    patch = patch || {};
    if (STATUS_OPTIONS.indexOf(patch.status) < 0) {
      throw new Error('不正なステータスです: ' + patch.status);
    }
    if (patch.month && !/^\d{4}\/\d{2}$/.test(String(patch.month))) {
      throw new Error('採用月の形式が不正です: ' + patch.month);
    }
    const sh = getIdeaSheet_();
    const rowNum = findIdeaRow_(sh, id);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + id);

    applyIdeaTextFormat_(sh, rowNum);
    // ステータス・採用月・イベントは隣り合っているのでまとめて書ける
    sh.getRange(rowNum, IDEA_COL.status, 1, 3).setValues([[
      patch.status,
      patch.month || '',
      patch.event || ''
    ]]);
    sh.getRange(rowNum, IDEA_COL.updatedAt, 1, 2)
      .setValues([[nowStr_(), String(patch.updatedBy || '').trim()]]);

    const row = sh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];
    return { idea: rowToIdea_(row, rowNum) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// アイデア 削除
// ============================================================
function deleteIdea_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getIdeaSheet_();
    const rowNum = findIdeaRow_(sh, id);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + id);
    const photoUrls = sh.getRange(rowNum, IDEA_COL.photo1, 1, 2).getValues()[0];
    sh.deleteRow(rowNum);
    deleteMaterialsFor_(id);
    photoUrls.forEach(url => deletePhotoByUrl_(url));
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

function validateIdeaInput_(data) {
  if (!data || !String(data.name || '').trim())   throw new Error('商品名を入力してください');
  if (!String(data.author || '').trim())          throw new Error('提案者を入力してください');
  if (!String(data.store || '').trim())           throw new Error('提案店舗を選択してください');
  if (data.status && STATUS_OPTIONS.indexOf(data.status) < 0) {
    throw new Error('不正なステータスです: ' + data.status);
  }
  if (data.month && !/^\d{4}\/\d{2}$/.test(String(data.month))) {
    throw new Error('採用月の形式が不正です: ' + data.month);
  }
}

// ============================================================
// 店舗マスタ
// ============================================================
function getStoreList_() {
  const sh = getStoreSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, STORE_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({ name: String(r[0]).trim(), order: numOrZero_(r[1]) }))
    .sort((a, b) => (a.order - b.order) || (a.name < b.name ? -1 : 1));
}

function addStore_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const name = validateMasterName_(data && data.name, '店舗名');
    if (getStoreList_().some(s => s.name === name)) {
      throw new Error('同じ名前の店舗が既にあります');
    }
    const sh = getStoreSheet_();
    const rowNum = sh.getLastRow() + 1;
    STORE_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, STORE_HEADERS.length).setValues([[name, numOrZero_(data.order)]]);
    return getStoreList_();
  } finally {
    lock.releaseLock();
  }
}

// 店舗名がそのままキーなので、改名したらアイデア側の「提案店舗」列も一括で書き換える
function updateStore_(oldName, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const newName = validateMasterName_(data && data.name, '店舗名');
    const sh = getStoreSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('店舗が見つかりません: ' + oldName);
    const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const idx = names.findIndex(r => String(r[0]).trim() === String(oldName).trim());
    if (idx < 0) throw new Error('店舗が見つかりません: ' + oldName);
    if (newName !== oldName && names.some(r => String(r[0]).trim() === newName)) {
      throw new Error('同じ名前の店舗が既にあります');
    }
    const rowNum = idx + 2;
    STORE_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, STORE_HEADERS.length).setValues([[newName, numOrZero_(data.order)]]);

    let renamed = 0;
    if (newName !== oldName) renamed = cascadeRename_(IDEA_COL.store, oldName, newName);
    return { stores: getStoreList_(), renamed: renamed };
  } finally {
    lock.releaseLock();
  }
}

function deleteStore_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const inUse = countIdeasByColumn_(IDEA_COL.store, name);
    if (inUse > 0) {
      throw new Error('「' + name + '」は' + inUse + '件のアイデアで使われているため削除できません。'
        + '先にそのアイデアの提案店舗を変更してください');
    }
    const sh = getStoreSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      const idx = names.findIndex(r => String(r[0]).trim() === String(name).trim());
      if (idx >= 0) sh.deleteRow(idx + 2);
    }
    return getStoreList_();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// イベント管理
// ============================================================
function getEvents_() {
  const sh = getEventSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, EVENT_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      name: String(r[0]).trim(),
      period: r[1] || '',
      order: numOrZero_(r[2]),
      memo: r[3] || ''
    }))
    .sort((a, b) => (a.order - b.order) || (a.name < b.name ? -1 : 1));
}

function addEvent_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const name = validateMasterName_(data && data.name, 'イベント名');
    if (getEvents_().some(e => e.name === name)) {
      throw new Error('同じ名前のイベントが既にあります');
    }
    const sh = getEventSheet_();
    const rowNum = sh.getLastRow() + 1;
    EVENT_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, EVENT_HEADERS.length).setValues([[
      name, data.period || '', numOrZero_(data.order), data.memo || ''
    ]]);
    return getEvents_();
  } finally {
    lock.releaseLock();
  }
}

// イベント名がそのままキーなので、改名したらアイデア側の「イベント」列も一括で書き換える
function updateEvent_(oldName, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const newName = validateMasterName_(data && data.name, 'イベント名');
    const sh = getEventSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('イベントが見つかりません: ' + oldName);
    const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const idx = names.findIndex(r => String(r[0]).trim() === String(oldName).trim());
    if (idx < 0) throw new Error('イベントが見つかりません: ' + oldName);
    if (newName !== oldName && names.some(r => String(r[0]).trim() === newName)) {
      throw new Error('同じ名前のイベントが既にあります');
    }
    const rowNum = idx + 2;
    EVENT_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, EVENT_HEADERS.length).setValues([[
      newName, data.period || '', numOrZero_(data.order), data.memo || ''
    ]]);

    let renamed = 0;
    if (newName !== oldName) renamed = cascadeRename_(IDEA_COL.event, oldName, newName);
    return { events: getEvents_(), renamed: renamed };
  } finally {
    lock.releaseLock();
  }
}

function deleteEvent_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const inUse = countIdeasByColumn_(IDEA_COL.event, name);
    if (inUse > 0) {
      throw new Error('「' + name + '」は' + inUse + '件のアイデアで使われているため削除できません。'
        + '先にそのアイデアのイベントを変更してください');
    }
    const sh = getEventSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      const idx = names.findIndex(r => String(r[0]).trim() === String(name).trim());
      if (idx >= 0) sh.deleteRow(idx + 2);
    }
    return getEvents_();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// カテゴリ管理
// ============================================================
function getCategoryList_() {
  const sh = getCategorySheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v !== '');
}

function addCategory_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const trimmed = validateMasterName_(name, 'カテゴリ名');
    if (getCategoryList_().indexOf(trimmed) >= 0) throw new Error('同じ名前のカテゴリが既にあります');
    const sh = getCategorySheet_();
    const rowNum = sh.getLastRow() + 1;
    sh.getRange(rowNum, 1, 1, 1).setNumberFormat('@');
    sh.getRange(rowNum, 1, 1, 1).setValues([[trimmed]]);
    return getCategoryList_();
  } finally {
    lock.releaseLock();
  }
}

function deleteCategory_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // カテゴリは1セルにカンマ区切りで複数入るので、完全一致ではなく要素として含むかで判定する
    const inUse = countIdeasByListColumn_(IDEA_COL.categories, name);
    if (inUse > 0) {
      throw new Error('「' + name + '」は' + inUse + '件のアイデアで使われているため削除できません。'
        + '先にそのアイデアのカテゴリを変更してください');
    }
    const sh = getCategorySheet_();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      const idx = names.findIndex(r => String(r[0]).trim() === String(name).trim());
      if (idx >= 0) sh.deleteRow(idx + 2);
    }
    return getCategoryList_();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// マスタ共通のヘルパー
// ============================================================
// マスタ名の共通バリデーション。カンマ・改行を含む名前を許すと
// カテゴリ・タグのカンマ区切り分解が壊れるので弾く。
function validateMasterName_(name, label) {
  const trimmed = String(name == null ? '' : name).trim();
  if (!trimmed) throw new Error(label + 'を入力してください');
  if (/[,\r\n]/.test(trimmed)) throw new Error(label + 'にカンマや改行は使えません');
  return trimmed;
}

// マスタを名前でキーにしているため、改名時はアイデア側の該当列を一括で書き換える必要がある
function cascadeRename_(col, oldName, newName) {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const range = sh.getRange(2, col, lastRow - 1, 1);
  const values = range.getValues();
  let count = 0;
  values.forEach(r => {
    if (String(r[0]).trim() === String(oldName).trim()) {
      r[0] = newName;
      count++;
    }
  });
  if (count) range.setValues(values);
  return count;
}

function countIdeasByColumn_(col, value) {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  return sh.getRange(2, col, lastRow - 1, 1).getValues()
    .filter(r => String(r[0]).trim() === String(value).trim()).length;
}

function countIdeasByListColumn_(col, value) {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const target = String(value).trim();
  return sh.getRange(2, col, lastRow - 1, 1).getValues()
    .filter(r => splitList_(r[0]).indexOf(target) >= 0).length;
}

// ============================================================
// 写真アップロード（Googleドライブ）
// ============================================================
// data:image/jpeg;base64,... 形式の新規アップロード時のみアップロードし、
// 既存のURL（変更なし）やクリア（空文字）はそのまま透過する
function resolvePhoto_(newValue, oldUrl) {
  const v = newValue === undefined || newValue === null ? '' : String(newValue);
  if (v.indexOf('data:') === 0) {
    const url = uploadPhoto_(v);
    if (oldUrl) deletePhotoByUrl_(oldUrl);
    return url;
  }
  if (!v) {
    if (oldUrl) deletePhotoByUrl_(oldUrl);
    return '';
  }
  return v;
}

function uploadPhoto_(dataUrl) {
  const m = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!m) throw new Error('不正な画像データです');
  const contentType = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  const ext = contentType.split('/')[1] || 'jpg';
  const blob = Utilities.newBlob(bytes, contentType, 'idea_' + new Date().getTime() + '.' + ext);
  const folder = getPhotoFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // uc?export=view形式はGoogle側のウイルススキャン確認画面にリダイレクトされ<img>で表示できないことがあるため、
  // サムネイル配信エンドポイントを使う
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

function getPhotoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function deletePhotoByUrl_(url) {
  const id = extractFileId_(url);
  if (!id) return;
  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (e) {
    // 既に削除済み・アクセス不可の場合は無視
  }
}

function extractFileId_(url) {
  if (!url) return null;
  const m = String(url).match(/[?&]id=([^&]+)/);
  return m ? m[1] : null;
}
