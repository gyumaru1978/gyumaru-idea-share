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
 * 5.5. 社長用の合い言葉を設定する（社長コメント機能を使う場合は必須）
 *    同じスクリプト プロパティに、プロパティ名 = PRESIDENT_PASSCODE、
 *    値 = 社長だけが知る合い言葉 を追加する。
 *    アプリの「マスタ」画面でこれを入力すると「社長モード」になり、
 *    各商品に「社長からのコメント・方針」を投稿できる。
 *    投稿のたびにサーバー側でも検証するため、この合い言葉を知らない人は
 *    画面をいじっても社長として投稿することはできない。
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
const CATEGORY_SHEET_NAME = 'カテゴリ';
const TAG_SHEET_NAME      = 'タグ';
const COMMENT_SHEET_NAME  = 'コメント';
const POSTER_SHEET_NAME   = 'ポスター';
const PHOTO_FOLDER_NAME   = '新商品アイデア_写真';
const POSTER_FOLDER_NAME  = '新商品アイデア_ポスター';

// ============================================================
// アイデアシートの列定義
// ============================================================
// 列の位置は必ずこの IDEA_COL 経由で参照する（1始まり＝シートの列番号）。
// getRange(row, 26) のような数字の直書きをすると、列を1本足しただけで
// 写真・更新日時などが静かにズレて壊れるため。
const IDEA_COL = {
  id: 1, name: 2, catch: 3, store: 4, date: 5,
  status: 6, categories: 7, tags: 8,
  yieldQty: 9, yieldUnit: 10, cost: 11, costPer: 12, price: 13, costRate: 14,
  steps: 15, allergens: 16,
  kcal: 17, protein: 18, fat: 19, carb: 20, salt: 21,
  storage: 22, bestBefore: 23, memo: 24,
  photo1: 25, photo2: 26, updatedAt: 27, updatedBy: 28,
  dept: 29,     // 部門（店舗部門／ぎゅう丸ラボ）
  presAt: 30    // 社長コメントの最終投稿日時。一覧の未読判定に使う非正規化値で、
                // 正本はコメントシート。postComment_ が社長投稿のたびに書き換える
};

// 分類の軸は3つ。それぞれ役割が違う。
//   店舗     … どの店舗から出た案か。マスタで管理（店舗間で共有するため必須）
//   カテゴリ … 商品の種類。マスタで管理し、勝手に増やせない（惣菜／デザート／麺 など）
//   タグ     … 誰でも自由に付けられる横断ラベル（SNS映え／春の定番／低カロリー など）。マスタで管理
// ※「採用月」「イベント」列は運用をやめて廃止した（migrateIdeaSheet_ が既存シートから列を削除する）。
const IDEA_HEADERS = [
  'アイデアID', '商品名', 'キャッチコピー', '提案店舗', '提案日',
  'ステータス', 'カテゴリ', 'タグ',
  '出来上がり数', '出来上がり単位', '原価合計', '1食あたり原価', '想定売価', '原価率(%)',
  '作り方', 'アレルギー',
  'エネルギー(kcal)', 'たんぱく質(g)', '脂質(g)', '炭水化物(g)', '食塩相当量(g)',
  '保存方法', '賞味期限目安', 'メモ',
  '写真URL1', '写真URL2', '更新日時', '更新者',
  '部門', '社長コメント日時'
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

const STORE_HEADERS   = ['店舗名', '表示順'];
const STORE_TEXT_COLS = [1];

// ポスターは店舗ごとに保管する。1店舗に何枚でも。画像はDriveに置き、シートにはファイルIDを持つ
// （URLでなくIDにしておくと、表示用サムネイルとダウンロード用リンクの両方を組み立てられる）。
const POSTER_HEADERS   = ['ポスターID', '店舗', 'タイトル', 'ファイルID', '登録日時', '登録者', 'メモ'];
const POSTER_TEXT_COLS = [1, 2, 3, 4, 5, 6, 7];

const CATEGORY_HEADER = 'カテゴリ名';
const TAG_HEADER      = 'タグ名';

// コメントは「社長から」（社長の方針）と「社長へ」（それに対する質問・意見）の
// 2種別を1シートに貯める。行の蓄積がそのまま履歴になる。
// 編集・削除のAPIはあえて作らない（「以前はどういう方針だったか」を後から消せない設計）。
//
// 投稿者は任意入力。社員ごとのアカウント管理はしない方針なので、
// 「誰の質問か分かると社長が返事しやすい」程度のラベルにとどめる。
const COMMENT_HEADERS   = ['アイデアID', '種別', '投稿者', '本文', '投稿日時'];
const COMMENT_TEXT_COLS = [1, 2, 3, 4, 5];
const KIND_FROM_PRES    = '社長から';
const KIND_TO_PRES      = '社長へ';
const COMMENT_KINDS     = [KIND_FROM_PRES, KIND_TO_PRES];

// 部門。ステータス（開発がどの段階か）とは別の軸で、
// 「どちらの事業の商品か」を表す。一覧のタブで切り替える。
const DEPT_OPTIONS = ['店舗部門', 'ぎゅう丸ラボ'];

// ステータスは2択に簡略化した。
//   アイデア … まだ検討段階のもの（旧「提案中／検討中／見送り」はすべてこれに寄せる）
//   採用     … 商品化が決まったもの
// 既存データは migrateIdeaSheet_ が旧ステータスを自動でこの2択へ寄せる。
const STATUS_OPTIONS = ['アイデア', '採用'];
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

// タグシートの初期値。運用が始まったら「マスタ」画面から自由に追加・改名・削除・並び替えできる。
// カテゴリと違い、アイデア入力画面からその場で新しいタグを作ることもでき、
// 作られたタグは自動的にこのマスタへ追加される。
const TAG_SEED = [
  'SNS映え', '春の定番', '夏の定番', '秋の定番', '冬の定番',
  '低カロリー', '高単価', 'テイクアウト向き', '子供向け', '土産向き',
  '地元食材', '時短', '既存設備で作れる', '試作済み'
];

// 店舗の初期値。実際の店舗名は「マスタ」画面から追加・削除できる。
const STORE_SEED    = [['嬉野本店', 1], ['武雄店', 2]];
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
    case 'getInitialData':   return getInitialData_(args[0]);
    case 'getIdeaDetail':    return getIdeaDetail_(args[0]);
    case 'postComment':      return postComment_(args[0], args[1], args[2]);
    case 'verifyPresident':  return verifyPresident_(args[0]);
    case 'addTag':           return addTag_(args[0]);
    case 'renameTag':        return renameTag_(args[0], args[1]);
    case 'deleteTag':        return deleteTag_(args[0]);
    case 'reorderTags':      return reorderTags_(args[0]);
    case 'addIdea':          return addIdea_(args[0]);
    case 'updateIdea':       return updateIdea_(args[0], args[1]);
    case 'updateIdeaStatus': return updateIdeaStatus_(args[0], args[1]);
    case 'deleteIdea':       return deleteIdea_(args[0]);
    case 'addStore':         return addStore_(args[0]);
    case 'updateStore':      return updateStore_(args[0], args[1]);
    case 'deleteStore':      return deleteStore_(args[0]);
    case 'reorderStores':    return reorderStores_(args[0]);
    case 'getPosters':       return getPosters_(args[0]);
    case 'addPoster':        return addPoster_(args[0]);
    case 'deletePoster':     return deletePoster_(args[0]);
    case 'addCategory':      return addCategory_(args[0]);
    case 'deleteCategory':   return deleteCategory_(args[0]);
    case 'reorderCategories':return reorderCategories_(args[0]);
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

// 社長用の合い言葉。社長コメントの投稿時に毎回サーバー側で検証する。
// 画面の「社長モード」はあくまで表示の切り替えで、本人確認はここが本体。
function checkPresident_(pass) {
  const stored = PropertiesService.getScriptProperties().getProperty('PRESIDENT_PASSCODE');
  if (!stored) {
    throw new Error('社長用合い言葉(PRESIDENT_PASSCODE)が未設定です。GASのプロジェクト設定＞スクリプト プロパティで設定してください。');
  }
  if (String(pass) !== String(stored)) {
    throw new Error('社長用の合い言葉が違います');
  }
}

function verifyPresident_(pass) {
  checkPresident_(pass);
  return { ok: true };
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
  } else {
    migrateIdeaSheet_(sh);
  }
  return sh;
}

// 既存シートの列構成を最新に合わせる移行処理。上から順に1回ずつ適用され、
// 適用済みの項目は何もしないため、毎回呼んでも安全。
//   v2: 「提案者」列の廃止（更新者と重複していたため削除）
//   v3: 「部門」「社長コメント日時」列を末尾に追加。
//       既存行の部門は、提案店舗が「ぎゅう丸ラボ」ならぎゅう丸ラボ、それ以外は店舗部門で初期化する
//   v4: ステータスを2択（アイデア／採用）へ寄せる。旧「提案中／検討中／見送り」→「アイデア」
//   v5: 運用をやめた「採用月」「イベント」列を削除する
function migrateIdeaSheet_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  let header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

  // 廃止した列を削除する。右にあるものから順に消すと、削除で左に詰まっても残りの位置がずれない。
  // （提案者=v2、採用月・イベント=v5。イベントの方が右なので先に消す）
  let structural = false;
  ['イベント', '採用月', '提案者'].forEach(name => {
    const idx = header.indexOf(name);
    if (idx >= 0) {
      sh.deleteColumn(idx + 1);
      header.splice(idx, 1);
      structural = true;
    }
  });

  // v3: 部門・社長コメント日時の追加（必ず末尾に足す。途中に挿すと既存データがズレる）
  const needDept = header.indexOf('部門') < 0;
  if (structural || needDept) {
    sh.getRange(1, 1, 1, IDEA_HEADERS.length).setValues([IDEA_HEADERS]);
    const rows = Math.max(sh.getMaxRows() - 1, 1);
    IDEA_TEXT_COLS.forEach(col => sh.getRange(2, col, rows, 1).setNumberFormat('@'));
  }
  if (needDept) {
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const stores = sh.getRange(2, IDEA_COL.store, lastRow - 1, 1).getValues();
      const depts = stores.map(r => {
        if (String(r[0]).trim() === '') return [''];   // 空行には書かない
        return [String(r[0]).trim() === 'ぎゅう丸ラボ' ? 'ぎゅう丸ラボ' : DEPT_OPTIONS[0]];
      });
      sh.getRange(2, IDEA_COL.dept, lastRow - 1, 1).setValues(depts);
    }
  }

  // v4: 旧ステータスを新2択へ正規化（該当する値が残っていなければ何もしない）
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const range = sh.getRange(2, IDEA_COL.status, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;
    values.forEach(r => {
      const s = String(r[0]).trim();
      if (s === '' || s === 'アイデア' || s === '採用') return;
      r[0] = (s === '採用') ? '採用' : 'アイデア';   // 提案中/検討中/見送り など → アイデア
      changed = true;
    });
    if (changed) {
      sh.getRange(2, IDEA_COL.status, lastRow - 1, 1).setNumberFormat('@');
      range.setValues(values);
    }
  }
}

// 使わなくなったシートの後片付け。
// 「既読」シートは、既読を名前ごとにサーバー管理していた頃の名残。
// 既読は端末ローカル(localStorage)方式に変えたので不要になった。
// データが入っている場合は消さない（想定外の使われ方をしていたら残す方が安全）。
function cleanupLegacySheets_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('既読');
  if (sh && sh.getLastRow() <= 1) ss.deleteSheet(sh);
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

function getPosterSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(POSTER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(POSTER_SHEET_NAME);
    initSheet_(sh, POSTER_HEADERS, POSTER_TEXT_COLS, 1000);
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

// タグシートを新規作成するときは、既にアイデアで使われているタグも取り込む。
// タグをマスタ化する前に登録された案のタグが、マスタ一覧から漏れないようにするため。
function getTagSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(TAG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(TAG_SHEET_NAME);
    initSheet_(sh, [TAG_HEADER], [1], 500);
    const seen = {};
    const seed = collectUsedTags_().concat(TAG_SEED).filter(t => {
      if (seen[t]) return false;
      seen[t] = true;
      return true;
    });
    if (seed.length) sh.getRange(2, 1, seed.length, 1).setValues(seed.map(t => [t]));
  }
  return sh;
}

function getCommentSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(COMMENT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(COMMENT_SHEET_NAME);
    initSheet_(sh, COMMENT_HEADERS, COMMENT_TEXT_COLS, 5000);
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
// 既読はサーバーでは持たない。「その端末で読んだか」を端末側(localStorage)に記録する方式にした。
// 社員ごとのアカウントを作らない運用なので、既読のためだけに名前を管理させるのは割に合わない。
// 代償は「同じ人がPCとスマホで見ると片方でまた未読に見える」ことだが、
// 見逃す方向には壊れないので実害は小さい。
function getInitialData_() {
  cleanupLegacySheets_();
  const ideas = getIdeas_();
  return {
    ideas: ideas,
    stores: getStoreList_(),
    categories: getCategoryList_(),
    tags: getTagList_(),
    deptOptions: DEPT_OPTIONS,
    statusOptions: STATUS_OPTIONS,
    unitOptions: UNIT_OPTIONS,
    yieldUnitOptions: YIELD_UNIT_OPTIONS,
    allergenOptions: ALLERGEN_OPTIONS,
    storageOptions: STORAGE_OPTIONS,
    today: todayStr_()
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
    materials: getMaterialsFor_(id),
    comments: getCommentsFor_(id)
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
    store: cell_(row, IDEA_COL.store),
    date: cell_(row, IDEA_COL.date),
    status: cell_(row, IDEA_COL.status) || STATUS_OPTIONS[0],
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
    updatedBy: cell_(row, IDEA_COL.updatedBy),
    // 移行前の既存行など部門が空のものは店舗部門として扱う（絞り込みから漏れないように）
    dept: cell_(row, IDEA_COL.dept) || DEPT_OPTIONS[0],
    presAt: cell_(row, IDEA_COL.presAt)
  };
}

// presAt はフォームからは編集できない値なので、呼び出し側が
// 「新規なら空」「更新なら旧行の値」を明示的に渡す（dataに紛れ込ませない）
function ideaToRow_(id, data, costs, photo1, photo2, updatedAt, presAt) {
  const row = new Array(IDEA_HEADERS.length).fill('');
  const set = (col, v) => { row[col - 1] = v; };

  set(IDEA_COL.id, id);
  set(IDEA_COL.name, String(data.name || '').trim());
  set(IDEA_COL.catch, data.catch || '');
  set(IDEA_COL.store, String(data.store || '').trim());
  set(IDEA_COL.date, data.date || todayStr_());
  set(IDEA_COL.status, data.status || STATUS_OPTIONS[0]);
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
  set(IDEA_COL.updatedBy, String(data.updatedBy || '').trim());
  set(IDEA_COL.dept, DEPT_OPTIONS.indexOf(data.dept) >= 0 ? data.dept : DEPT_OPTIONS[0]);
  set(IDEA_COL.presAt, presAt || '');

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

// アイデアで実際に使われているタグを、使用頻度の高い順に集める
function collectUsedTags_() {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const count = {};
  sh.getRange(2, IDEA_COL.tags, lastRow - 1, 1).getValues().forEach(r => {
    splitList_(r[0]).forEach(t => { count[t] = (count[t] || 0) + 1; });
  });
  return Object.keys(count).sort((a, b) => count[b] - count[a]);
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
    const row = ideaToRow_(id, data, costs, photo1, photo2, updatedAt, '');

    const startRow = sh.getLastRow() + 1;
    applyIdeaTextFormat_(sh, startRow);
    sh.getRange(startRow, 1, 1, row.length).setValues([row]);
    replaceMaterialBlock_(id, materials);
    syncTagsToMaster_(data.tags);

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
    // 工程写真(photo2)は画面から廃止したが、過去に登録した画像は消さずそのまま引き継ぐ
    const photo2 = cell_(oldRow, IDEA_COL.photo2);
    const updatedAt = nowStr_();
    // 提案日はフォームでは変えない値なので旧行から引き継ぐ（編集で今日に書き換わらないように）
    const merged = Object.assign({}, data, {
      date: data.date || cell_(oldRow, IDEA_COL.date)
    });
    const row = ideaToRow_(id, merged, costs, photo1, photo2, updatedAt,
                           cell_(oldRow, IDEA_COL.presAt));

    applyIdeaTextFormat_(sh, rowNum);
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    replaceMaterialBlock_(id, materials);
    syncTagsToMaster_(data.tags);

    return { idea: rowToIdea_(row, rowNum), materials: getMaterialsFor_(id) };
  } finally {
    lock.releaseLock();
  }
}

// 一覧・詳細からのステータス変更専用（アイデア⇄採用のワンタップ切り替え）。
// フォーム全体を送り直さないので写真の再アップロードも材料の再書き込みも起きない。
// patch = { status, updatedBy }
function updateIdeaStatus_(id, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    patch = patch || {};
    if (STATUS_OPTIONS.indexOf(patch.status) < 0) {
      throw new Error('不正なステータスです: ' + patch.status);
    }
    const sh = getIdeaSheet_();
    const rowNum = findIdeaRow_(sh, id);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + id);

    applyIdeaTextFormat_(sh, rowNum);
    sh.getRange(rowNum, IDEA_COL.status, 1, 1).setValues([[patch.status]]);
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
  if (!String(data.store || '').trim())           throw new Error('提案店舗を選択してください');
  if (data.status && STATUS_OPTIONS.indexOf(data.status) < 0) {
    throw new Error('不正なステータスです: ' + data.status);
  }
  if (data.dept && DEPT_OPTIONS.indexOf(data.dept) < 0) {
    throw new Error('不正な部門です: ' + data.dept);
  }
}

// ============================================================
// コメント（社長からのコメント・方針／社長へのコメント）
// ============================================================
function getCommentsFor_(ideaId) {
  const sh = getCommentSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, COMMENT_HEADERS.length).getValues()
    .filter(r => String(r[0]) === String(ideaId))
    .map(r => ({
      kind: r[1] || KIND_TO_PRES,
      author: r[2] || '',
      text: r[3] || '',
      at: r[4] || ''
    }));
    // 追記しかしないシートなので行順＝時系列。日時文字列でソートし直すと
    // 同じ「分」に投稿された複数コメントの前後関係が崩れるため、行順のまま返す
}

// payload = { kind: '社長から'|'社長へ', text, name }
// name は任意。社長からの投稿で未入力なら「社長」とだけ記録する。
// 種別が「社長から」のときは presidentPass をサーバー側で毎回検証する。
// 社長からの投稿はアイデア行の「社長コメント日時」も更新し、これが未読バッジの引き金になる。
function postComment_(ideaId, payload, presidentPass) {
  payload = payload || {};
  const kind = payload.kind;
  const text = String(payload.text || '').trim();
  if (COMMENT_KINDS.indexOf(kind) < 0) throw new Error('不正なコメント種別です: ' + kind);
  if (!text) throw new Error('コメントを入力してください');
  const fromPres = kind === KIND_FROM_PRES;
  if (fromPres) checkPresident_(presidentPass);
  const name = String(payload.name || '').trim() || (fromPres ? '社長' : '');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ideaSh = getIdeaSheet_();
    const rowNum = findIdeaRow_(ideaSh, ideaId);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + ideaId);

    const now = nowStr_();
    const sh = getCommentSheet_();
    const commentRow = sh.getLastRow() + 1;
    COMMENT_TEXT_COLS.forEach(col => sh.getRange(commentRow, col, 1, 1).setNumberFormat('@'));
    sh.getRange(commentRow, 1, 1, COMMENT_HEADERS.length)
      .setValues([[ideaId, kind, name, text, now]]);

    if (fromPres) {
      ideaSh.getRange(rowNum, IDEA_COL.presAt, 1, 1).setNumberFormat('@');
      ideaSh.getRange(rowNum, IDEA_COL.presAt, 1, 1).setValues([[now]]);
    }

    const row = ideaSh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];
    return {
      idea: rowToIdea_(row, rowNum),
      comments: getCommentsFor_(ideaId)
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// タグマスタ
// ============================================================
// マスタに登録された全タグを、シートの行順（＝画面で並び替えた順）で返す。
// マスタに無いのにアイデアで使われているタグ（手作業でシートを編集した等）は
// 末尾に足して返す。そうしないと画面から編集する手段が無くなってしまうため。
function getTagList_() {
  const sh = getTagSheet_();
  const lastRow = sh.getLastRow();
  const list = lastRow < 2 ? [] : sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v !== '');
  const seen = {};
  list.forEach(t => { seen[t] = true; });
  collectUsedTags_().forEach(t => {
    if (!seen[t]) { seen[t] = true; list.push(t); }
  });
  return list;
}

// アイデア登録・更新時に、その場で作られた新しいタグをマスタへ取り込む。
// 既にあるものは黙って無視するので、毎回呼んで問題ない。
function syncTagsToMaster_(tags) {
  splitList_(joinList_(tags)).forEach(t => {
    try { addTagToMaster_(t, false); } catch (e) { /* マスタ追加の失敗で保存を止めない */ }
  });
}

function addTag_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    addTagToMaster_(validateMasterName_(name, 'タグ名'), true);
    return getTagList_();
  } finally {
    lock.releaseLock();
  }
}

// マスタへ1件足す。throwIfDup=false なら既にあっても静かに無視する
// （アイデア登録時にその場で作られたタグを取り込む用）。
function addTagToMaster_(name, throwIfDup) {
  const sh = getTagSheet_();
  const lastRow = sh.getLastRow();
  const existing = lastRow < 2 ? [] : sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim());
  if (existing.indexOf(name) >= 0) {
    if (throwIfDup) throw new Error('同じ名前のタグが既にあります');
    return;
  }
  const rowNum = sh.getLastRow() + 1;
  sh.getRange(rowNum, 1, 1, 1).setNumberFormat('@');
  sh.getRange(rowNum, 1, 1, 1).setValues([[name]]);
}

// 改名はマスタと、そのタグが付いた全アイデアの両方を書き換える。
// 表記ゆれの統一（例：「SNSばえ」→「SNS映え」）にも使える。
function renameTag_(oldName, newName) {
  const from = String(oldName).trim();
  const to = validateMasterName_(newName, 'タグ名');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    renameInMasterColumn_(getTagSheet_(), from, to);
    return { renamed: rewriteTagColumn_(from, to), tags: getTagList_() };
  } finally {
    lock.releaseLock();
  }
}

// タグはカテゴリと違い「使用中でも削除できる」。自由に付け外しするラベルなので、
// 消したいときに全アイデアから外せる方が実用的なため。
function deleteTag_(name) {
  const target = String(name).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    deleteFromMasterColumn_(getTagSheet_(), target);
    return { removed: rewriteTagColumn_(target, null), tags: getTagList_() };
  } finally {
    lock.releaseLock();
  }
}

function reorderTags_(names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    reorderMasterRows_(getTagSheet_(), [TAG_HEADER], [1], names, 0);
    return getTagList_();
  } finally {
    lock.releaseLock();
  }
}

// 1列だけのマスタシート（カテゴリ・タグ）で使う共通処理
function renameInMasterColumn_(sh, oldName, newName) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const range = sh.getRange(2, 1, lastRow - 1, 1);
  const values = range.getValues();
  const idx = values.findIndex(r => String(r[0]).trim() === oldName);
  if (idx < 0) return;
  // 改名先が既にある場合は、行を統合するため旧行を消すだけにする
  const dup = values.findIndex(r => String(r[0]).trim() === newName);
  if (dup >= 0) { sh.deleteRow(idx + 2); return; }
  sh.getRange(idx + 2, 1, 1, 1).setNumberFormat('@');
  sh.getRange(idx + 2, 1, 1, 1).setValues([[newName]]);
}

function deleteFromMasterColumn_(sh, name) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const idx = values.findIndex(r => String(r[0]).trim() === name);
  if (idx >= 0) sh.deleteRow(idx + 2);
}

// タグ列を全行走査し、target を newName に置換（newName=null なら削除）した行数を返す。
// joinList_ を通すので、改名先が既に付いている行では自動的に重複が1つにまとまる。
function rewriteTagColumn_(target, newName) {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const range = sh.getRange(2, IDEA_COL.tags, lastRow - 1, 1);
  const values = range.getValues();
  let count = 0;
  values.forEach(r => {
    const tags = splitList_(r[0]);
    if (tags.indexOf(target) < 0) return;
    const next = newName === null
      ? tags.filter(t => t !== target)
      : tags.map(t => (t === target ? newName : t));
    r[0] = joinList_(next);
    count++;
  });
  if (count) range.setValues(values);
  return count;
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
// ポスター（店舗ごとに保管、ダウンロード可）
// ============================================================
// 店舗を指定して、その店舗のポスター一覧を返す（store 未指定なら全店舗ぶん）。
// 表示用サムネイルとダウンロード用リンクの両方を組み立てて返す。
function getPosters_(store) {
  const sh = getPosterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const target = String(store == null ? '' : store).trim();
  return sh.getRange(2, 1, lastRow - 1, POSTER_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .filter(r => !target || String(r[1]).trim() === target)
    .map(r => {
      const fileId = String(r[3]).trim();
      return {
        id: String(r[0]).trim(),
        store: String(r[1]).trim(),
        title: r[2] || '',
        fileId: fileId,
        at: r[4] || '',
        by: r[5] || '',
        memo: r[6] || '',
        thumbUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000',
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId
      };
    })
    .sort((a, b) => (a.at < b.at ? 1 : -1));   // 新しい順
}

// data = { store, title, image(dataURL), memo, by }
function addPoster_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);   // 画像アップロードを含むので長めに
  try {
    data = data || {};
    const store = String(data.store || '').trim();
    if (!store) throw new Error('店舗を選択してください');
    if (getStoreList_().every(s => s.name !== store)) throw new Error('存在しない店舗です: ' + store);
    const img = String(data.image || '');
    if (img.indexOf('data:') !== 0) throw new Error('ポスター画像を選択してください');

    const fileId = uploadToFolder_(img, getPosterFolder_(), 'poster_');
    const sh = getPosterSheet_();
    const id = generateNextPosterId_(sh);
    const rowNum = sh.getLastRow() + 1;
    POSTER_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, POSTER_HEADERS.length).setValues([[
      id, store, String(data.title || '').trim(), fileId, nowStr_(),
      String(data.by || '').trim(), String(data.memo || '').trim()
    ]]);
    return getPosters_(store);
  } finally {
    lock.releaseLock();
  }
}

function deletePoster_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getPosterSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('ポスターが見つかりません: ' + id);
    const rows = sh.getRange(2, 1, lastRow - 1, POSTER_HEADERS.length).getValues();
    const idx = rows.findIndex(r => String(r[0]).trim() === String(id).trim());
    if (idx < 0) throw new Error('ポスターが見つかりません: ' + id);
    const store = String(rows[idx][1]).trim();
    const fileId = String(rows[idx][3]).trim();
    sh.deleteRow(idx + 2);
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
    return getPosters_(store);
  } finally {
    lock.releaseLock();
  }
}

function generateNextPosterId_(sh) {
  const lastRow = sh.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(/^P(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return 'P' + String(max + 1).padStart(4, '0');
}

function getPosterFolder_() {
  const it = DriveApp.getFoldersByName(POSTER_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(POSTER_FOLDER_NAME);
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

// ============================================================
// マスタの並び替え
// ============================================================
// 画面から渡された名前の並び順どおりにシートの行を並べ替える。
// 「表示順」列を持つマスタ（店舗・イベント）は 1..N で振り直し、
// 持たないマスタ（カテゴリ）は行の物理的な順序がそのまま表示順になる。
//
// 画面に無い名前（別の人が同時に追加した等）は消さずに末尾へ回す。
// ここで消してしまうと、2人が同時にマスタを触ったときに片方の追加が失われる。
function reorderMasterRows_(sh, headers, textCols, orderedNames, orderCol) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .filter(r => String(r[0]).trim() !== '');

  const pos = {};
  (orderedNames || []).forEach((n, i) => { pos[String(n).trim()] = i; });
  rows.forEach((r, i) => { r._i = i; });   // 並びが同点のときに元の順序を保つ
  rows.sort((a, b) => {
    const ia = pos[String(a[0]).trim()], ib = pos[String(b[0]).trim()];
    const va = ia === undefined ? 9999 : ia;
    const vb = ib === undefined ? 9999 : ib;
    return va !== vb ? va - vb : a._i - b._i;
  });
  rows.forEach(r => { delete r._i; });

  if (orderCol) rows.forEach((r, i) => { r[orderCol - 1] = i + 1; });

  sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (rows.length) {
    // 書式は必ず setValues の前に当てる
    textCols.forEach(c => sh.getRange(2, c, rows.length, 1).setNumberFormat('@'));
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function reorderStores_(names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    reorderMasterRows_(getStoreSheet_(), STORE_HEADERS, STORE_TEXT_COLS, names, 2);
    return getStoreList_();
  } finally {
    lock.releaseLock();
  }
}

function reorderCategories_(names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // カテゴリシートは「カテゴリ名」1列だけなので、行の並び順がそのまま表示順になる
    reorderMasterRows_(getCategorySheet_(), [CATEGORY_HEADER], [1], names, 0);
    return getCategoryList_();
  } finally {
    lock.releaseLock();
  }
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
  const fileId = uploadToFolder_(dataUrl, getPhotoFolder_(), 'idea_');
  // uc?export=view形式はGoogle側のウイルススキャン確認画面にリダイレクトされ<img>で表示できないことがあるため、
  // サムネイル配信エンドポイントを使う
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
}

// data:image/...;base64,... をDriveの指定フォルダに保存し、ファイルIDを返す。
// 写真とポスターで共通に使う（呼ぶ側が返り値のIDからURLを組み立てる）。
function uploadToFolder_(dataUrl, folder, prefix) {
  const m = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!m) throw new Error('不正な画像データです');
  const contentType = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  const ext = contentType.split('/')[1] || 'jpg';
  const blob = Utilities.newBlob(bytes, contentType, prefix + new Date().getTime() + '.' + ext);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
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
