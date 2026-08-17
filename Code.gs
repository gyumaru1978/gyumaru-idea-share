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
const PLAN_SHEET_NAME     = '開発計画';
const DEVTASK_SHEET_NAME  = '開発タスク';
const OWNER_SHEET_NAME    = '担当';
const PROG_SHEET_NAME     = '開発進捗';
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
  presAt: 30,   // 社長コメントの最終投稿日時。postComment_ が社長投稿のたびに書き換える
  seenAt: 31    // 「見ました」確認の最終日時（全員共有・端末非依存）。
                // presAt より古ければ New（新しい社長コメントあり）。
                // 誰かが「見ました」チェック or 社長へ返信すると now に更新され、全員のNewが消える
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
  '部門', '社長コメント日時', '確認日時'
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

// 材料マスタ。単価・栄養は材料ごとに「基準量＋単位」を自由に決めて登録する。
// 例：鶏もも肉なら「1000g で 900円・1900kcal…」、玉ねぎなら「1個 で 30円・37kcal…」。
// レシピではこのマスタから材料を選び、量を入れるだけで、基準量との比で原価・栄養・アレルギーを計算する。
const MMASTER_SHEET_NAME = '材料マスタ';
const MMASTER_HEADERS = [
  '材料名', '表示順', '基準量', '単位', '単価(円)',
  'エネルギー(kcal)', 'たんぱく質(g)', '脂質(g)', '炭水化物(g)', '食塩相当量(g)',
  'アレルギー', '仕入れ先', 'メモ'
];
// 数値列（表示順・基準量・単価・栄養5項目）以外をテキスト書式にする
const MMASTER_NUM_COLS  = [2, 3, 5, 6, 7, 8, 9, 10];
const MMASTER_TEXT_COLS = MMASTER_HEADERS.map((_, i) => i + 1).filter(c => MMASTER_NUM_COLS.indexOf(c) < 0);
const MMASTER_SEED = [
  // [名前, 表示順, 基準量, 単位, 単価, kcal, たんぱく質, 脂質, 炭水化物, 食塩, アレルギー, 仕入れ先, メモ]
  ['鶏もも肉', 1, 1000, 'g', 900, 1900, 166, 142, 0, 2, '鶏肉', '', ''],
  ['牛豚合いびき肉', 2, 1000, 'g', 1300, 2500, 175, 200, 3, 3, '牛肉,豚肉', '', ''],
  ['玉ねぎ', 3, 1, '個', 30, 74, 2, 0.2, 17.6, 0, '', '', '約200g/個'],
  ['牛乳', 4, 1000, 'ml', 220, 610, 33, 38, 48, 1, '乳', '', ''],
  ['卵', 5, 1, '個', 25, 76, 6.1, 5.1, 0.2, 0.2, '卵', '', '約50g/個']
];

// ポスターは商品（アイデア）ごとに保管し、1枚ずつ「どの店舗で使ったか」のラベルを付ける。
// 商品の編集画面で追加し、詳細画面で一覧・ダウンロードする。
// 画像はDriveに置き、シートにはファイルIDを持つ
// （URLでなくIDにしておくと、表示用サムネイルとダウンロード用リンクの両方を組み立てられる）。
const POSTER_HEADERS   = ['ポスターID', 'アイデアID', '店舗', 'ファイルID', '登録日時', '登録者'];
const POSTER_TEXT_COLS = [1, 2, 3, 4, 5, 6];

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
// 表示は「採用→アイデア」の順。ただし新規・空欄の既定は「アイデア」なので別に持つ
const STATUS_OPTIONS = ['採用', 'アイデア'];
const STATUS_DEFAULT = 'アイデア';
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

// ============================================================
// 開発計画（季節メニューのタスク管理）
// ============================================================
// 「年度×季節」で1つの計画（例：2026年度の春）。年度は5月始まり（2026年度＝2026/5〜2027/4）。
// 計画に提供期間・開発期間を持たせ、タスク（工程）をぶら下げる。
// 進捗はタスクの状態（完了=1・進行中=0.5）から自動計算し、別途は保存しない。
const PLAN_HEADERS   = ['計画ID', '年度', '季節', '提供開始', '提供終了', '開発開始', '開発終了', 'メモ', '更新日時'];
const PLAN_TEXT_COLS = [1, 2, 3, 4, 5, 6, 7, 8, 9];   // 日付はすべて yyyy/MM/dd の文字列
const PLAN_COL = { id: 1, year: 2, season: 3, serveStart: 4, serveEnd: 5, devStart: 6, devEnd: 7, memo: 8, updatedAt: 9 };

// タスク（工程）は全グループ共通：名前・期間・並び順を1本だけ持つ。
// 担当/状態/結果の3列は旧方式（グループごとにタスクをコピー）の名残で、v2移行後は使わない（空になる）。
const DEVTASK_HEADERS   = ['タスクID', '計画ID', '表示順', 'タスク名', '開始日', '終了日', '担当', '状態', '結果', '更新日時'];
const DEVTASK_TEXT_COLS = [1, 2, 4, 5, 6, 7, 8, 9, 10];   // 表示順(3)だけ数値
const DEVTASK_COL = { id: 1, planId: 2, order: 3, name: 4, start: 5, end: 6, owners: 7, status: 8, result: 9, updatedAt: 10 };

// グループごとの進捗（状態・結果）。タスク×グループで1行。
const PROG_HEADERS   = ['タスクID', 'グループ', '状態', '結果', '更新日時'];
const PROG_TEXT_COLS = [1, 2, 3, 4, 5];
const PROG_COL = { taskId: 1, group: 2, status: 3, result: 4, updatedAt: 5 };

const SEASONS = ['春', '夏', '秋', '冬'];
const TASK_STATUS_OPTIONS = ['未着手', '進行中', '完了'];
const TASK_STATUS_DEFAULT = '未着手';

// タスクの担当マスタ（＝開発グループ）。店舗マスタとは別物。
// ロードマップはこのグループごとのセクションに分かれる。顔ぶれはよく変わる想定なので、
// マスタ画面から追加・改名・削除・並び替え自由（改名はタスクへ一括反映）。
const OWNER_HEADER = '担当名';
const OWNER_SEED   = ['社長', '佐賀・長崎エリア', '福岡エリア'];

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
    case 'deleteComment':    return deleteComment_(args[0], args[1], args[2]);
    case 'markSeen':         return markSeen_(args[0], args[1]);
    case 'verifyPresident':  return verifyPresident_(args[0]);
    case 'addTag':           return addTag_(args[0]);
    case 'renameTag':        return renameTag_(args[0], args[1]);
    case 'deleteTag':        return deleteTag_(args[0]);
    case 'reorderTags':      return reorderTags_(args[0]);
    case 'saveMaterial':     return saveMaterial_(args[0], args[1]);
    case 'deleteMaterial':   return deleteMaterial_(args[0]);
    case 'reorderMaterials': return reorderMaterialsMaster_(args[0]);
    case 'addIdea':          return addIdea_(args[0]);
    case 'updateIdea':       return updateIdea_(args[0], args[1]);
    case 'updateIdeaStatus': return updateIdeaStatus_(args[0], args[1]);
    case 'deleteIdea':       return deleteIdea_(args[0]);
    case 'addStore':         return addStore_(args[0]);
    case 'updateStore':      return updateStore_(args[0], args[1]);
    case 'deleteStore':      return deleteStore_(args[0]);
    case 'reorderStores':    return reorderStores_(args[0]);
    case 'addPoster':        return addPoster_(args[0]);
    case 'deletePoster':     return deletePoster_(args[0]);
    case 'addCategory':      return addCategory_(args[0]);
    case 'renameCategory':   return renameCategory_(args[0], args[1]);
    case 'deleteCategory':   return deleteCategory_(args[0]);
    case 'reorderCategories':return reorderCategories_(args[0]);
    case 'savePlan':         return savePlan_(args[0]);
    case 'addDevTask':       return addDevTask_(args[0]);
    case 'addDevTasksBulk':  return addDevTasksBulk_(args[0]);
    case 'updateDevTask':    return updateDevTask_(args[0], args[1]);
    case 'deleteDevTask':    return deleteDevTask_(args[0]);
    case 'setTaskProgress':  return setTaskProgress_(args[0], args[1], args[2]);
    case 'reorderDevTasks':  return reorderDevTasks_(args[0], args[1]);
    case 'addOwner':         return addOwner_(args[0]);
    case 'renameOwner':      return renameOwner_(args[0], args[1]);
    case 'deleteOwner':      return deleteOwner_(args[0]);
    case 'reorderOwners':    return reorderOwners_(args[0]);
    default: throw new Error('不正な操作です: ' + action);
  }
}

// 合い言葉はコードに直書きせず、GASの「プロジェクトの設定 ＞ スクリプト プロパティ」に
// キー APP_PASSCODE で保存する。未設定なら全操作を拒否する。
function checkPasscode_(passcode) {
  const stored = PropertiesService.getScriptProperties().getProperty('APP_PASSCODE');
  if (!stored) {
    throw new Error('あいことば(APP_PASSCODE)が未設定です。GASのプロジェクト設定＞スクリプト プロパティで設定してください。');
  }
  if (String(passcode) !== String(stored)) {
    throw new Error('あいことばが違います');
  }
}

// 社長用の合い言葉。社長コメントの投稿時に毎回サーバー側で検証する。
// 画面の「社長モード」はあくまで表示の切り替えで、本人確認はここが本体。
function checkPresident_(pass) {
  const stored = PropertiesService.getScriptProperties().getProperty('PRESIDENT_PASSCODE');
  if (!stored) {
    throw new Error('社長用あいことば(PRESIDENT_PASSCODE)が未設定です。GASのプロジェクト設定＞スクリプト プロパティで設定してください。');
  }
  if (String(pass) !== String(stored)) {
    throw new Error('社長用のあいことばが違います');
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
//   v6: 「確認日時」列を末尾に追加（New表示を全員共有の1フラグで管理するため）
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

  // v3/v6: 末尾に列を追加（部門・社長コメント日時・確認日時）。途中に挿すと既存データがズレるので必ず末尾。
  const needDept = header.indexOf('部門') < 0;
  const needSeen = header.indexOf('確認日時') < 0;
  if (structural || needDept || needSeen) {
    // 物理的な列数が足りなければ先に足す（足りない状態で範囲書き込みするとエラーになる）
    if (sh.getMaxColumns() < IDEA_HEADERS.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), IDEA_HEADERS.length - sh.getMaxColumns());
    }
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

// 使わなくなったシートの後片付け。運用をやめたシートを削除する。
//   「既読」  … 既読を名前ごとにサーバー管理していた頃の名残（今はNew方式で不要）
//   「イベント」… 採用月/イベント運用の廃止で不要
// どちらもアプリのコードからは一切参照していないので、残っていれば削除して問題ない。
function cleanupLegacySheets_() {
  const ss = getSpreadsheet_();
  ['既読', 'イベント'].forEach(function(nm) {
    const sh = ss.getSheetByName(nm);
    // シートは最低1枚必要なので、最後の1枚のときは消さない（通常は該当しない）
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
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

function getMaterialMasterSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(MMASTER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MMASTER_SHEET_NAME);
    initSheet_(sh, MMASTER_HEADERS, MMASTER_TEXT_COLS, 1000);
    sh.getRange(2, 1, MMASTER_SEED.length, MMASTER_HEADERS.length).setValues(MMASTER_SEED);
  }
  return sh;
}

function getPosterSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(POSTER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(POSTER_SHEET_NAME);
    initSheet_(sh, POSTER_HEADERS, POSTER_TEXT_COLS, 1000);
  } else {
    // 廃止した「タイトル」列があれば削除する（1回だけ。以降は何もしない）
    const lastCol = sh.getLastColumn();
    if (lastCol >= 1) {
      const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
      const idx = header.indexOf('タイトル');
      if (idx >= 0) {
        sh.deleteColumn(idx + 1);
        sh.getRange(1, 1, 1, POSTER_HEADERS.length).setValues([POSTER_HEADERS]);
      }
    }
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

function getPlanSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PLAN_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PLAN_SHEET_NAME);
    initSheet_(sh, PLAN_HEADERS, PLAN_TEXT_COLS, 500);
  }
  return sh;
}

function getDevTaskSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(DEVTASK_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DEVTASK_SHEET_NAME);
    initSheet_(sh, DEVTASK_HEADERS, DEVTASK_TEXT_COLS, 2000);
  }
  return sh;
}

function getOwnerSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(OWNER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(OWNER_SHEET_NAME);
    initSheet_(sh, [OWNER_HEADER], [1], 200);
    sh.getRange(2, 1, OWNER_SEED.length, 1).setValues(OWNER_SEED.map(o => [o]));
  }
  return sh;
}

function getProgressSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PROG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PROG_SHEET_NAME);
    initSheet_(sh, PROG_HEADERS, PROG_TEXT_COLS, 3000);
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
  migrateOwners_();
  migrateDevTasksV2_();
  const ideas = getIdeas_();
  return {
    ideas: ideas,
    stores: getStoreList_(),
    materialMasters: getMaterialMasterList_(),
    categories: getCategoryList_(),
    tags: getTagList_(),
    deptOptions: DEPT_OPTIONS,
    statusOptions: STATUS_OPTIONS,
    unitOptions: UNIT_OPTIONS,
    yieldUnitOptions: YIELD_UNIT_OPTIONS,
    allergenOptions: ALLERGEN_OPTIONS,
    storageOptions: STORAGE_OPTIONS,
    plans: getPlanList_(),
    devTasks: getDevTaskList_(),
    owners: getOwnerList_(),
    seasons: SEASONS,
    taskStatusOptions: TASK_STATUS_OPTIONS,
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
    comments: getCommentsFor_(id),
    posters: getPostersFor_(id)
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
    status: cell_(row, IDEA_COL.status) || STATUS_DEFAULT,
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
    presAt: cell_(row, IDEA_COL.presAt),
    seenAt: cell_(row, IDEA_COL.seenAt)
  };
}

// presAt・seenAt はフォームからは編集できない値なので、呼び出し側が
// 「新規なら空」「更新なら旧行の値」を明示的に渡す（dataに紛れ込ませない）
function ideaToRow_(id, data, costs, photo1, photo2, updatedAt, presAt, seenAt) {
  const row = new Array(IDEA_HEADERS.length).fill('');
  const set = (col, v) => { row[col - 1] = v; };

  set(IDEA_COL.id, id);
  set(IDEA_COL.name, String(data.name || '').trim());
  set(IDEA_COL.catch, data.catch || '');
  set(IDEA_COL.store, String(data.store || '').trim());
  set(IDEA_COL.date, data.date || todayStr_());
  set(IDEA_COL.status, data.status || STATUS_DEFAULT);
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
  set(IDEA_COL.seenAt, seenAt || '');

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
    const row = ideaToRow_(id, data, costs, photo1, photo2, updatedAt, '', '');

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
                           cell_(oldRow, IDEA_COL.presAt), cell_(oldRow, IDEA_COL.seenAt));

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
    deletePostersFor_(id);
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
  // row（物理行番号）も返す。社長コメントの削除はこの行番号を指定して行う。
  return sh.getRange(2, 1, lastRow - 1, COMMENT_HEADERS.length).getValues()
    .map((r, idx) => ({
      row: idx + 2,
      _id: String(r[0]),
      kind: r[1] || KIND_TO_PRES,
      author: r[2] || '',
      text: r[3] || '',
      at: r[4] || ''
    }))
    .filter(c => c._id === String(ideaId))
    .map(c => ({ row: c.row, kind: c.kind, author: c.author, text: c.text, at: c.at }));
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

// 社長からのコメントを1件削除する（社長モードのときだけ）。
// row は getCommentsFor_ が返す物理行番号。念のため ideaId と種別を照合してから消す。
// 削除後は残っている社長コメントの最新日時で presAt を貼り直し、Newの状態を実態に合わせる。
function deleteComment_(ideaId, row, presidentPass) {
  checkPresident_(presidentPass);
  row = Number(row);
  if (!(row >= 2)) throw new Error('コメントが見つかりません');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getCommentSheet_();
    if (row > sh.getLastRow()) throw new Error('コメントが見つかりません（画面を再読み込みしてください）');
    const vals = sh.getRange(row, 1, 1, COMMENT_HEADERS.length).getValues()[0];
    if (String(vals[0]) !== String(ideaId))
      throw new Error('コメントが変わっています（画面を再読み込みしてください）');
    if (vals[1] !== KIND_FROM_PRES)
      throw new Error('社長からのコメントだけ削除できます');
    sh.deleteRow(row);

    const ideaSh = getIdeaSheet_();
    const rowNum = findIdeaRow_(ideaSh, ideaId);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + ideaId);
    const remaining = getCommentsFor_(ideaId);
    const presCmts = remaining.filter(c => c.kind === KIND_FROM_PRES);
    const newPresAt = presCmts.length ? presCmts[presCmts.length - 1].at : '';
    ideaSh.getRange(rowNum, IDEA_COL.presAt, 1, 1).setNumberFormat('@');
    ideaSh.getRange(rowNum, IDEA_COL.presAt, 1, 1).setValues([[newPresAt]]);

    const irow = ideaSh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];
    return { idea: rowToIdea_(irow, rowNum), comments: remaining };
  } finally {
    lock.releaseLock();
  }
}

// 「見ました」確認の切り替え。seen=true で確認日時を now にしてNewを全員ぶん消す。
// seen=false なら確認日時を空にしてNewを戻す（誰でも自由にON/OFFできる共有フラグ）。
// 引数省略時は true（従来互換）。
function markSeen_(ideaId, seen) {
  const on = (seen === undefined || seen === null) ? true : !!seen;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getIdeaSheet_();
    const rowNum = findIdeaRow_(sh, ideaId);
    if (rowNum < 0) throw new Error('アイデアが見つかりません: ' + ideaId);
    sh.getRange(rowNum, IDEA_COL.seenAt, 1, 1).setNumberFormat('@');
    sh.getRange(rowNum, IDEA_COL.seenAt, 1, 1).setValues([[on ? nowStr_() : '']]);
    const row = sh.getRange(rowNum, 1, 1, IDEA_HEADERS.length).getValues()[0];
    return { idea: rowToIdea_(row, rowNum) };
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
  return rewriteListColumn_(IDEA_COL.tags, target, newName);
}

// カンマ区切りの列（タグ・カテゴリ）の中の target を newName に置換する。
// newName が null なら target を取り除く（削除）。joinList_ を通すので重複は自動でまとまる。
// 変更した行数を返す。
function rewriteListColumn_(col, target, newName) {
  const sh = getIdeaSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const range = sh.getRange(2, col, lastRow - 1, 1);
  const values = range.getValues();
  let count = 0;
  values.forEach(r => {
    const items = splitList_(r[0]);
    if (items.indexOf(target) < 0) return;
    const next = newName === null
      ? items.filter(t => t !== target)
      : items.map(t => (t === target ? newName : t));
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
// ポスター（商品ごとに保管、店舗ラベル付き、ダウンロード可）
// ============================================================
// 指定アイデアのポスター一覧。表示用サムネイルとダウンロード用リンクを組み立てて返す。
function getPostersFor_(ideaId) {
  const sh = getPosterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const target = String(ideaId).trim();
  return sh.getRange(2, 1, lastRow - 1, POSTER_HEADERS.length).getValues()
    .filter(r => String(r[1]).trim() === target)
    .map(r => {
      const fileId = String(r[3]).trim();
      return {
        id: String(r[0]).trim(),
        ideaId: String(r[1]).trim(),
        store: String(r[2]).trim(),
        fileId: fileId,
        at: r[4] || '',
        by: r[5] || '',
        thumbUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000',
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId
      };
    })
    .sort((a, b) => (a.at < b.at ? 1 : -1));   // 新しい順
}

// data = { ideaId, store, image(dataURL), by }  ※タイトルは廃止
function addPoster_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);   // 画像アップロードを含むので長めに
  try {
    data = data || {};
    const ideaId = String(data.ideaId || '').trim();
    if (!ideaId) throw new Error('商品が特定できません');
    if (findIdeaRow_(getIdeaSheet_(), ideaId) < 0) throw new Error('商品が見つかりません: ' + ideaId);
    const store = String(data.store || '').trim();
    if (!store) throw new Error('使った店舗を選択してください');
    if (getStoreList_().every(s => s.name !== store)) throw new Error('存在しない店舗です: ' + store);
    const img = String(data.image || '');
    if (img.indexOf('data:') !== 0) throw new Error('ポスター画像を選択してください');

    const fileId = uploadToFolder_(img, getPosterFolder_(), 'poster_');
    const sh = getPosterSheet_();
    const id = generateNextPosterId_(sh);
    const rowNum = sh.getLastRow() + 1;
    POSTER_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, POSTER_HEADERS.length).setValues([[
      id, ideaId, store, fileId, nowStr_(), String(data.by || '').trim()
    ]]);
    return getPostersFor_(ideaId);
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
    const ideaId = String(rows[idx][1]).trim();
    const fileId = String(rows[idx][3]).trim();
    sh.deleteRow(idx + 2);
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
    return getPostersFor_(ideaId);
  } finally {
    lock.releaseLock();
  }
}

// アイデア削除時に、そのアイデアのポスター行とDriveファイルをまとめて消す
function deletePostersFor_(ideaId) {
  const sh = getPosterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const rows = sh.getRange(2, 1, lastRow - 1, POSTER_HEADERS.length).getValues();
  // 下から消す（上から消すと行番号がずれる）
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][1]).trim() === String(ideaId).trim()) {
      const fileId = String(rows[i][4]).trim();
      sh.deleteRow(i + 2);
      if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
    }
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

// カテゴリの改名。マスタ側を書き換え、アイデア側のカテゴリ列も一括で追随させる。
function renameCategory_(oldName, newName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const to = validateMasterName_(newName, 'カテゴリ名');
    const from = String(oldName || '').trim();
    const list = getCategoryList_();
    if (list.indexOf(from) < 0) throw new Error('カテゴリが見つかりません: ' + from);
    if (to !== from && list.indexOf(to) >= 0) throw new Error('同じ名前のカテゴリが既にあります');

    const sh = getCategorySheet_();
    const lastRow = sh.getLastRow();
    const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const idx = names.findIndex(r => String(r[0]).trim() === from);
    if (idx >= 0) {
      sh.getRange(idx + 2, 1, 1, 1).setNumberFormat('@');
      sh.getRange(idx + 2, 1, 1, 1).setValues([[to]]);
    }
    const renamed = (to !== from) ? rewriteListColumn_(IDEA_COL.categories, from, to) : 0;
    return { categories: getCategoryList_(), renamed: renamed };
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
// 材料マスタ
// ============================================================
// 各値は「基準量(baseQty) の 単位(unit) あたり」の値。レシピでは (使う量 / baseQty) を掛ける。
function mmasterRowToObj_(r) {
  return {
    name: String(r[0]).trim(),
    order: numOrZero_(r[1]),
    baseQty: numOrZero_(r[2]) || 1,   // 0だと除算できないので最低1
    unit: r[3] || 'g',
    price: numOrZero_(r[4]),
    kcal: numOrBlank_(r[5]),
    protein: numOrBlank_(r[6]),
    fat: numOrBlank_(r[7]),
    carb: numOrBlank_(r[8]),
    salt: numOrBlank_(r[9]),
    allergens: splitList_(r[10]),
    supplier: r[11] || '',
    memo: r[12] || ''
  };
}

function getMaterialMasterList_() {
  const sh = getMaterialMasterSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, MMASTER_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(mmasterRowToObj_)
    .sort((a, b) => (a.order - b.order) || (a.name < b.name ? -1 : 1));
}

// 追加も更新も1つの関数で扱う。oldName が空なら新規、あれば更新（改名時はレシピ側の材料名も追随）。
// data = { name, baseQty, unit, price, kcal, protein, fat, carb, salt, allergens[], supplier, memo }
function saveMaterial_(oldName, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    data = data || {};
    const name = validateMasterName_(data.name, '材料名');
    if (numOrZero_(data.baseQty) <= 0) throw new Error('基準量は1以上で入力してください');
    const sh = getMaterialMasterSheet_();
    const lastRow = sh.getLastRow();
    const names = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
    const old = String(oldName || '').trim();
    const editingIdx = old ? names.indexOf(old) : -1;

    // 同名重複チェック（自分自身は除く）
    const dup = names.indexOf(name);
    if (dup >= 0 && dup !== editingIdx) throw new Error('同じ名前の材料が既にあります');

    const order = editingIdx >= 0
      ? numOrZero_(sh.getRange(editingIdx + 2, 2, 1, 1).getValue())
      : (names.length + 1);
    const row = [
      name, order, numOrZero_(data.baseQty), String(data.unit || 'g').trim(),
      numOrZero_(data.price),
      numOrBlank_(data.kcal), numOrBlank_(data.protein), numOrBlank_(data.fat),
      numOrBlank_(data.carb), numOrBlank_(data.salt),
      joinList_(data.allergens), String(data.supplier || '').trim(), String(data.memo || '').trim()
    ];

    const rowNum = editingIdx >= 0 ? editingIdx + 2 : sh.getLastRow() + 1;
    MMASTER_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, MMASTER_HEADERS.length).setValues([row]);

    // 改名したらレシピ（材料シート）の材料名も一括で追随させる
    if (old && old !== name) renameRecipeMaterial_(old, name);

    return getMaterialMasterList_();
  } finally {
    lock.releaseLock();
  }
}

function deleteMaterial_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getMaterialMasterSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      const idx = names.findIndex(r => String(r[0]).trim() === String(name).trim());
      if (idx >= 0) sh.deleteRow(idx + 2);
    }
    // レシピ側の行はあえて消さない（過去のレシピの材料名として残す）。
    // マスタから外れた材料は栄養・アレルギーの自動集計に乗らなくなるだけ。
    return getMaterialMasterList_();
  } finally {
    lock.releaseLock();
  }
}

function reorderMaterialsMaster_(names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    reorderMasterRows_(getMaterialMasterSheet_(), MMASTER_HEADERS, MMASTER_TEXT_COLS, names, 2);
    return getMaterialMasterList_();
  } finally {
    lock.releaseLock();
  }
}

// 材料シート（レシピの子）の「材料名」列を一括置換する（材料マスタ改名時のカスケード）
function renameRecipeMaterial_(oldName, newName) {
  const sh = getMaterialSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const range = sh.getRange(2, 3, lastRow - 1, 1);   // 3列目 = 材料名
  const values = range.getValues();
  let count = 0;
  values.forEach(r => {
    if (String(r[0]).trim() === String(oldName).trim()) { r[0] = newName; count++; }
  });
  if (count) range.setValues(values);
  return count;
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

// ============================================================
// 開発計画（季節メニューのタスク管理）
// ============================================================
// 「年度×季節」で1計画。年度は5月始まり（2026年度＝2026/5〜2027/4）。
// 進捗はタスクの状態から画面側で自動計算する（シートには保存しない）。

function getPlanList_() {
  const sh = getPlanSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, PLAN_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      id: String(r[PLAN_COL.id - 1]),
      year: String(r[PLAN_COL.year - 1]),
      season: String(r[PLAN_COL.season - 1]),
      serveStart: String(r[PLAN_COL.serveStart - 1] || ''),
      serveEnd: String(r[PLAN_COL.serveEnd - 1] || ''),
      devStart: String(r[PLAN_COL.devStart - 1] || ''),
      devEnd: String(r[PLAN_COL.devEnd - 1] || ''),
      memo: String(r[PLAN_COL.memo - 1] || '')
    }));
}

function getDevTaskList_() {
  const sh = getDevTaskSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const prog = getTaskProgressMap_();
  return sh.getRange(2, 1, lastRow - 1, DEVTASK_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      const id = String(r[DEVTASK_COL.id - 1]);
      return {
        id: id,
        planId: String(r[DEVTASK_COL.planId - 1]),
        order: numOrZero_(r[DEVTASK_COL.order - 1]),
        name: String(r[DEVTASK_COL.name - 1] || ''),
        start: String(r[DEVTASK_COL.start - 1] || ''),
        end: String(r[DEVTASK_COL.end - 1] || ''),
        progress: prog[id] || {}   // { グループ名: { status, result } }
      };
    })
    .sort((a, b) => (a.order - b.order));
}

// 進捗シートを タスクID → { グループ: {status, result} } の索引にする
function getTaskProgressMap_() {
  const sh = getProgressSheet_();
  const lastRow = sh.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  sh.getRange(2, 1, lastRow - 1, PROG_HEADERS.length).getValues().forEach(r => {
    const id = String(r[PROG_COL.taskId - 1]).trim();
    const g = String(r[PROG_COL.group - 1]).trim();
    if (!id || !g) return;
    (map[id] = map[id] || {})[g] = {
      status: String(r[PROG_COL.status - 1] || TASK_STATUS_DEFAULT),
      result: String(r[PROG_COL.result - 1] || '')
    };
  });
  return map;
}

// グループごとの進捗（状態・結果）を更新。タスク×グループの行を上書き（無ければ追加）
function setTaskProgress_(taskId, group, data) {
  data = data || {};
  const g = String(group || '').trim();
  if (!g) throw new Error('グループが指定されていません');
  const status = TASK_STATUS_OPTIONS.indexOf(data.status) >= 0 ? data.status : TASK_STATUS_DEFAULT;
  const result = String(data.result || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (findDevTaskRow_(getDevTaskSheet_(), taskId) < 0) throw new Error('タスクが見つかりません: ' + taskId);
    const sh = getProgressSheet_();
    const lastRow = sh.getLastRow();
    let rowNum = -1;
    if (lastRow >= 2) {
      const vals = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]) === String(taskId) && String(vals[i][1]).trim() === g) { rowNum = i + 2; break; }
      }
    }
    if (rowNum < 0) rowNum = sh.getLastRow() + 1;
    PROG_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, PROG_HEADERS.length).setValues([[String(taskId), g, status, result, nowStr_()]]);
    return { devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

function getOwnerList_() {
  const sh = getOwnerSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v !== '');
}

// 計画ID・タスクIDの採番。既存IDの最大値+1（P0001 / T0001 形式）
function maxPrefixedNum_(sh, col, prefix) {
  const lastRow = sh.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    sh.getRange(2, col, lastRow - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(new RegExp('^' + prefix + '(\\d+)$'));
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return max;
}
function nextPrefixedId_(sh, col, prefix) {
  return prefix + ('0000' + (maxPrefixedNum_(sh, col, prefix) + 1)).slice(-4);
}

// 年度＋季節の計画行を返す（無ければ作る）。タスク追加時に日程未設定でも困らないように。
function ensurePlan_(year, season) {
  const y = String(year || '').trim();
  const s = String(season || '').trim();
  if (!/^\d{4}$/.test(y)) throw new Error('年度が不正です: ' + year);
  if (SEASONS.indexOf(s) < 0) throw new Error('季節が不正です: ' + season);
  const sh = getPlanSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, PLAN_HEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][PLAN_COL.year - 1]) === y && String(rows[i][PLAN_COL.season - 1]) === s) {
        return { rowNum: i + 2, id: String(rows[i][PLAN_COL.id - 1]) };
      }
    }
  }
  const id = nextPrefixedId_(sh, PLAN_COL.id, 'P');
  const rowNum = sh.getLastRow() + 1;
  const row = new Array(PLAN_HEADERS.length).fill('');
  row[PLAN_COL.id - 1] = id;
  row[PLAN_COL.year - 1] = y;
  row[PLAN_COL.season - 1] = s;
  row[PLAN_COL.updatedAt - 1] = nowStr_();
  sh.getRange(rowNum, 1, 1, PLAN_HEADERS.length).setNumberFormat('@');
  sh.getRange(rowNum, 1, 1, PLAN_HEADERS.length).setValues([row]);
  return { rowNum: rowNum, id: id };
}

// 計画（日程・メモ）の保存。year+season で行を特定し、上書き（無ければ新規作成）。
// data = { year, season, serveStart, serveEnd, devStart, devEnd, memo }
function savePlan_(data) {
  data = data || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const found = ensurePlan_(data.year, data.season);
    const sh = getPlanSheet_();
    const row = sh.getRange(found.rowNum, 1, 1, PLAN_HEADERS.length).getValues()[0];
    row[PLAN_COL.serveStart - 1] = dateStrOrEmpty_(data.serveStart);
    row[PLAN_COL.serveEnd - 1] = dateStrOrEmpty_(data.serveEnd);
    row[PLAN_COL.devStart - 1] = dateStrOrEmpty_(data.devStart);
    row[PLAN_COL.devEnd - 1] = dateStrOrEmpty_(data.devEnd);
    row[PLAN_COL.memo - 1] = String(data.memo || '');
    row[PLAN_COL.updatedAt - 1] = nowStr_();
    sh.getRange(found.rowNum, 1, 1, PLAN_HEADERS.length).setNumberFormat('@');
    sh.getRange(found.rowNum, 1, 1, PLAN_HEADERS.length).setValues([row]);
    return { plans: getPlanList_() };
  } finally {
    lock.releaseLock();
  }
}

// 'yyyy/MM/dd' 形式だけ通す。<input type="date"> の 'yyyy-MM-dd' も受けて変換する
function dateStrOrEmpty_(v) {
  const s = String(v == null ? '' : v).trim().replace(/-/g, '/');
  if (!s) return '';
  if (!/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(s)) return '';
  const p = s.split('/');
  return p[0] + '/' + ('0' + p[1]).slice(-2) + '/' + ('0' + p[2]).slice(-2);
}

// タスク追加。data = { year, season, name, start, end }（名前・期間は全グループ共通）
function addDevTask_(data) {
  data = data || {};
  const name = String(data.name || '').trim();
  if (!name) throw new Error('タスク名を入力してください');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const plan = ensurePlan_(data.year, data.season);
    const sh = getDevTaskSheet_();
    const maxOrder = getDevTaskList_().filter(t => t.planId === plan.id)
      .reduce((m, t) => Math.max(m, t.order), 0);
    const id = nextPrefixedId_(sh, DEVTASK_COL.id, 'T');
    const rowNum = sh.getLastRow() + 1;
    const row = new Array(DEVTASK_HEADERS.length).fill('');
    row[DEVTASK_COL.id - 1] = id;
    row[DEVTASK_COL.planId - 1] = plan.id;
    row[DEVTASK_COL.order - 1] = maxOrder + 1;
    row[DEVTASK_COL.name - 1] = name;
    row[DEVTASK_COL.start - 1] = dateStrOrEmpty_(data.start);
    row[DEVTASK_COL.end - 1] = dateStrOrEmpty_(data.end);
    row[DEVTASK_COL.updatedAt - 1] = nowStr_();
    DEVTASK_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, DEVTASK_HEADERS.length).setValues([row]);
    return { plans: getPlanList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

// タスクの一括追加（標準工程の自動投入用）。data = { year, season, tasks: [{ name }] }
// 名前だけで作る（期間は後から設定・進捗は各グループが更新）。
function addDevTasksBulk_(data) {
  data = data || {};
  const items = (data.tasks || [])
    .map(t => ({ name: String(t && t.name || '').trim() }))
    .filter(t => t.name);
  if (!items.length) throw new Error('追加するタスクがありません');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const plan = ensurePlan_(data.year, data.season);
    const sh = getDevTaskSheet_();
    let order = getDevTaskList_().filter(t => t.planId === plan.id)
      .reduce((m, t) => Math.max(m, t.order), 0);
    let seq = maxPrefixedNum_(sh, DEVTASK_COL.id, 'T');
    const now = nowStr_();
    const startRow = sh.getLastRow() + 1;
    const rows = items.map(t => {
      const row = new Array(DEVTASK_HEADERS.length).fill('');
      row[DEVTASK_COL.id - 1] = 'T' + ('0000' + (++seq)).slice(-4);
      row[DEVTASK_COL.planId - 1] = plan.id;
      row[DEVTASK_COL.order - 1] = ++order;
      row[DEVTASK_COL.name - 1] = t.name;
      row[DEVTASK_COL.updatedAt - 1] = now;
      return row;
    });
    DEVTASK_TEXT_COLS.forEach(col => sh.getRange(startRow, col, rows.length, 1).setNumberFormat('@'));
    sh.getRange(startRow, 1, rows.length, DEVTASK_HEADERS.length).setValues(rows);
    return { plans: getPlanList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

function findDevTaskRow_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  const vals = sh.getRange(2, DEVTASK_COL.id, lastRow - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateDevTask_(taskId, data) {
  data = data || {};
  const name = String(data.name || '').trim();
  if (!name) throw new Error('タスク名を入力してください');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getDevTaskSheet_();
    const rowNum = findDevTaskRow_(sh, taskId);
    if (rowNum < 0) throw new Error('タスクが見つかりません: ' + taskId);
    const row = sh.getRange(rowNum, 1, 1, DEVTASK_HEADERS.length).getValues()[0];
    row[DEVTASK_COL.name - 1] = name;
    row[DEVTASK_COL.start - 1] = dateStrOrEmpty_(data.start);
    row[DEVTASK_COL.end - 1] = dateStrOrEmpty_(data.end);
    row[DEVTASK_COL.updatedAt - 1] = nowStr_();
    DEVTASK_TEXT_COLS.forEach(col => sh.getRange(rowNum, col, 1, 1).setNumberFormat('@'));
    sh.getRange(rowNum, 1, 1, DEVTASK_HEADERS.length).setValues([row]);
    return { plans: getPlanList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

function deleteDevTask_(taskId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getDevTaskSheet_();
    const rowNum = findDevTaskRow_(sh, taskId);
    if (rowNum < 0) throw new Error('タスクが見つかりません: ' + taskId);
    sh.deleteRow(rowNum);
    // このタスクの進捗行も下から順に消す
    const psh = getProgressSheet_();
    const plast = psh.getLastRow();
    if (plast >= 2) {
      const vals = psh.getRange(2, 1, plast - 1, 1).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][0]) === String(taskId)) psh.deleteRow(i + 2);
      }
    }
    return { plans: getPlanList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

// 並び替え。taskIds の順に「表示順」を1..Nで振り直す（他計画のタスクは触らない）
function reorderDevTasks_(planId, taskIds) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getDevTaskSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { devTasks: [] };
    const pos = {};
    (taskIds || []).forEach((id, i) => { pos[String(id)] = i + 1; });
    const rows = sh.getRange(2, 1, lastRow - 1, DEVTASK_HEADERS.length).getValues();
    rows.forEach((r, i) => {
      if (String(r[DEVTASK_COL.planId - 1]) !== String(planId)) return;
      const p = pos[String(r[DEVTASK_COL.id - 1])];
      if (p === undefined) return;   // 画面に無いタスク（同時追加など）は順序を変えない
      sh.getRange(i + 2, DEVTASK_COL.order, 1, 1).setValues([[p]]);
    });
    return { devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

// ---- 担当マスタ ----
// 統合移行（2026-08: 長崎エリア＋佐賀エリア → 佐賀・長崎エリア）。
// 旧名がマスタに残っていたら改名・統合し、タスクの担当も書き換える。
// 一度走れば旧名は無くなるので、毎回呼んでも何もしない（アプリを開いたときに実行）。
function migrateOwners_() {
  const merged = '佐賀・長崎エリア';
  const olds = ['長崎エリア', '佐賀エリア'];
  const list = getOwnerList_();
  if (!olds.some(o => list.indexOf(o) >= 0)) return;
  const sh = getOwnerSheet_();
  olds.forEach(o => {
    if (getOwnerList_().indexOf(o) < 0) return;
    renameInMasterColumn_(sh, o, merged);   // 統合先が既にあれば旧行は消える
    rewriteOwnerColumn_(o, merged);
  });
}

function addOwner_(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const trimmed = validateMasterName_(name, '担当名');
    if (getOwnerList_().indexOf(trimmed) >= 0) throw new Error('同じ名前の担当が既にあります');
    const sh = getOwnerSheet_();
    const rowNum = sh.getLastRow() + 1;
    sh.getRange(rowNum, 1, 1, 1).setNumberFormat('@');
    sh.getRange(rowNum, 1, 1, 1).setValues([[trimmed]]);
    return getOwnerList_();
  } finally {
    lock.releaseLock();
  }
}

// 改名はマスタと、そのグループの進捗行の両方を書き換える
function renameOwner_(oldName, newName) {
  const from = String(oldName).trim();
  const to = validateMasterName_(newName, '担当名');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    renameInMasterColumn_(getOwnerSheet_(), from, to);
    rewriteOwnerColumn_(from, to);   // 旧方式の担当列にも念のため反映（v2移行後は空）
    return { renamed: rewriteProgressGroup_(from, to), owners: getOwnerList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

// 担当は使用中でも削除できる（そのグループの進捗行も消える）。顔ぶれがよく変わる前提のため
function deleteOwner_(name) {
  const target = String(name).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    deleteFromMasterColumn_(getOwnerSheet_(), target);
    rewriteOwnerColumn_(target, null);
    return { removed: rewriteProgressGroup_(target, null), owners: getOwnerList_(), devTasks: getDevTaskList_() };
  } finally {
    lock.releaseLock();
  }
}

// 進捗シートのグループ名を書き換える（newName=null なら行ごと削除）。書き換えた行数を返す。
// 改名先の行が同じタスクに既にある場合は旧行を消す（重複を作らない）
function rewriteProgressGroup_(target, newName) {
  const sh = getProgressSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const rows = sh.getRange(2, 1, lastRow - 1, PROG_HEADERS.length).getValues();
  const hasNew = {};
  if (newName) rows.forEach(r => {
    if (String(r[PROG_COL.group - 1]).trim() === newName) hasNew[String(r[PROG_COL.taskId - 1])] = true;
  });
  let count = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][PROG_COL.group - 1]).trim() !== target) continue;
    count++;
    if (newName === null || hasNew[String(rows[i][PROG_COL.taskId - 1])]) { sh.deleteRow(i + 2); continue; }
    sh.getRange(i + 2, PROG_COL.group, 1, 1).setNumberFormat('@');
    sh.getRange(i + 2, PROG_COL.group, 1, 1).setValues([[newName]]);
  }
  return count;
}

// 旧方式（グループごとにタスクをコピー）→ 新方式（タスク1本＋グループ別進捗）への一括移行。
// 計画ID×タスク名でまとめ、正本1行に統合。各コピーの担当×状態・結果は進捗シートへ移す。
// 同名タスクを意図的に複数作る運用と衝突しないよう、実行済みフラグ（Script Properties）で1回だけ走らせる
function migrateDevTasksV2_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DEVTASK_MODEL_V2')) return;
  const sh = getDevTaskSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, DEVTASK_HEADERS.length).getValues();
    const byKey = {};
    rows.forEach((r, i) => {
      if (String(r[DEVTASK_COL.id - 1]).trim() === '') return;
      const key = String(r[DEVTASK_COL.planId - 1]) + ' ' + String(r[DEVTASK_COL.name - 1]).trim();
      (byKey[key] = byKey[key] || []).push({ r: r, row: i + 2 });
    });
    const progAdd = [];
    const delRows = [];
    Object.keys(byKey).forEach(key => {
      const arr = byKey[key];
      arr.sort((a, b) => numOrZero_(a.r[DEVTASK_COL.order - 1]) - numOrZero_(b.r[DEVTASK_COL.order - 1]));
      const canon = arr[0];
      const canonId = String(canon.r[DEVTASK_COL.id - 1]);
      let mnS = '', mxE = '';
      arr.forEach(x => {
        const st = String(x.r[DEVTASK_COL.start - 1] || '');
        const en = String(x.r[DEVTASK_COL.end - 1] || '');
        if (st && (!mnS || st < mnS)) mnS = st;
        if (en && (!mxE || en > mxE)) mxE = en;
        const status = String(x.r[DEVTASK_COL.status - 1] || '');
        const result = String(x.r[DEVTASK_COL.result - 1] || '');
        splitList_(x.r[DEVTASK_COL.owners - 1]).forEach(g => {
          progAdd.push([canonId, g, status || TASK_STATUS_DEFAULT, result, nowStr_()]);
        });
        if (x !== canon) delRows.push(x.row);
      });
      // 正本行：期間はコピー群の最小開始〜最大終了、旧3列（担当/状態/結果）は空にする
      canon.r[DEVTASK_COL.start - 1] = mnS;
      canon.r[DEVTASK_COL.end - 1] = mxE;
      canon.r[DEVTASK_COL.owners - 1] = '';
      canon.r[DEVTASK_COL.status - 1] = '';
      canon.r[DEVTASK_COL.result - 1] = '';
      DEVTASK_TEXT_COLS.forEach(col => sh.getRange(canon.row, col, 1, 1).setNumberFormat('@'));
      sh.getRange(canon.row, 1, 1, DEVTASK_HEADERS.length).setValues([canon.r]);
    });
    // 重複行の削除は行番号が大きい方から
    delRows.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
    if (progAdd.length) {
      const psh = getProgressSheet_();
      const startRow = psh.getLastRow() + 1;
      PROG_TEXT_COLS.forEach(col => psh.getRange(startRow, col, progAdd.length, 1).setNumberFormat('@'));
      psh.getRange(startRow, 1, progAdd.length, PROG_HEADERS.length).setValues(progAdd);
    }
  }
  props.setProperty('DEVTASK_MODEL_V2', '1');
}

function reorderOwners_(names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    reorderMasterRows_(getOwnerSheet_(), [OWNER_HEADER], [1], names, 0);
    return getOwnerList_();
  } finally {
    lock.releaseLock();
  }
}

// 開発タスクの「担当」列（カンマ区切り）の中の target を newName に置換（null で削除）。
// 変更した行数を返す。アイデアのタグ列に対する rewriteListColumn_ の開発タスク版。
function rewriteOwnerColumn_(target, newName) {
  const sh = getDevTaskSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const range = sh.getRange(2, DEVTASK_COL.owners, lastRow - 1, 1);
  const values = range.getValues();
  let count = 0;
  values.forEach(r => {
    const items = splitList_(r[0]);
    if (items.indexOf(target) < 0) return;
    const next = newName === null
      ? items.filter(x => x !== target)
      : items.map(x => (x === target ? newName : x));
    r[0] = joinList_(next);
    count++;
  });
  if (count) {
    range.setNumberFormat('@');
    range.setValues(values);
  }
  return count;
}
