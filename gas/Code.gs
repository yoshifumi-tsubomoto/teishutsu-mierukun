/**
 * 提出物チェックアプリ v6(GAS + スプレッドシート)
 * 主な機能:
 *  - QRスキャン記録(手入力/ハンディリーダー/ライブカメラ/写真読み取り📸)
 *  - UNDO/REDO(取消フラグ方式のソフトデリート)
 *  - 生徒ID・教材コードの自動生成(教材シート「教材名略称」・略称シート「教科」)
 *  - 対象組による生徒マイページの表示制御(学年・組一致 または 本人に記録あり)
 *  - 生徒マイページ: 緊急度順の表示(締切7日以内=急ぎ)・提出をあきらめる/元に戻す・自作タスク
 *  - 集計シート: 3段見出し(教科/教材名/範囲)・行フィルタ・絞り込みメニュー4種
 *  - QRラベル印刷(エーワン各種番手の実寸配置)・教材一覧表印刷
 *  - Classroom通知(再提出者のみ)・設定シートへのWebアプリURL/印刷ページリンク自動反映
 *  - 自作タスクは共有していない別スプレッドシートに保存(他職員から不可視)
 *  - 教員画面: 学年/組/教科/教材名/範囲で絞り込む「未提出・再提出リスト」表示
 *    (スプレッドシートを開かずスマホだけで完結、教材ごとにClassroom通知も可能)
 *  - 開発版(β環境)連携: IS_DEV_BUILDフラグ・DEV_SS_IDへのデータコピー機能
 *  - 初期シート作成時に「はじめにお読みください」シートを自動生成(他校への
 *    テンプレート配布時のセルフガイド)
 *  - デイリー教材モード(v6.5〜): 教材シートの「毎日」フラグで、毎日集める
 *    提出物を「教材×日付」単位で記録・判定(未提出判定・マイページ・
 *    Classroom通知・デイリー集計シートすべて日付キー対応)
 *  - 受付終了機能(v6.6〜): 教材単位で受付を締め切り、生徒の「やることリスト」から外す
 *  - 絞り込み連動・条件記憶(v6.7〜): 未提出・再提出リストの絞り込みを学年→教科→
 *    教材名→範囲で連動させ、前回の条件を教員ごとに記憶する
 *  - 教材タブ(v6.8〜): 教員画面から教材の登録・編集・削除。教材コードの自動採番、
 *    登録者・更新者の記録、推奨の教材名略称(略称シートの種別「教材」)
 *  - 締切での絞り込み(v6.9〜): 未提出リスト・教材タブで締切あり/過ぎた/7日以内/
 *    期間指定/締切なしに絞り、締切が近い順に並べる。教材タブにQRラベル印刷ボタン
 */

/** バージョン表示用。新しい版をリリースするときはここだけ書き換える */
const APP_VERSION = 'v6.9';
const IS_DEMO_BUILD = false; // 体験版コピーではtrueにする(画面表示・体験版モードの目印)
/** 開発者自身が新機能を試す「開発版」コピーでは true にする。
 *  本番と全く同じCode.gs/HTMLをそのまま使い、この1行だけを変える運用を想定
 *  (体験版とは別物。開発版は本番データのコピーを使った内部テスト専用)。 */
const IS_DEV_BUILD = false;

const SHEET_ROSTER   = '生徒名簿';
const SHEET_TEACHERS = '教員名簿';
const SHEET_MATERIAL = '教材';
const SHEET_LOG      = 'ログ';
const SHEET_SUMMARY  = '集計';
const SHEET_SETTINGS = '設定';
const SHEET_ABBR     = '略称';
const SHEET_TASKS    = '個人タスク';
const SHEET_READFIRST = 'はじめにお読みください';
const SHEET_DAILY_SUMMARY = 'デイリー集計';

/** 「提出をあきらめる」を表す予約語。教材シートの評価ラベルにこの文字列を
 *  使わないこと(通常は起こり得ないが念のため)。 */
const GIVEUP_LABEL = '提出しない';
/** 教材シート末尾の「誰が登録・更新したか」の列(v6.8〜、メニュー⑥で追加)。
 *  ヘッダー名で探すので、列が無い環境では記録をとばすだけで動作は変わらない */
const MATERIAL_AUDIT_COLS = ['登録者', '登録日時', '更新者', '更新日時'];
/** 追加列が無いときに画面へ出す案内。「⑥」だけでは画面上で何のことかわからないため、
 *  スプレッドシートのメニューであることまで書く */
const SETUP_COLS_HINT = 'スプレッドシートを開き、上部メニュー「提出物アプリ」→「⑥ 追加列を準備」を1回実行してください';
/** 推奨の教材名略称(v6.8〜)。メニュー⑥で略称シートに種別「教材」として足す。
 *  教材タブの略称欄に「推奨」として表示し、教材名から略称を自動で入れるのに使う */
const RECOMMENDED_MATERIAL_ABBRS = [
  ['ワーク', 'WK'], ['プリント', 'PR'], ['ノート', 'NT'], ['補充教材', 'HJ'],
  ['自学ノート', 'JG'], ['レポート', 'RP'], ['テスト', 'TS'], ['ドリル', 'DR'],
  ['資料集', 'SR'], ['振り返りシート', 'FR'],
];

/* ============ メニュー ============ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('提出物アプリ')
    .addItem('① 初期シートを作成', 'setupSheets')
    .addItem('② 生徒IDを自動生成(空欄のみ)', 'generateStudentIds')
    .addItem('③ 教材コードを自動生成(空欄のみ)', 'generateMaterialCodes')
    .addItem('④ 集計シートを更新', 'buildSummary')
    .addItem('⑤ 再提出者へClassroom通知(教材コード指定)', 'menuNotifyClassroom')
    .addItem('⑥ 追加列を準備(毎日・提出日・受付終了・登録者)', 'setupDailyColumns')
    .addItem('⑦ デイリー集計を更新', 'buildDailySummary')
    .addItem('🖨 印刷ページのリンクを設定シートに表示', 'showPrintPageLink')
    .addSeparator()
    .addItem('補充教材一覧(教材名で絞り込み)', 'showMaterialNamePicker')
    .addItem('教科別一覧(教科で絞り込み)', 'showSubjectPicker')
    .addItem('連番で絞り込み(教材コード末尾2桁)', 'showSerialRangePicker')
    .addItem('組み合わせて絞り込み(教科×教材名×連番)', 'showCombinedPicker')
    .addItem('集計の絞り込みを解除(すべて表示)', 'resetSummaryColumns')
    .addSeparator()
    .addItem('🧪 開発版へ名簿・教材・略称をコピー(開発者用)', 'copyDataToDevBuild')
    .addToUi();
}

/* ============ 初期セットアップ ============ */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 移行措置: 旧バージョンの「名簿」シートが残っていれば「生徒名簿」に
  // 改名する(データはそのまま引き継がれる。作り直しではない)
  const oldRoster = ss.getSheetByName('名簿');
  if (oldRoster && !ss.getSheetByName(SHEET_ROSTER)) {
    oldRoster.setName(SHEET_ROSTER);
  }

  const defs = [
    { name: SHEET_READFIRST,
      headers: [] },
    { name: SHEET_ROSTER,
      headers: ['生徒ID', '学年', '組', '番号', '氏名', 'Googleアカウント'] },
    { name: SHEET_TEACHERS,
      headers: ['教員名', 'Googleアカウント'] },
    { name: SHEET_MATERIAL,
      headers: ['教材コード', '学年', '対象組(空欄=全組)', '教科', '教材名', '教材名略称', '範囲',
                '提出期限(空欄可)', '評価ラベル(カンマ区切り)', '再提出扱い(カンマ区切り)'] },
    { name: SHEET_ABBR,
      headers: ['種別(教科)', '名称', '略称(半角英数)'] },
    { name: SHEET_LOG,
      headers: ['日時', '教材コード', '教材名', '生徒ID', '氏名', '評価', '回数', '取消'] },
    { name: SHEET_SETTINGS, headers: ['キー', '値'] },
    { name: SHEET_SUMMARY, headers: [] },
  ];
  defs.forEach(def => {
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (def.headers.length && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers])
        .setFontWeight('bold').setBackground('#e8eaed');
      sh.setFrozenRows(1);
    }
  });

  // 「はじめにお読みください」シートに簡易セットアップ手順を書き込む(他校へ
  // テンプレートとして配布したとき、このチャットの案内なしでも最初の一歩を
  // 踏み出せるようにするため)。既に何か書かれていれば上書きしない
  const readFirst = ss.getSheetByName(SHEET_READFIRST);
  if (readFirst.getLastRow() === 0) {
    readFirst.getRange(1, 1).setValue(
      '提出物チェックアプリ ' + APP_VERSION + ' - はじめにお読みください\n\n' +
      '① このシートは読むだけでOK(削除しても動作に影響しません)\n' +
      '② 「教員名簿」シートに、この アプリを使う先生のGoogleアカウントを入力\n' +
      '③ 「生徒名簿」シートに、生徒の学年・組・番号・氏名・Googleアカウントを入力\n' +
      '   (生徒ID列は空欄のままでOK。メニュー②で自動生成されます)\n' +
      '④ 「教材」シートに、チェックしたい教材を1行=1範囲で入力\n' +
      '   (教材コード列は空欄でOK。メニュー③で自動生成されます)\n' +
      '⑤ 「設定」シートの「職員アドレス」を、実際の先生のアドレスに書き換える\n' +
      '⑥ Apps Scriptエディタで「デプロイ→新しいデプロイ→ウェブアプリ」\n' +
      '   実行ユーザー:自分 / アクセス:組織内の全員、でデプロイ\n' +
      '⑦ 発行されたURLを一度開く(設定シートにURLが自動で記録されます)\n\n' +
      '詳しい手順は、このスプレッドシートを配布した先生から受け取った\n' +
      '「説明書」を参照してください。\n\n' +
      '――――――――――――――――――――\n' +
      'ご利用にあたって\n' +
      '本アプリは作者が個人で開発・公開しているものであり、学校や教育委員会の\n' +
      '公式なシステムではありません。\n' +
      '・内容の確認・対応を約束するものではありません\n' +
      '・ソースコードは自由に改変していただいて構いません(改変前に元ファイル\n' +
      '  のコピーを残しておくことをおすすめします)\n' +
      '・ただし改変された場合、そのコピーについてのサポートは対象外とさせて\n' +
      '  いただきます\n' +
      '・改変された場合は、元にしたバージョン番号と変更内容を教えていただける\n' +
      '  と、今後の対応の参考になります(任意ですが助かります)\n' +
      '・いただいたご要望は、反映するかどうか・時期を含めて開発者の判断と\n' +
      '  させていただきます\n' +
      '・このテンプレートを、さらに他の学校へ配布し直すのはご遠慮ください\n' +
      '  (配布は元のリンクからお願いします)\n\n' +
      'ご意見・不具合報告はこちら:\n' +
      '〔ここにフィードバックフォームの回答用URLを貼り付けてください〕'
    ).setWrap(true).setVerticalAlignment('top');
    readFirst.setColumnWidth(1, 640);
    readFirst.setRowHeight(1, 560);
  }

  const roster = ss.getSheetByName(SHEET_ROSTER);
  if (roster.getLastRow() === 1) {
    roster.getRange(2, 1, 2, 6).setValues([
      ['', 1, 3, 1, '見本 生徒A', 'student-a@example.com'],
      ['', 1, 3, 2, '見本 生徒B', 'student-b@example.com'],
    ]);
  }
  const teachers = ss.getSheetByName(SHEET_TEACHERS);
  if (teachers.getLastRow() === 1) {
    teachers.getRange(2, 1, 2, 2).setValues([
      ['見本 先生A', 'teacher-a@example.com'],
      ['見本 先生B', 'teacher-b@example.com'],
    ]);
  }
  const mat = ss.getSheetByName(SHEET_MATERIAL);
  if (mat.getLastRow() === 1) {
    mat.getRange(2, 1, 3, 10).setValues([
      ['', 1, '', '国語', 'ワーク', 'WK', 'p.1〜20',  '', '合格,再提出', '再提出'],
      ['', 1, '', '国語', 'ワーク', 'WK', 'p.21〜40', '', '合格,再提出', '再提出'],
      ['', 1, 'A,B', '数学', 'プリント', 'PR', '正負の数', '', '5,4,3,2,1', '2,1'],
    ]);
  }
  const abbr = ss.getSheetByName(SHEET_ABBR);
  if (abbr.getLastRow() === 1) {
    abbr.getRange(2, 1, 9, 3).setValues([
      ['教科', '国語', 'JPN'], ['教科', '数学', 'MAT'], ['教科', '英語', 'ENG'],
      ['教科', '理科', 'SCI'], ['教科', '社会', 'SOC'], ['教科', '音楽', 'MUS'],
      ['教科', '美術', 'ART'], ['教科', '保健体育', 'PE'], ['教科', '技術家庭', 'TEC'],
    ]);
  }
  const set = ss.getSheetByName(SHEET_SETTINGS);
  if (set.getLastRow() === 1) {
    set.getRange(2, 1, 3, 2).setValues([
      ['職員アドレス', 'teacher-a@example.com, teacher-b@example.com'],
      ['年度(西暦)', _fiscalYear()],
      ['ClassroomコースID', ''],
    ]);
  }
  SpreadsheetApp.getUi().alert(
    'シートを作成しました。\n' +
    '1) 略称シートで教科の略称を確認/追加\n' +
    '2) 教員名簿(教員名・Googleアカウント)を入力\n' +
    '3) 生徒名簿(ID空欄でOK)・教材(コード空欄・教材名略称は入力してOK)を入力\n' +
    '4) メニュー②③でID・コードを自動生成\n' +
    '5) 設定シートの年度を確認\n' +
    '(職員アドレスは今後「教員名簿」シートで管理します。設定シートの' +
    '「職員アドレス」は、緊急時の追加登録用として引き続き使えます)');
}

/** 日本の年度(4月始まり)を西暦で返す */
function _fiscalYear() {
  const d = new Date();
  return d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1; // 0=1月
}

/* ============ ID・コードの自動生成 ============ */
/**
 * 生徒ID自動生成: S-年度-学年組番号4桁(例 S-2025-1203)
 * ・空欄の行だけ生成。既にIDがある行は絶対に触らない(3年間固定のため)
 * ・年度は設定シート「年度(西暦)」を使用(未設定なら今年度)
 */
function generateStudentIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ROSTER);
  const vals = sh.getDataRange().getValues();
  const year = String(_settings()['年度(西暦)'] || _fiscalYear()).trim();

  const existing = {};
  for (let i = 1; i < vals.length; i++) {
    const id = String(vals[i][0]).trim();
    if (id) existing[id] = true;
  }

  let made = 0, skipped = 0;
  const errors = [];
  for (let i = 1; i < vals.length; i++) {
    const [id, grade, cls, num, name] = vals[i];
    if (String(id).trim()) { skipped++; continue; }        // 既存IDは不変
    if (!name) continue;                                   // 空行
    if (!grade || !cls || num === '' || num === null) {
      errors.push((i + 1) + '行目: 学年・組・番号が未入力');
      continue;
    }
    const num4 = String(grade) + String(cls) +
                 ('0' + String(num)).slice(-2);            // 例 1203
    let newId = 'S-' + year + '-' + num4;
    if (existing[newId]) {
      errors.push((i + 1) + '行目: ' + newId + ' が重複(手動で調整してください)');
      continue;
    }
    sh.getRange(i + 1, 1).setValue(newId);
    existing[newId] = true;
    made++;
  }
  SpreadsheetApp.getUi().alert(
    '生徒ID: ' + made + '件生成 / 既存' + skipped + '件は変更なし' +
    (errors.length ? '\n\n【要確認】\n' + errors.join('\n') : ''));
}

/**
 * 教材コード自動生成: 教科略称-教材名略称-連番2桁(例 JPN-WK-01)
 * ・空欄の行だけ生成。連番は同じ「教科-教材」の既存最大値の続きから
 * ・教科略称は「略称」シートを参照(教科名から自動判定)
 * ・教材名略称は教材シートの「教材名略称」列に直接入力してもらう
 *   (教材名からの自動推測はしない。誤爆を避けるため先生が確定する)
 */
function generateMaterialCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MATERIAL);
  const vals = sh.getDataRange().getValues();
  const abbr = _abbrMaps();
  const maxSerial = _materialSerialMax(vals);

  let made = 0;
  const errors = [];
  for (let i = 1; i < vals.length; i++) {
    const [code, grade, targetCls, subject, name, matAbbrRaw] = vals[i];
    if (String(code).trim()) continue;                     // 既存は不変
    if (!subject && !name) continue;                       // 空行

    const subAb = _findAbbr(abbr.subject, String(subject));
    const matAb = String(matAbbrRaw || '').trim().toUpperCase();

    const missing = [];
    if (!subAb) missing.push('教科「' + subject + '」の略称が略称シートにない');
    if (!matAb) missing.push('教材名略称が空欄');
    if (missing.length) {
      errors.push((i + 1) + '行目: ' + missing.join(' / '));
      continue;
    }

    const prefix = subAb + '-' + matAb;
    const serial = (maxSerial[prefix] || 0) + 1;
    maxSerial[prefix] = serial;
    const newCode = prefix + '-' + ('0' + serial).slice(-2);
    sh.getRange(i + 1, 1).setValue(newCode);
    made++;
  }
  SpreadsheetApp.getUi().alert(
    '教材コード: ' + made + '件生成' +
    (errors.length ? '\n\n【要確認】\n' + errors.join('\n') : ''));
}

/** 既存の教材コードから、prefix(教科略称-教材名略称)ごとの最大連番を数える。
 *  メニュー③と教材タブ(createMaterials)の両方がこれを使い、採番規則を1か所にする */
function _materialSerialMax(vals) {
  const maxSerial = {};
  for (let i = 1; i < vals.length; i++) {
    const code = String(vals[i][0]).trim();
    const m = code.match(/^([A-Za-z0-9]+-[A-Za-z0-9]+)-(\d+)$/);
    if (m) {
      const n = parseInt(m[2], 10);
      if (!maxSerial[m[1]] || n > maxSerial[m[1]]) maxSerial[m[1]] = n;
    }
  }
  return maxSerial;
}

function _abbrMaps() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ABBR);
  const vals = sh ? sh.getDataRange().getValues() : [];
  const subject = [], material = [];
  for (let i = 1; i < vals.length; i++) {
    const [kind, name, ab] = vals[i];
    if (!name || !ab) continue;
    const rec = { name: String(name).trim(), ab: String(ab).trim().toUpperCase() };
    if (String(kind).indexOf('教科') !== -1) subject.push(rec);
    else material.push(rec);
  }
  return { subject, material };
}

/** 完全一致を優先し、なければ「名称が含まれる」部分一致で略称を探す */
function _findAbbr(list, text) {
  text = String(text || '').trim();
  if (!text) return null;
  for (const r of list) if (r.name === text) return r.ab;
  let best = null;
  for (const r of list) {
    if (text.indexOf(r.name) !== -1 && (!best || r.name.length > best.name.length)) best = r;
  }
  return best ? best.ab : null;
}

/* ============ 権限チェック(セキュリティの要) ============ */
function _me() {
  return (Session.getActiveUser().getEmail() || '').toLowerCase();
}

/**
 * 職員専用APIの入口で必ず呼ぶ。
 * google.script.run で公開される関数は、URLを開ける全員(生徒含む)が
 * 直接呼べるため、サーバー側での本人確認が唯一の防御。
 * ※新しい職員向け関数を追加するときは必ず先頭に入れること
 */
function _assertTeacher() {
  if (!_teacherSet()[_me()]) {
    throw new Error('この操作には職員権限が必要です');
  }
}

/* ============ 設定・参照 ============ */
function _settings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const vals = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0]) {
      let v = vals[i][1];
      if (typeof v === 'boolean') v = v ? 'TRUE' : 'FALSE';
      map[String(vals[i][0]).trim()] = String(v || '').trim();
    }
  }
  return map;
}

function _teacherSet() {
  const set = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEACHERS);
  if (sh) {
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      const a = String(vals[i][1] || '').trim().toLowerCase();
      if (a) set[a] = true;
    }
  }
  const raw = _settings()['職員アドレス'] || '';
  raw.split(',').forEach(a => { a = a.trim().toLowerCase(); if (a) set[a] = true; });
  return set;
}

function _rosterMaps() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  const vals = sh.getDataRange().getValues();
  const byId = {}, byEmail = {};
  for (let i = 1; i < vals.length; i++) {
    const [id, grade, cls, num, name, email] = vals[i];
    if (!id) continue;
    const rec = { id: String(id).trim(), grade, cls, num, name,
                  email: String(email || '').trim().toLowerCase() };
    byId[rec.id] = rec;
    if (rec.email) byEmail[rec.email] = rec;
  }
  return { byId, byEmail };
}

function _parseDueDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(String(v).trim());
  return isNaN(d.getTime()) ? null : d;
}

/* ============ デイリー教材モード(v6.5〜) 共通関数 ============ */

/** 教材シートの「毎日」列の値がONかどうか(boolean/文字列/数値すべて対応)。
 *  IME変換で人によって表記がぶれる「まる」記号(〇U+3007/○U+25CB/◯U+25EF)は
 *  すべて同じ意味として扱う(見た目は同じでも文字コードが違うため要注意)。
 *  チェックボックス列(受付終了)の値もbooleanで来るため、この関数をそのまま流用する。 */
function _isDailyFlag(v) {
  if (v === true || v === 1) return true;
  const s = String(v == null ? '' : v).trim().toUpperCase();
  if (s === 'TRUE' || s === '1' || s === 'YES') return true;
  return s === '○' || s === '〇' || s === '◯';
}

/** サーバー側の「今日」をyyyy-MM-dd形式で返す。デイリー教材の記録・判定はすべてこれを使う */
function _todayKey() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Dateまたは日付文字列をyyyy-MM-ddに正規化(ログ読み出し時に使用)。
 *  スプレッドシートに保存したyyyy-MM-dd文字列は、読み出し時にDate型へ
 *  自動変換されることがあるため、比較の前には必ずこれを通す。 */
function _dateKeyOf(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v).trim().slice(0, 10);
}

/** _latestMapからデイリー教材の指定日の記録を引くためのキー */
function _dailyKey(studentId, materialCode, dateKey) {
  return studentId + '|' + materialCode + '|' + dateKey;
}

/** _latestMapから「生徒×教材」の該当記録を引く共通ヘルパー。
 *  デイリー教材は自動的に「今日」の記録を引く(通常教材は従来どおり)。 */
function _latestFor(latest, matMap, studentId, code) {
  const m = matMap[code];
  if (m && m.daily) {
    return latest[_dailyKey(studentId, code, _todayKey())];
  }
  return latest[studentId + '|' + code];
}

/** 教材シートに「毎日」「受付終了」列、ログシートに「提出日」列がなければ追加する(冪等)。
 *  「受付終了」列はチェックボックスにする。再実行すると、後から追加された教材行にも
 *  チェックボックスが張り直される。 */
function setupDailyColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const added = [];
  [[SHEET_MATERIAL, '毎日'], [SHEET_LOG, '提出日'],
   [SHEET_MATERIAL, '受付終了']]
    .concat(MATERIAL_AUDIT_COLS.map(c => [SHEET_MATERIAL, c]))
    .forEach(([sheetName, colName]) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (headers.indexOf(colName) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(colName);
      added.push(sheetName + 'シートに「' + colName + '」列');
    }
  });
  // 「受付終了」列のデータ範囲にチェックボックスを設定(毎回張り直す=冪等。
  // 後から教材行が増えても、このメニューを再実行すれば行末まで反映される)
  const mat = ss.getSheetByName(SHEET_MATERIAL);
  if (mat) {
    const headers = mat.getRange(1, 1, 1, mat.getLastColumn()).getValues()[0].map(String);
    const closedCol = headers.indexOf('受付終了');
    if (closedCol !== -1 && mat.getLastRow() >= 2) {
      mat.getRange(2, closedCol + 1, mat.getLastRow() - 1, 1).insertCheckboxes();
    }
  }
  // 略称シートに推奨の教材名略称(種別「教材」)を、まだ無い名称の分だけ足す。
  // 学校で直した略称や消した行は、同じ名称が残っていれば上書きしない
  const abbrSh = ss.getSheetByName(SHEET_ABBR);
  if (abbrSh) {
    const have = {};
    _abbrMaps().material.forEach(r => { have[r.name] = true; });
    const rows = RECOMMENDED_MATERIAL_ABBRS.filter(([name]) => !have[name])
      .map(([name, ab]) => ['教材', name, ab]);
    if (rows.length) {
      abbrSh.getRange(abbrSh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
      added.push('略称シートに推奨の教材名略称' + rows.length + '件');
    }
  }
  // 設定シートに「表示名」の行を空欄で用意する(既にあれば値はそのまま)
  const setSh = ss.getSheetByName(SHEET_SETTINGS);
  if (setSh && !setSh.getDataRange().getValues().slice(1)
      .some(r => String(r[0]).trim() === '表示名')) {
    _setSettingsRow('表示名', '', false);
    added.push('設定シートに「表示名」の行(例: 〇〇中学校 R7入学生 と入力すると画面に表示)');
  }
  ui.alert(added.length ? added.join('、') + 'を追加しました。' : '列は準備済みです。');
}

function _materialMap() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MATERIAL);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const dailyCol = headers.indexOf('毎日'); // マイグレーション前は-1(=通常教材扱い)
  const closedCol = headers.indexOf('受付終了'); // マイグレーション前は-1(=受付中扱い)
  const vals = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++) {
    const [code, grade, targetCls, subject, name, matAbbr, range, due, labels, fails] = vals[i];
    if (!code) continue;
    const rangeStr = String(range || '').trim();
    map[String(code).trim()] = {
      grade: String(grade || '').trim(),
      grades: String(grade || '').split(',').map(x => x.trim()).filter(Boolean),
      clses:  String(targetCls || '').split(',').map(x => x.trim()).filter(Boolean),
      subject: subject, name: name, abbr: String(matAbbr || '').trim(), range: rangeStr,
      due: _parseDueDate(due),
      disp: String(name) + (rangeStr ? '(' + rangeStr + ')' : ''),
      labels: String(labels || '合格,再提出').split(',').map(s => s.trim()).filter(Boolean),
      fails:  String(fails  || '再提出').split(',').map(s => s.trim()).filter(Boolean),
      daily: dailyCol !== -1 ? _isDailyFlag(vals[i][dailyCol]) : false,
      closed: closedCol !== -1 ? _isDailyFlag(vals[i][closedCol]) : false,
    };
  }
  return map;
}

/* ============ ルーティング ============ */
function doGet(e) {
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) {
      const props = PropertiesService.getScriptProperties();
      const changed = props.getProperty('WEBAPP_URL') !== url;
      if (changed || _settingsRowIsEmpty('WebアプリURL')) {
        props.setProperty('WEBAPP_URL', url);
        _writeWebAppUrlToSettings(url);
      }
    }
  } catch (err) { /* 記憶・反映できなくても表示処理は続行 */ }

  try {
    if (_settingsRowIsEmpty('バージョン') ||
        _settingsRowValue('バージョン') !== APP_VERSION + (IS_DEMO_BUILD ? '(体験版)' : '')) {
      _setSettingsRow('バージョン', APP_VERSION + (IS_DEMO_BUILD ? '(体験版)' : ''), false);
    }
  } catch (err) { /* 反映できなくても表示処理は続行 */ }

  const email = _me();
  const page = (e && e.parameter && e.parameter.page) || 'auto';
  const isTeacher = !!_teacherSet()[email];
  const student = _rosterMaps().byEmail[email] || null;

  let name = '';
  try { name = getDisplayName(); } catch (err) { /* 読めなくても表示は続行 */ }
  const serve = f => HtmlService.createHtmlOutputFromFile(f)
    .setTitle('提出物チェック' + (IS_DEV_BUILD ? '(開発版)' : '') + (name ? '|' + name : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');

  if (page === 'print')   return isTeacher ? serve('print')   : _deny(email);
  if (page === 'scan')    return isTeacher ? serve('index')   : _deny(email);
  if (page === 'student') return (student || isTeacher) ? serve('student') : _deny(email);

  if (isTeacher) return serve('index');
  if (student)   return serve('student');
  return _deny(email);
}

function _deny(email) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:24px">' +
    '<h2>アクセス権がありません</h2>' +
    '<p>このアカウント(' + (email || '不明') + ')は名簿・職員リストに登録されていません。<br>' +
    '担当の先生に連絡してください。</p></div>');
}

/** 設定シート「表示名」(例: 〇〇中学校 R7入学生)。学年ごとにファイルを分けて運用するとき、
 *  どのファイルの画面かわかるようにタブ・画面上部に出す。空欄なら何も出さない。
 *  校名はコードに書かず、ここで設定する(配布テンプレートでもそのまま使えるように)。
 *  認証不要(画面の見出しに使うだけの read-only) */
function getDisplayName() {
  return String(_settings()['表示名'] || '').trim();
}

/** 画面表示用のバージョン文字列。認証不要(harmless read-only) */
function getAppVersion() {
  return APP_VERSION + (IS_DEMO_BUILD ? '(体験版)' : IS_DEV_BUILD ? '(開発版)' : '');
}

/** 画面上部に警告バナーを出すべきか。開発版のときだけ文言を返し、
 *  本番・体験版(体験版は別途固定バナーを持つ)ではnull。 */
function getBuildBanner() {
  if (IS_DEV_BUILD) {
    return '🔧 これは開発版です。本番データのコピーで動作確認中。生徒・保護者には案内しないでください';
  }
  return null;
}

/* ============ スキャナ用API(職員専用) ============ */
function lookupMaterial(code) {
  _assertTeacher();
  const m = _materialMap()[String(code).trim()];
  if (!m) return null;
  return { code: String(code).trim(), subject: m.subject,
           name: m.disp, labels: m.labels, fails: m.fails, daily: !!m.daily };
}

function recordScan(studentId, materialCode, label) {
  _assertTeacher();
  studentId = String(studentId).trim();
  materialCode = String(materialCode).trim();
  label = String(label).trim();

  const student = _rosterMaps().byId[studentId];
  if (!student) return { ok: false, error: '名簿にない生徒IDです: ' + studentId };
  const material = _materialMap()[materialCode];
  if (!material) return { ok: false, error: '教材コードが不正です' };
  if (material.labels.indexOf(label) === -1) {
    return { ok: false, error: 'この教材で使えない評価です: ' + label };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
    const lastCol = Math.max(sh.getLastColumn(), 8);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const submitCol = headers.indexOf('提出日'); // マイグレーション前は-1
    const vals = sh.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][3]).trim() === studentId &&
          String(vals[i][1]).trim() === materialCode &&
          vals[i][7] !== '取消') count++;
    }
    const nth = count + 1;
    const row = new Array(lastCol).fill('');
    row[0] = new Date(); row[1] = materialCode; row[2] = material.disp;
    row[3] = studentId; row[4] = student.name; row[5] = label;
    row[6] = nth; row[7] = '';
    if (submitCol !== -1 && material.daily) row[submitCol] = _todayKey();
    sh.appendRow(row);
    return {
      ok: true, row: sh.getLastRow(),
      name: student.name,
      cls: student.grade + '年' + student.cls + '組' + student.num + '番',
      label: label, nth: nth,
      isFail: material.fails.indexOf(label) !== -1,
    };
  } finally {
    lock.releaseLock();
  }
}

/** UNDO: 取消フラグを立てる */
function cancelRow(row) {
  _assertTeacher();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (row < 2 || row > sh.getLastRow()) return { ok: false, error: '行が不正です' };
  sh.getRange(row, 8).setValue('取消');
  sh.getRange(row, 1, 1, 8).setFontColor('#999999').setFontLine('line-through');
  return { ok: true,
    name: sh.getRange(row, 5).getValue(),
    judge: sh.getRange(row, 6).getValue() };
}

/** REDO: 取消フラグを外して記録を復活させる */
function restoreRow(row) {
  _assertTeacher();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (row < 2 || row > sh.getLastRow()) return { ok: false, error: '行が不正です' };
  sh.getRange(row, 8).setValue('');
  sh.getRange(row, 1, 1, 8).setFontColor('#000000').setFontLine('none');
  return { ok: true,
    name: sh.getRange(row, 5).getValue(),
    judge: sh.getRange(row, 6).getValue() };
}

/* ============ 最新状態の共通ロジック ============ */
/** ログシートから「生徒×教材」(デイリー教材は「生徒×教材×提出日」)ごとの
 *  最新の有効ログを1件ずつ抽出する。matMapを渡さない場合は内部で取得する
 *  (呼び出し側で既に_materialMap()を取得済みなら渡して二重読み込みを避ける)。 */
function _latestMap(matMap) {
  matMap = matMap || _materialMap();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  const lastCol = Math.max(sh.getLastColumn(), 8);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const submitCol = headers.indexOf('提出日'); // マイグレーション前は-1
  const vals = sh.getDataRange().getValues();
  const latest = {};
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0] || r[7] === '取消') continue;
    const materialCode = String(r[1]).trim();
    const m = matMap[materialCode];
    let key = String(r[3]).trim() + '|' + materialCode;
    if (m && m.daily) {
      const dateVal = submitCol !== -1 ? r[submitCol] : '';
      key += '|' + _dateKeyOf(dateVal);
    }
    latest[key] = { label: String(r[5]), nth: r[6], date: r[0], row: i + 1 };
  }
  return latest;
}

function _daysUntil(due) {
  if (!due) return null;
  const now = new Date();
  const d0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const d1 = new Date(due.getFullYear(), due.getMonth(), due.getDate());
  return Math.round((d1 - d0) / 86400000);
}

/** 締切での絞り込み(v6.8〜、未提出リスト・教材タブ共通の考え方)。
 *  mode: ''=すべて / has=締切あり / overdue=締切を過ぎた / week=7日以内(今日〜7日後) /
 *        range=期間指定(from〜to、yyyy-MM-dd、片方だけでも可) / none=締切なし
 *  毎日の教材は締切日を持たない扱い(none 側) */
const DUE_MODES = ['', 'has', 'overdue', 'week', 'range', 'none'];
function _dueMatches(m, mode, from, to) {
  if (!mode) return true;
  const due = m.daily ? null : m.due;
  if (mode === 'none') return !due;
  if (!due) return false;
  if (mode === 'has') return true;
  const left = _daysUntil(due);
  if (mode === 'overdue') return left < 0;
  if (mode === 'week') return left >= 0 && left <= 7;
  if (mode === 'range') {
    const k = _dateKeyOf(due);
    return (!from || k >= from) && (!to || k <= to);
  }
  return true;
}

function _urgencyOf(daysLeft) {
  if (daysLeft === null) return 'none';
  if (daysLeft < 0) return 'overdue';
  if (daysLeft === 0) return 'today';
  if (daysLeft <= 7) return 'soon';
  return 'none';
}

function _badgeText(daysLeft, urgency) {
  if (urgency === 'overdue') return '⚠期限超過(' + Math.abs(daysLeft) + '日経過)';
  if (urgency === 'today') return '⏰本日締切';
  if (urgency === 'soon') return '⏰あと' + daysLeft + '日';
  return '';
}

function _requireStudent() {
  const email = _me();
  const maps = _rosterMaps();
  let student = maps.byEmail[email];
  if (!student && _teacherSet()[email]) {
    const ids = Object.keys(maps.byId);
    if (ids.length) student = maps.byId[ids[0]];
  }
  if (!student) throw new Error('生徒として認識できませんでした');
  return student;
}

/* ============ 生徒本人ページ用API(本人のデータのみ) ============ */
function getMyStatus() {
  const email = _me();
  const maps = _rosterMaps();
  let student = maps.byEmail[email];
  if (!student && _teacherSet()[email]) {
    const ids = Object.keys(maps.byId);
    if (ids.length) student = maps.byId[ids[0]];
  }
  if (!student) return { ok: false, error: 'アカウントが名簿に登録されていません' };

  const materials = _materialMap();
  const latest = _latestMap(materials);
  const items = [];
  let resubmit = 0;

  Object.keys(materials).forEach(code => {
    const m = materials[code];
    const s = _latestFor(latest, materials, student.id, code);
    const gOk = !m.grades.length || m.grades.indexOf(String(student.grade)) !== -1;
    const cOk = !m.clses.length  || m.clses.indexOf(String(student.cls))  !== -1;
    if (!(gOk && cOk) && !s) return;

    let status = '未';
    if (s) {
      if (s.label === GIVEUP_LABEL) status = '断念';
      else status = (m.fails.indexOf(s.label) !== -1) ? '再提出' : '済';
    }
    // 受付終了教材: 済(合格)以外はすべて「受付終了」に置き換える
    // (未提出・再提出・断念はもう督促しない。済は従来どおり)
    if (m.closed && status !== '済') status = '受付終了';
    if (status === '再提出') resubmit++;   // ← 受付終了に置き換え済みならカウントされない

    // デイリー教材は締切バッジの代わりに「毎日」バッジを出す(今日出したかの表示)
    const daysLeft = (status === '未' && !m.daily) ? _daysUntil(m.due) : null;
    const urgency = (status === '未') ? (m.daily ? 'today' : _urgencyOf(daysLeft)) : 'none';
    const badge = (status === '未' && m.daily) ? '📅毎日提出' : _badgeText(daysLeft, urgency);

    items.push({
      kind: 'material',
      materialCode: code,
      subject: m.subject, name: m.disp,
      daily: !!m.daily,
      label: s ? s.label : '', nth: s ? s.nth : 0,
      date: s ? Utilities.formatDate(new Date(s.date), 'JST', 'M/d') : '',
      status: status,
      dueText: (m.daily || !m.due) ? '' : Utilities.formatDate(m.due, 'JST', 'M/d'),
      daysLeft: daysLeft, urgency: urgency, badge: badge,
      row: (s && status !== '再提出' && status !== '済' && status !== '受付終了') ? s.row : null,
    });
  });

  _myTasksRaw(student.id).forEach(t => {
    const daysLeft = t.done ? null : _daysUntil(t.due);
    const urgency = t.done ? 'none' : _urgencyOf(daysLeft);
    items.push({
      kind: 'custom', row: t.row,
      subject: '自分の課題', name: t.text,
      status: t.done ? '完了' : '未完了',
      dueText: t.due ? Utilities.formatDate(t.due, 'JST', 'M/d') : '',
      daysLeft: daysLeft, urgency: urgency, badge: _badgeText(daysLeft, urgency),
    });
  });

  const bucketOf = it => {
    if (it.status === '受付終了') return 5;             // ← 追加(最下位)
    if ((it.status === '未' || it.status === '未完了') && it.urgency === 'overdue') return -1;
    if (it.status === '再提出') return 0;
    if (it.status === '未' || it.status === '未完了') {
      return (it.urgency === 'today' || it.urgency === 'soon') ? 1 : 2;
    }
    if (it.status === '済' || it.status === '完了') return 3;
    return 4;
  };
  items.forEach(it => { it.bucket = bucketOf(it); });
  items.sort((a, b) => {
    if (a.bucket !== b.bucket) return a.bucket - b.bucket;
    if (a.bucket === 1) {
      const da = a.daysLeft === null ? 9999 : a.daysLeft;
      const db = b.daysLeft === null ? 9999 : b.daysLeft;
      if (da !== db) return da - db;
    }
    return String(a.subject).localeCompare(String(b.subject), 'ja');
  });

  return { ok: true,
    name: student.name,
    cls: student.grade + '年' + student.cls + '組' + student.num + '番',
    items: items, resubmit: resubmit };
}

/* ============ 生徒本人による「提出をあきらめる」/取消 ============ */
function giveUpSubmission(materialCode) {
  const student = _requireStudent();
  materialCode = String(materialCode).trim();
  const material = _materialMap()[materialCode];
  if (!material) return { ok: false, error: '教材コードが不正です' };
  if (material.daily) {
    return { ok: false, error: 'この教材(毎日提出)は「あきらめる」の対象外です' };
  }
  if (material.closed) {
    return { ok: false, error: 'この教材は受付を終了しています' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
    const vals = sh.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][3]).trim() === student.id &&
          String(vals[i][1]).trim() === materialCode &&
          vals[i][7] !== '取消') count++;
    }
    const nth = count + 1;
    sh.appendRow([new Date(), materialCode, material.disp, student.id,
                  student.name, GIVEUP_LABEL, nth, '']);
    return { ok: true, row: sh.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function cancelMyGiveUp(row) {
  const student = _requireStudent();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (row < 2 || row > sh.getLastRow()) return { ok: false, error: '行が不正です' };
  const rowId = String(sh.getRange(row, 4).getValue()).trim();
  const rowLabel = String(sh.getRange(row, 6).getValue()).trim();
  if (rowId !== student.id || rowLabel !== GIVEUP_LABEL) {
    return { ok: false, error: 'この記録は取り消せません' };
  }
  sh.getRange(row, 8).setValue('取消');
  sh.getRange(row, 1, 1, 8).setFontColor('#999999').setFontLine('line-through');
  return { ok: true };
}

/* ============ 生徒本人の自作タスク ============ */
function _tasksSheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('TASKS_SS_ID');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('提出物チェック_個人タスク(共有しないこと)');
    props.setProperty('TASKS_SS_ID', ss.getId());
    const sh0 = ss.getSheets()[0];
    sh0.setName(SHEET_TASKS);
    sh0.getRange(1, 1, 1, 5)
      .setValues([['生徒ID', '内容', '期限(空欄可)', '完了', '作成日時']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh0.setFrozenRows(1);
  }
  let sh = ss.getSheetByName(SHEET_TASKS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_TASKS);
    sh.getRange(1, 1, 1, 5)
      .setValues([['生徒ID', '内容', '期限(空欄可)', '完了', '作成日時']])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  }
  return sh;
}

function _myTasksRaw(studentId) {
  const sh = _tasksSheet();
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0] || String(r[0]).trim() !== studentId) continue;
    out.push({
      row: i + 1,
      text: String(r[1] || ''),
      due: _parseDueDate(r[2]),
      done: String(r[3]).trim() === '完了',
    });
  }
  return out;
}

function addMyTask(text, dueDateStr) {
  const student = _requireStudent();
  text = String(text || '').trim();
  if (!text) return { ok: false, error: '内容を入力してください' };
  if (text.length > 100) return { ok: false, error: '内容が長すぎます(100文字以内)' };
  const due = dueDateStr ? _parseDueDate(dueDateStr) : null;
  const sh = _tasksSheet();
  sh.appendRow([student.id, text, due || '', '', new Date()]);
  return { ok: true };
}

function toggleMyTask(row) {
  const student = _requireStudent();
  const sh = _tasksSheet();
  if (row < 2 || row > sh.getLastRow()) return { ok: false, error: '行が不正です' };
  const owner = String(sh.getRange(row, 1).getValue()).trim();
  if (owner !== student.id) return { ok: false, error: '他の生徒のタスクは操作できません' };
  const cur = String(sh.getRange(row, 4).getValue()).trim();
  sh.getRange(row, 4).setValue(cur === '完了' ? '' : '完了');
  return { ok: true };
}

function deleteMyTask(row) {
  const student = _requireStudent();
  const sh = _tasksSheet();
  if (row < 2 || row > sh.getLastRow()) return { ok: false, error: '行が不正です' };
  const owner = String(sh.getRange(row, 1).getValue()).trim();
  if (owner !== student.id) return { ok: false, error: '他の生徒のタスクは操作できません' };
  sh.deleteRow(row);
  return { ok: true };
}

/* ============ 集計 ============ */
function _summaryMatCodesSorted(matMap) {
  matMap = matMap || _materialMap();
  return Object.keys(matMap).sort((a, b) => {
    const ma = matMap[a], mb = matMap[b];
    return String(ma.subject).localeCompare(String(mb.subject), 'ja') ||
           String(ma.name).localeCompare(String(mb.name), 'ja') ||
           String(ma.range).localeCompare(String(mb.range), 'ja', { numeric: true });
  });
}

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roster = ss.getSheetByName(SHEET_ROSTER).getDataRange().getValues()
    .slice(1).filter(r => r[0]);
  const matMap = _materialMap();
  const matCodes = _summaryMatCodesSorted(matMap);
  const latest = _latestMap(matMap);
  const FIXED = 4;

  const sh = ss.getSheetByName(SHEET_SUMMARY);
  const oldFilter = sh.getFilter();
  if (oldFilter) oldFilter.remove();
  if (sh.getMaxColumns() > FIXED) {
    sh.getRange(1, FIXED + 1, 1, sh.getMaxColumns() - FIXED).shiftColumnGroupDepth(-8);
  }
  if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  }
  sh.clear();

  const rowSubject = ['', '', '', ''].concat(matCodes.map(c => matMap[c].subject));
  const rowName    = ['', '', '', ''].concat(matCodes.map(c => matMap[c].name));
  const rowRange   = ['学年', '組', '番号', '氏名'].concat(matCodes.map(c => matMap[c].range || ''));
  const values = [rowSubject, rowName, rowRange];
  const colors = [rowSubject.map(() => '#e8eaed'), rowName.map(() => '#e8eaed'),
                   rowRange.map(() => '#e8eaed')];

  roster.forEach(st => {
    const row = [st[1], st[2], st[3], st[4]];
    const crow = ['', '', '', ''];
    matCodes.forEach(code => {
      const s = _latestFor(latest, matMap, String(st[0]).trim(), code);
      if (!s) { row.push('未'); crow.push('#fff2cc'); return; }
      if (s.label === GIVEUP_LABEL) { row.push('断念'); crow.push('#e0e0e0'); return; }
      const isFail = matMap[code].fails.indexOf(s.label) !== -1;
      row.push(s.label + '(' + s.nth + ')');
      crow.push(isFail ? '#f4cccc' : '#d9ead3');
    });
    values.push(row);
    colors.push(crow);
  });

  const header = rowSubject;
  sh.getRange(1, 1, values.length, header.length).setValues(values)
    .setBackgrounds(colors);
  sh.getRange(1, 1, 3, header.length).setFontWeight('bold');
  sh.setFrozenRows(3);
  sh.setFrozenColumns(FIXED);

  let i = 0;
  while (i < matCodes.length) {
    const subject = matMap[matCodes[i]].subject;
    let j = i;
    while (j < matCodes.length && matMap[matCodes[j]].subject === subject) j++;
    if (j - i > 1) sh.getRange(1, FIXED + i + 1, 1, j - i).merge();

    let k = i;
    while (k < j) {
      const name = matMap[matCodes[k]].name;
      let l = k;
      while (l < j && matMap[matCodes[l]].name === name) l++;
      if (l - k > 1) sh.getRange(2, FIXED + k + 1, 1, l - k).merge();
      k = l;
    }
    i = j;
  }

  sh.getRange(3, 1, 1 + roster.length, header.length).createFilter();

  return { ok: true };
}

/* ============ デイリー集計(v6.5〜) ============ */
/** デイリー教材(教材シートの「毎日」列がON)を対象に、日ごとの履歴グリッド
 *  (行=生徒、列=日付、値=評価ラベル)を「デイリー集計」シートへ出力する。
 *  実行時にui.promptで対象月(yyyy-MM)を尋ねる。空欄なら今月。
 *  教材ごとに1ブロックを縦に並べ、実行のたびにシート全体を書き換える。 */
function buildDailySummary() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('デイリー集計を更新',
    '対象月をyyyy-MM形式で入力してください(例: 2026-07)。空欄のままOKで今月。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const input = res.getResponseText().trim();
  const now = new Date();
  let year, month; // month: 1-12
  if (!input) {
    year = now.getFullYear(); month = now.getMonth() + 1;
  } else {
    const m = input.match(/^(\d{4})-(\d{1,2})$/);
    if (!m) { ui.alert('「2026-07」のような形式で入力してください: ' + input); return; }
    year = parseInt(m[1], 10); month = parseInt(m[2], 10);
    if (month < 1 || month > 12) { ui.alert('月は1〜12で入力してください: ' + input); return; }
  }
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthLabel = year + '-' + ('0' + month).slice(-2);
  // 各日の曜日を先に計算しておく(ヘッダー表記・土日グレー表示に使う)
  const WEEKDAY_CHARS = ['日', '月', '火', '水', '木', '金', '土'];
  const dayInfos = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    dayInfos.push({ day: d, dow: dow, isWeekend: (dow === 0 || dow === 6) });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 月ごとに別シートにする(例:「デイリー集計_2026-07」)。同じシートを使い回すと
  // 月が変わるたびに前の月のデータが消えてしまうため、月替わりでも過去分が残るように
  // シート名に対象月を含める。同じ月を再実行した場合はそのシートだけが更新される。
  const dailySheetName = SHEET_DAILY_SUMMARY + '_' + monthLabel;
  let sh = ss.getSheetByName(dailySheetName);
  if (!sh) sh = ss.insertSheet(dailySheetName);
  const oldFilter = sh.getFilter();
  if (oldFilter) oldFilter.remove();
  if (sh.getMaxRows() > 0 && sh.getMaxColumns() > 0) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  }
  sh.clear();

  const matMap = _materialMap();
  const dailyCodes = _summaryMatCodesSorted(matMap).filter(c => matMap[c].daily);

  if (!dailyCodes.length) {
    sh.getRange(1, 1).setValue(
      'デイリー教材(教材シートの「毎日」列がON)が見つかりません。');
    ui.alert('デイリー教材が見つかりませんでした。教材シートの「毎日」列を確認してください。');
    return;
  }

  const roster = ss.getSheetByName(SHEET_ROSTER).getDataRange().getValues()
    .slice(1).filter(r => r[0]);
  const latest = _latestMap(matMap);
  const FIXED = 4; // 学年・組・番号・氏名
  let curRow = 1;

  dailyCodes.forEach(code => {
    const m = matMap[code];
    const students = roster.filter(st => {
      const grade = String(st[1]), cls = String(st[2]);
      const gOk = !m.grades.length || m.grades.indexOf(grade) !== -1;
      const cOk = !m.clses.length  || m.clses.indexOf(cls)  !== -1;
      return gOk && cOk;
    });
    // 1行目: 教材名 + 対象月
    // ※セル結合(merge)は、列固定やシートの状態によって一部の列だけ結合が
    //   効かなくなる不具合が起きたため使わない。代わりに行全体の背景色だけを
    //   濃い色で揃え、文字は先頭セル1つだけに入れて右側の空セルへ自然に
    //   はみ出させることで、結合と同じ見た目を壊れずに再現する。
    const totalCols = FIXED + daysInMonth + 1;
    sh.getRange(curRow, 1, 1, totalCols).setBackground('#37474f');
    sh.getRange(curRow, FIXED + 1, 1, 1)
      .setValue(m.subject + '/' + m.name + (m.range ? '(' + m.range + ')' : '') +
                ' - ' + monthLabel + '分')
      .setFontWeight('bold').setFontColor('#ffffff');
    curRow++;

    // 2行目: ヘッダー(学年 | 組 | 番号 | 氏名 | 1(月) | 2(火) | … | 回数)。土日はグレー背景
    const header = ['学年', '組', '番号', '氏名'];
    const headerColors = ['#e8eaed', '#e8eaed', '#e8eaed', '#e8eaed'];
    dayInfos.forEach(info => {
      header.push(info.day + '(' + WEEKDAY_CHARS[info.dow] + ')');
      headerColors.push(info.isWeekend ? '#cfd8dc' : '#e8eaed');
    });
    header.push('回数');
    headerColors.push('#e8eaed');
    sh.getRange(curRow, 1, 1, header.length).setValues([header])
      .setFontWeight('bold');
    sh.getRange(curRow, 1, 1, headerColors.length).setBackgrounds([headerColors]);
    curRow++;

    // 3行目以降: 生徒ごとの日別記録(未提出は空欄、再提出扱いは赤系、それ以外は緑系)
    students.forEach(st => {
      const studentId = String(st[0]).trim();
      const row = [st[1], st[2], st[3], st[4]];
      const colors = ['#ffffff', '#ffffff', '#ffffff', '#ffffff'];
      let cnt = 0;
      dayInfos.forEach(info => {
        const dateKey = monthLabel + '-' + ('0' + info.day).slice(-2);
        const s = latest[_dailyKey(studentId, code, dateKey)];
        if (!s) { row.push(''); colors.push(info.isWeekend ? '#eeeeee' : '#ffffff'); return; }
        row.push(s.label);
        const isFail = m.fails.indexOf(s.label) !== -1;
        colors.push(isFail ? '#f4cccc' : '#d9ead3');
        cnt++;
      });
      row.push(cnt);
      colors.push('#ffffff');
      sh.getRange(curRow, 1, 1, row.length).setValues([row]).setBackgrounds([colors]);
      curRow++;
    });

    curRow += 1; // ブロック間の空行
  });

  sh.setFrozenColumns(FIXED);
  ui.alert('シート「' + dailySheetName + '」を更新しました。');
}

/* ============ 集計シートの列絞り込み(教材名・教科) ============ */
const SUMMARY_FIXED_COLS = 4;

function showMaterialNamePicker() {
  const matMap = _materialMap();
  const set = {};
  Object.keys(matMap).forEach(c => { set[matMap[c].name] = true; });
  const names = Object.keys(set).sort((a, b) => a.localeCompare(b, 'ja'));
  const chosen = _promptPickFromList('教材名で絞り込み', '教材名を選んでください', names);
  if (chosen) _filterSummaryColumns(m => m.name === chosen);
}

function showSubjectPicker() {
  const matMap = _materialMap();
  const set = {};
  Object.keys(matMap).forEach(c => { set[matMap[c].subject] = true; });
  const subjects = Object.keys(set).sort((a, b) => a.localeCompare(b, 'ja'));
  const chosen = _promptPickFromList('教科で絞り込み', '教科を選んでください', subjects);
  if (chosen) _filterSummaryColumns(m => m.subject === chosen);
}

function showSerialRangePicker() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('連番で絞り込み(教材コード末尾2桁)',
    '例: 1-5 と入力すると、末尾の連番が01〜05の教材だけを表示します。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const input = res.getResponseText().trim();
  const m = input.match(/^(\d{1,2})\s*[-〜~]\s*(\d{1,2})$/);
  if (!m) { ui.alert('「1-5」のような形式で入力してください: ' + input); return; }
  const lo = parseInt(m[1], 10), hi = parseInt(m[2], 10);
  if (lo > hi) { ui.alert('開始番号が終了番号より大きくなっています'); return; }
  _filterSummaryColumns((mat, code) => {
    const mm = code.match(/-(\d+)$/);
    if (!mm) return false;
    const n = parseInt(mm[1], 10);
    return n >= lo && n <= hi;
  });
}

function _promptPickFromList(title, label, options, allowBlank) {
  const ui = SpreadsheetApp.getUi();
  if (!options.length) {
    if (allowBlank) return null;
    ui.alert('候補がありません(教材シートを確認してください)');
    return undefined;
  }
  const list = options.map((v, i) => (i + 1) + '. ' + v).join('\n');
  const res = ui.prompt(title,
    label + '(番号か名前を入力' + (allowBlank ? '。空欄のままOKで「すべて」' : '') + ')\n\n' + list,
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return undefined;
  const input = res.getResponseText().trim();
  if (!input) {
    if (allowBlank) return null;
    ui.alert('入力してください');
    return undefined;
  }
  const asNum = parseInt(input, 10);
  if (String(asNum) === input && asNum >= 1 && asNum <= options.length) {
    return options[asNum - 1];
  }
  const hit = options.find(o => o === input) ||
              options.find(o => o.toLowerCase() === input.toLowerCase());
  if (!hit) { ui.alert('一致する候補が見つかりません: ' + input); return undefined; }
  return hit;
}

function _uniqueSortedJa(arr) {
  const set = {};
  arr.forEach(v => { set[v] = true; });
  return Object.keys(set).sort((a, b) => a.localeCompare(b, 'ja'));
}

function showCombinedPicker() {
  const matMap = _materialMap();

  const subjects = _uniqueSortedJa(Object.keys(matMap).map(c => matMap[c].subject));
  const subject = _promptPickFromList('①教科で絞り込み(任意)', '教科を選んでください',
    subjects, true);
  if (subject === undefined) return;

  const namePool = Object.keys(matMap)
    .filter(c => !subject || matMap[c].subject === subject)
    .map(c => matMap[c].name);
  const names = _uniqueSortedJa(namePool);
  const name = _promptPickFromList('②教材名で絞り込み(任意)', '教材名を選んでください',
    names, true);
  if (name === undefined) return;

  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('③連番で絞り込み(任意)',
    '教材コード末尾2桁の範囲。例: 1-5 。空欄のままOKで「すべて」',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const rangeInput = res.getResponseText().trim();
  let lo = null, hi = null;
  if (rangeInput) {
    const m = rangeInput.match(/^(\d{1,2})\s*[-〜~]\s*(\d{1,2})$/);
    if (!m) { ui.alert('「1-5」のような形式で入力してください: ' + rangeInput); return; }
    lo = parseInt(m[1], 10); hi = parseInt(m[2], 10);
    if (lo > hi) { ui.alert('開始番号が終了番号より大きくなっています'); return; }
  }

  _filterSummaryColumns((mat, code) => {
    if (subject && mat.subject !== subject) return false;
    if (name && mat.name !== name) return false;
    if (lo !== null) {
      const mm = code.match(/-(\d+)$/);
      const n = mm ? parseInt(mm[1], 10) : null;
      if (n === null || n < lo || n > hi) return false;
    }
    return true;
  });
}

function _filterSummaryColumns(predicate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_SUMMARY);
  const matMap = _materialMap();
  const matCodes = _summaryMatCodesSorted(matMap);
  if (sh.getLastColumn() !== SUMMARY_FIXED_COLS + matCodes.length) {
    SpreadsheetApp.getUi().alert(
      '集計シートの列数が教材シートの内容と一致していません。\n' +
      '先にメニュー「④ 集計シートを更新」を実行してから、もう一度お試しください。');
    return;
  }
  matCodes.forEach((code, i) => {
    const col = SUMMARY_FIXED_COLS + i + 1;
    if (predicate(matMap[code], code)) sh.showColumns(col); else sh.hideColumns(col);
  });
}

function resetSummaryColumns() {
  _filterSummaryColumns(() => true);
}

/* ============ Classroom通知(職員専用) ============ */
function getResubmitList(materialCode) {
  _assertTeacher();
  materialCode = String(materialCode).trim();
  const matMap = _materialMap();
  const m = matMap[materialCode];
  if (!m) return { ok: false, error: '教材コードが不正です' };
  if (m.closed) return { ok: false, error: 'この教材は受付を終了しています(通知対象外)' };
  const maps = _rosterMaps();
  const latest = _latestMap(matMap);
  const targets = [];
  Object.keys(maps.byId).forEach(id => {
    const s = _latestFor(latest, matMap, id, materialCode);
    if (s && m.fails.indexOf(s.label) !== -1) {
      const st = maps.byId[id];
      targets.push({ id: id, name: st.name, email: st.email,
                     cls: st.grade + '-' + st.cls + '-' + st.num });
    }
  });
  return { ok: true, material: m.disp, targets: targets,
           daily: !!m.daily, date: m.daily ? _todayKey() : null };
}

function sendClassroomNotice(materialCode) {
  _assertTeacher();
  const courseId = _settings()['ClassroomコースID'];
  if (!courseId) return { ok: false, error: '設定シートにClassroomコースIDがありません' };
  if (typeof Classroom === 'undefined') {
    return { ok: false,
      error: 'Classroom APIが未有効です。Apps Scriptエディタ左の「サービス」→「Google Classroom API」を追加してください' };
  }
  const list = getResubmitList(materialCode);
  if (!list.ok) return list;
  if (!list.targets.length) return { ok: false, error: '再提出者はいません' };

  const emails = {};
  list.targets.forEach(t => { if (t.email) emails[t.email] = true; });
  const ids = [];
  let pageToken = null;
  do {
    const res = Classroom.Courses.Students.list(courseId,
      { pageSize: 100, pageToken: pageToken });
    (res.students || []).forEach(s => {
      const em = (s.profile && s.profile.emailAddress || '').toLowerCase();
      if (emails[em]) ids.push(s.userId);
    });
    pageToken = res.nextPageToken;
  } while (pageToken);

  if (!ids.length) {
    return { ok: false,
      error: '対象生徒がこのClassroomコースに見つかりません(コースID・生徒の参加状況を確認)' };
  }

  const dateLabel = list.daily
    ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d') + 'の'
    : '';
  Classroom.Courses.Announcements.create({
    text: '【再提出のお知らせ】' + dateLabel + '「' + list.material + '」を再提出してください。' +
          '詳しくは提出物チェックページで確認できます。',
    state: 'PUBLISHED',
    assigneeMode: 'INDIVIDUAL_STUDENTS',
    individualStudentsOptions: { studentIds: ids },
  }, courseId);

  return { ok: true, sent: ids.length, material: list.material };
}

function menuNotifyClassroom() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Classroom通知', '教材コードを入力してください(例: JPN-WK-01)',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const code = res.getResponseText();
  const pre = getResubmitList(code);
  if (!pre.ok) { ui.alert(pre.error); return; }
  if (!pre.targets.length) { ui.alert('再提出者はいません'); return; }
  const go = ui.alert('確認',
    '「' + pre.material + '」の再提出者 ' + pre.targets.length + '名に通知します。よろしいですか?\n' +
    pre.targets.map(t => t.cls + ' ' + t.name).join('\n'),
    ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;
  const out = sendClassroomNotice(code);
  ui.alert(out.ok ? out.sent + '名に通知しました' : 'エラー: ' + out.error);
}

/* ============ 開発版(β環境)連携 ============ */
/**
 * 開発者自身が新機能を試す「開発版」スプレッドシートへ、本番の
 * 生徒名簿・教材・略称シートをコピーする(実データに近い状態で
 * 動作確認するため)。ログ・設定・個人タスクはコピーしない
 * (ログは量が多く不要、設定は開発版側で別管理したいため)。
 * 事前準備: 本番側のスクリプトプロパティ DEV_SS_ID に、
 * 開発版スプレッドシートのID(URLの/d/と/editの間の文字列)を設定しておくこと。
 */
function copyDataToDevBuild() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const devId = props.getProperty('DEV_SS_ID');
  if (!devId) {
    ui.alert(
      'スクリプトプロパティ「DEV_SS_ID」が未設定です。\n' +
      '(プロジェクトの設定 → スクリプトプロパティ) で、' +
      '開発版スプレッドシートのIDを設定してから、もう一度実行してください。');
    return;
  }
  const go = ui.alert('確認',
    '本番の「生徒名簿」「教材」「略称」シートを、開発版スプレッドシートへ' +
    'コピーします(開発版側の同名シートは上書きされます。本番データは変更されません)。' +
    'よろしいですか?',
    ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;

  let devSs;
  try {
    devSs = SpreadsheetApp.openById(devId);
  } catch (e) {
    ui.alert('開発版スプレッドシートを開けませんでした(IDを確認してください): ' + e.message);
    return;
  }

  const srcSs = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_ROSTER, SHEET_MATERIAL, SHEET_ABBR].forEach(name => {
    const src = srcSs.getSheetByName(name);
    if (!src) return;
    const old = devSs.getSheetByName(name);
    const copied = src.copyTo(devSs);
    if (old) devSs.deleteSheet(old);
    copied.setName(name);
  });

  ui.alert('コピーが完了しました。開発版スプレッドシートを開いて確認してください。\n' +
    '(生徒の記録状況を試したい場合は、開発版側でメニュー「④集計シートを更新」も' +
    '実行してください)');
}

/* ============ QR印刷用データ(職員専用) ============ */
function getPrintData() {
  _assertTeacher();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = ss.getSheetByName(SHEET_ROSTER).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      num4: String(r[1]) + String(r[2]) + ('0' + String(r[3])).slice(-2),
      grade: String(r[1]), cls: String(r[2]), num: String(r[3]),
      gakkyu: r[1] + '年' + r[2] + '組' + r[3] + '番',
      name: String(r[4] || ''),
    }));
  const matMap = _materialMap();
  const materials = Object.keys(matMap).map(c => ({
    id: c,
    grade: matMap[c].grade,
    cls: matMap[c].clses.join(','),
    subject: String(matMap[c].subject || ''),
    name: String(matMap[c].name || ''),
    range: matMap[c].range,
  }));
  return { students, materials };
}

/* ============ 未提出・要再提出リスト(職員専用) ============ */

/** 未提出・再提出リストの絞り込み条件を、教員ごとに記憶するためのキー接頭辞。
 *  ※UserPropertiesは使わないこと。このWebアプリはexecuteAs: USER_DEPLOYINGで
 *  動くため、UserPropertiesは「デプロイした先生」のものになり、全教員で同じ値を
 *  共有してしまう(お互いの条件を上書きし合う)。そのためScriptPropertiesに
 *  メールアドレスをキーに含めて保存する。 */
const LIST_FILTER_PREFIX = 'LIST_FILTERS_';

/** 保存してある絞り込み条件を返す。無ければ空オブジェクト。
 *  記憶は補助機能なので、読めない・壊れている場合も例外にせず{}を返して
 *  「条件なしで開く」に倒す。 */
function getMyListFilters() {
  _assertTeacher();
  const email = _me();
  if (!email) return {};
  try {
    const raw = PropertiesService.getScriptProperties()
      .getProperty(LIST_FILTER_PREFIX + email);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return (obj && typeof obj === 'object') ? obj : {};
  } catch (e) {
    return {};
  }
}

/** 絞り込み条件を保存する。想定した6項目だけを取り出して保存し、
 *  画面から渡された余計なプロパティは保存しない。 */
function saveMyListFilters(filters) {
  _assertTeacher();
  const email = _me();
  if (!email) return { ok: false, error: 'アカウントを特定できません' };
  filters = filters || {};
  const safe = {
    grade:   String(filters.grade   || '').trim(),
    cls:     String(filters.cls     || '').trim(),
    subject: String(filters.subject || '').trim(),
    name:    String(filters.name    || '').trim(),
    range:   String(filters.range   || '').trim(),
    includeClosed: !!filters.includeClosed,
    dueMode: DUE_MODES.indexOf(String(filters.dueMode || '')) !== -1 ? String(filters.dueMode || '') : '',
    dueFrom: String(filters.dueFrom || '').trim().slice(0, 10),
    dueTo:   String(filters.dueTo   || '').trim().slice(0, 10),
  };
  try {
    PropertiesService.getScriptProperties()
      .setProperty(LIST_FILTER_PREFIX + email, JSON.stringify(safe));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message }; // 保存できなくても運用に支障はない
  }
}

function getUnsubmittedList(filters) {
  _assertTeacher();
  filters = filters || {};
  const grade   = String(filters.grade   || '').trim();
  const cls     = String(filters.cls     || '').trim();
  const subject = String(filters.subject || '').trim();
  const name    = String(filters.name    || '').trim();
  const range   = String(filters.range   || '').trim();
  const dueMode = DUE_MODES.indexOf(String(filters.dueMode || '')) !== -1 ? String(filters.dueMode || '') : '';
  const dateOk  = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? String(v) : '';
  const dueFrom = dateOk(filters.dueFrom), dueTo = dateOk(filters.dueTo);

  const matMap = _materialMap();
  const roster = _rosterMaps().byId;
  const latest = _latestMap(matMap);
  const students = Object.keys(roster).map(id => roster[id]);

  let codes = Object.keys(matMap).filter(code => {
    const m = matMap[code];
    if (!filters.includeClosed && m.closed) return false;  // ← 既定は受付終了を隠す
    if (grade   && m.grades.length && m.grades.indexOf(grade) === -1) return false;
    if (subject && String(m.subject) !== subject) return false;
    if (name    && String(m.name) !== name) return false;
    if (range   && m.range !== range) return false;
    if (!_dueMatches(m, dueMode, dueFrom, dueTo)) return false;
    return true;
  });
  // 締切で絞ったとき(締切なし以外)は、締切が近い順を優先して並べる
  const byDue = dueMode && dueMode !== 'none';
  codes.sort((a, b) => {
    const ma = matMap[a], mb = matMap[b];
    if (byDue) {
      const da = _dateKeyOf(ma.due), db = _dateKeyOf(mb.due);
      if (da !== db) return da < db ? -1 : 1;
    }
    if (String(ma.subject) !== String(mb.subject))
      return String(ma.subject).localeCompare(String(mb.subject), 'ja');
    if (String(ma.name) !== String(mb.name))
      return String(ma.name).localeCompare(String(mb.name), 'ja');
    return String(ma.range).localeCompare(String(mb.range), 'ja', { numeric: true });
  });

  const materials = [];
  codes.forEach(code => {
    const m = matMap[code];
    const targets = [];
    students.forEach(st => {
      const gOk = !m.grades.length || m.grades.indexOf(String(st.grade)) !== -1;
      const cOkMat = !m.clses.length || m.clses.indexOf(String(st.cls)) !== -1;
      if (!(gOk && cOkMat)) return;
      if (cls && String(st.cls) !== cls) return;

      const s = _latestFor(latest, matMap, st.id, code);
      let status = null;
      if (!s) status = '未提出';
      else if (s.label === GIVEUP_LABEL) status = '断念';
      else if (m.fails.indexOf(s.label) !== -1) status = '再提出';

      if (!status) return;
      if (m.daily && status === '断念') return; // デイリー教材に断念は存在しない
      const daysLeft = (status === '未提出' && !m.daily) ? _daysUntil(m.due) : null;
      const urgency = status === '未提出' ? (m.daily ? 'today' : _urgencyOf(daysLeft)) : 'none';
      const badge = (status === '未提出' && m.daily) ? '📅毎日' : _badgeText(daysLeft, urgency);
      targets.push({
        id: st.id, name: String(st.name || ''),
        grade: String(st.grade), cls: String(st.cls), num: String(st.num),
        gakkyu: st.grade + '年' + st.cls + '組' + st.num + '番',
        status: status, urgency: urgency, badge: badge,
      });
    });
    if (!targets.length) return;

    const _order = { '再提出': 0, '未提出': 1, '断念': 2 };
    targets.sort((a, b) => {
      if (_order[a.status] !== _order[b.status]) return _order[a.status] - _order[b.status];
      if (a.grade !== b.grade) return a.grade.localeCompare(b.grade, 'ja', { numeric: true });
      if (a.cls !== b.cls) return a.cls.localeCompare(b.cls, 'ja');
      return Number(a.num) - Number(b.num);
    });

    materials.push({
      code: code, subject: String(m.subject), name: m.disp,
      dueText: (m.daily || !m.due) ? '' : Utilities.formatDate(m.due, 'JST', 'M/d'),
      daily: !!m.daily,
      closed: !!m.closed,
      dateLabel: m.daily ? _todayKey() : '',
      resubmitCount: targets.filter(t => t.status === '再提出').length,
      unsubmittedCount: targets.filter(t => t.status === '未提出').length,
      giveupCount: targets.filter(t => t.status === '断念').length,
      targets: targets,
    });
  });

  return { ok: true, materials: materials };
}

/** 教材の「受付終了」フラグを切り替える(職員専用)。
 *  教材シートの「受付終了」列(チェックボックス)にtrue/falseを書くだけ。
 *  ログは一切変更しないため、フラグを外せば生徒側の表示は元(未提出・再提出)に戻る。
 *  ※チェックボックス列なので '○' 等の文字列ではなくbooleanを書くこと
 *  (文字列を書くと入力規則違反の警告マークが付く)。 */
function setMaterialClosed(materialCode, closed) {
  _assertTeacher();
  materialCode = String(materialCode).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MATERIAL);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const closedCol = headers.indexOf('受付終了');
  if (closedCol === -1) {
    return { ok: false,
      error: '教材シートに「受付終了」列がありません。' + SETUP_COLS_HINT };
  }
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === materialCode) {
      sh.getRange(i + 1, closedCol + 1).setValue(!!closed);
      _stampMaterialRow(sh, headers, i + 1, false);
      return { ok: true, code: materialCode, closed: !!closed };
    }
  }
  return { ok: false, error: '教材コードが見つかりません: ' + materialCode };
}

/* ============ 教材タブ(v6.8〜、職員専用) ============
 * 教材シートの先頭10列は _materialMap と同じく「位置」で、
 * 毎日・受付終了・登録者などの追加列は「ヘッダー名」で扱う。
 * 行は必ず教材コードで探す(画面から行番号を受け取らない)。 */

/** 登録者/更新者の列に、実行した先生(_me)と現在時刻を書く。
 *  isNew=true なら登録者・登録日時も書く。列が無ければ何もしない */
function _stampMaterialRow(sh, headers, row, isNew) {
  const me = _me();
  const now = new Date();
  const put = (name, v) => {
    const c = headers.indexOf(name);
    if (c !== -1) sh.getRange(row, c + 1).setValue(v);
  };
  if (isNew) { put('登録者', me); put('登録日時', now); }
  put('更新者', me); put('更新日時', now);
}

/** 教員名簿から メール→教員名 の対応を作る(画面で名前表示するため) */
function _teacherNameMap() {
  const map = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEACHERS);
  if (!sh) return map;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const a = String(r[1] || '').trim().toLowerCase();
    if (a) map[a] = String(r[0] || '').trim() || a;
  });
  return map;
}

/** 教材タブの一覧・選択肢をまとめて返す */
function getMaterialAdminData() {
  _assertTeacher();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MATERIAL);
  const lastCol = Math.max(sh.getLastColumn(), 10);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const col = name => headers.indexOf(name);
  const vals = sh.getDataRange().getValues();

  // ログ件数(取消を含む)。削除してよいかの判定にも使う
  const logCount = {};
  const log = ss.getSheetByName(SHEET_LOG);
  if (log && log.getLastRow() >= 2) {
    log.getRange(2, 2, log.getLastRow() - 1, 1).getValues().forEach(r => {
      const c = String(r[0]).trim();
      if (c) logCount[c] = (logCount[c] || 0) + 1;
    });
  }

  const names = _teacherNameMap();
  const who = a => { a = String(a || '').trim().toLowerCase(); return a ? (names[a] || a) : ''; };
  const fmt = (d, p) => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, 'JST', p) : '';
  const me = _me();

  const materials = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const code = String(r[0] || '').trim();
    if (!code) continue;
    const due = _parseDueDate(r[7]);
    const get = name => col(name) !== -1 ? r[col(name)] : '';
    materials.push({
      code: code,
      grade: String(r[1] || '').trim(), cls: String(r[2] || '').trim(),
      subject: String(r[3] || '').trim(), name: String(r[4] || '').trim(),
      abbr: String(r[5] || '').trim(), range: String(r[6] || '').trim(),
      due: fmt(due, 'yyyy-MM-dd'), dueText: fmt(due, 'M/d'),
      labels: String(r[8] || '合格,再提出'), fails: String(r[9] || '再提出'),
      daily: _isDailyFlag(get('毎日')), closed: _isDailyFlag(get('受付終了')),
      logCount: logCount[code] || 0,
      createdBy: who(get('登録者')), createdAt: fmt(get('登録日時'), 'M/d'),
      updatedBy: who(get('更新者')), updatedAt: fmt(get('更新日時'), 'M/d'),
      mine: String(get('登録者') || '').trim().toLowerCase() === me,
    });
  }

  const abbrs = _abbrMaps();
  const subjects = abbrs.subject.map(s => s.name);
  const roster = ss.getSheetByName(SHEET_ROSTER).getDataRange().getValues().slice(1);
  const uniq = arr => [...new Set(arr.map(v => String(v).trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, 'ja', { numeric: true }));

  return {
    ok: true,
    materials: materials,
    subjects: subjects,
    recommended: abbrs.material.map(r => ({ name: r.name, abbr: r.ab })),  // 略称シートの種別「教材」
    grades: uniq(roster.map(r => r[1])),
    classes: uniq(roster.map(r => r[2])),
    hasDaily: col('毎日') !== -1,
    hasClosed: col('受付終了') !== -1,
    hasAudit: col('登録者') !== -1,
    demo: !!IS_DEMO_BUILD,
    printUrl: ScriptApp.getService().getUrl() + '?page=print',
  };
}

/** 画面から来た教材の入力を検証・整形する。問題があれば {error} を返す。
 *  forUpdate=true のときは教科・教材名略称を見ない(コードの一部なので変えない) */
function _normalizeMaterialInput(p, forUpdate) {
  p = p || {};
  const out = {};
  const csv = v => (Array.isArray(v) ? v : String(v || '').split(','))
    .map(s => String(s).trim()).filter(Boolean);

  if (!forUpdate) {
    out.subject = String(p.subject || '').trim();
    if (!out.subject) return { error: '教科を選んでください' };
    out.subjectAbbr = _findAbbr(_abbrMaps().subject, out.subject);
    if (!out.subjectAbbr) return { error: '教科「' + out.subject + '」の略称が略称シートにありません' };
    out.abbr = String(p.abbr || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{1,10}$/.test(out.abbr)) return { error: '教材名略称は半角英数字(10文字まで)で入力してください' };
  }
  out.name = String(p.name || '').trim();
  if (!out.name) return { error: '教材名を入力してください' };

  out.grades = csv(p.grades);
  if (out.grades.some(g => !/^\d$/.test(g))) return { error: '学年が不正です' };
  out.classes = csv(p.classes);
  if (out.classes.some(c => c.length > 10)) return { error: '組が不正です' };

  out.due = null;
  const dueStr = String(p.due || '').trim();
  if (dueStr) {
    const m = dueStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return { error: '提出期限の形式が不正です' };
    out.due = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  out.labels = csv(p.labels);
  if (!out.labels.length) return { error: '評価ラベルを1つ以上入力してください' };
  if (new Set(out.labels).size !== out.labels.length) return { error: '評価ラベルが重複しています' };
  if (out.labels.indexOf(GIVEUP_LABEL) !== -1) return { error: '「' + GIVEUP_LABEL + '」は評価ラベルに使えません' };
  out.fails = csv(p.fails);
  if (out.fails.some(f => out.labels.indexOf(f) === -1)) return { error: '再提出扱いは評価ラベルの中から選んでください' };

  out.daily = !!p.daily;
  out.closed = !!p.closed;
  return out;
}

/** 範囲・学年・組まで同じ教材が既にあるか(重複警告用)。見つかったコードを返す */
function _findDuplicateMaterials(vals, subject, name, range, grades, classes) {
  const key = (g, c) => [...g].sort().join(',') + '|' + [...c].sort().join(',');
  const want = key(grades, classes);
  const split = v => String(v || '').split(',').map(s => s.trim()).filter(Boolean);
  const hits = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!String(r[0] || '').trim()) continue;
    if (String(r[3]).trim() === subject && String(r[4]).trim() === name &&
        String(r[6] || '').trim() === range && key(split(r[1]), split(r[2])) === want) {
      hits.push(String(r[0]).trim());
    }
  }
  return hits;
}

/** 教材を新規登録する。範囲(1行1範囲)の数だけ行を作り、コードを自動採番する。
 *  p.allowDuplicate=true でなければ、同じ教材があるとき登録せずに確認を求める */
function createMaterials(p) {
  _assertTeacher();
  const v = _normalizeMaterialInput(p, false);
  if (v.error) return { ok: false, error: v.error };
  const ranges = String((p && p.ranges) || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (!ranges.length) ranges.push('');
  if (ranges.length > 50) return { ok: false, error: '一度に登録できる範囲は50件までです' };
  if (new Set(ranges).size !== ranges.length) return { ok: false, error: '範囲が重複しています' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MATERIAL);
    const width = Math.max(sh.getLastColumn(), 10);
    const headers = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
    const dailyCol = headers.indexOf('毎日');
    const closedCol = headers.indexOf('受付終了');
    if (v.daily && dailyCol === -1) {
      return { ok: false, error: '教材シートに「毎日」列がありません。' + SETUP_COLS_HINT };
    }
    const vals = sh.getDataRange().getValues();

    if (!(p && p.allowDuplicate)) {
      const dup = [];
      ranges.forEach(rg => {
        _findDuplicateMaterials(vals, v.subject, v.name, rg, v.grades, v.classes)
          .forEach(c => dup.push(c));
      });
      if (dup.length) return { ok: false, duplicate: dup };
    }

    const prefix = v.subjectAbbr + '-' + v.abbr;
    let serial = _materialSerialMax(vals)[prefix] || 0;
    const rows = [], items = [];
    ranges.forEach(rg => {
      serial++;
      const code = prefix + '-' + ('0' + serial).slice(-2);
      const row = new Array(width).fill('');
      row[0] = code; row[1] = v.grades.join(','); row[2] = v.classes.join(',');
      row[3] = v.subject; row[4] = v.name; row[5] = v.abbr; row[6] = rg;
      row[7] = v.due || ''; row[8] = v.labels.join(','); row[9] = v.fails.join(',');
      if (dailyCol !== -1) row[dailyCol] = v.daily ? '○' : '';
      if (closedCol !== -1) row[closedCol] = false;
      rows.push(row);
      items.push({ code: code, disp: v.name + (rg ? '(' + rg + ')' : '') });
    });

    if (IS_DEMO_BUILD) return { ok: true, demo: true, items: items };

    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, width).setValues(rows);
    if (closedCol !== -1) sh.getRange(start, closedCol + 1, rows.length, 1).insertCheckboxes();
    for (let k = 0; k < rows.length; k++) _stampMaterialRow(sh, headers, start + k, true);
    return { ok: true, items: items };
  } finally {
    lock.releaseLock();
  }
}

/** 教材を1件更新する。教材コード・教科・教材名略称は変えない(コードの一部のため) */
function updateMaterial(code, p) {
  _assertTeacher();
  code = String(code || '').trim();
  const v = _normalizeMaterialInput(p, true);
  if (v.error) return { ok: false, error: v.error };
  const range = String((p && p.range) || '').trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MATERIAL);
    const width = Math.max(sh.getLastColumn(), 10);
    const headers = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
    const dailyCol = headers.indexOf('毎日');
    const closedCol = headers.indexOf('受付終了');
    if (v.daily && dailyCol === -1) {
      return { ok: false, error: '教材シートに「毎日」列がありません。' + SETUP_COLS_HINT };
    }
    if (v.closed && closedCol === -1) {
      return { ok: false, error: '教材シートに「受付終了」列がありません。' + SETUP_COLS_HINT };
    }
    const vals = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === code) { row = i + 1; break; }
    }
    if (row === -1) return { ok: false, error: '教材コードが見つかりません: ' + code };

    if (IS_DEMO_BUILD) return { ok: true, demo: true };

    sh.getRange(row, 2, 1, 2).setValues([[v.grades.join(','), v.classes.join(',')]]);
    sh.getRange(row, 5).setValue(v.name);
    sh.getRange(row, 7, 1, 4).setValues([[range, v.due || '', v.labels.join(','), v.fails.join(',')]]);
    if (dailyCol !== -1) {
      // 既にONの行は元の表記(○やTRUE)を残し、OFF→ONのときだけ○を書く
      const cur = _isDailyFlag(vals[row - 1][dailyCol]);
      if (cur !== v.daily) sh.getRange(row, dailyCol + 1).setValue(v.daily ? '○' : '');
    }
    if (closedCol !== -1) {
      const cell = sh.getRange(row, closedCol + 1);
      cell.insertCheckboxes();
      cell.setValue(v.closed);
    }
    _stampMaterialRow(sh, headers, row, false);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** 教材を削除する。ログに1件でも記録(取消を含む)があれば削除しない */
function deleteMaterial(code) {
  _assertTeacher();
  code = String(code || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName(SHEET_LOG);
    if (log && log.getLastRow() >= 2) {
      const used = log.getRange(2, 2, log.getLastRow() - 1, 1).getValues()
        .some(r => String(r[0]).trim() === code);
      if (used) return { ok: false, error: '記録があるため削除できません。使わない場合は受付終了にしてください' };
    }
    const sh = ss.getSheetByName(SHEET_MATERIAL);
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === code) {
        if (IS_DEMO_BUILD) return { ok: true, demo: true };
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: '教材コードが見つかりません: ' + code };
  } finally {
    lock.releaseLock();
  }
}

/* ============ 印刷ページへのリンク ============ */
function _settingsRowIsEmpty(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sh) return true;
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) {
      return String(vals[i][1] || '').trim() === '';
    }
  }
  return true;
}

function _settingsRowValue(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sh) return '';
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) return String(vals[i][1] || '').trim();
  }
  return '';
}

function _setSettingsRow(key, value, asFormula) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) { row = i + 1; break; }
  }
  if (row === -1) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue(key);
  }
  if (asFormula) sh.getRange(row, 2).setFormula(value);
  else sh.getRange(row, 2).setValue(value);
}

function _writeWebAppUrlToSettings(url) {
  _setSettingsRow('WebアプリURL',
    '=HYPERLINK("' + url + '", "🔗 ' + url + '")', true);
  _setSettingsRow('印刷ページ',
    '=HYPERLINK("' + url + '?page=print", "🖨 QRコード印刷ページを開く")', true);
}

function showPrintPageLink() {
  const ui = SpreadsheetApp.getUi();
  let base = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '';
  if (!base) {
    const res = ui.prompt('印刷ページへのリンクを作成',
      'まだアプリのURLを記憶できていません。\n' +
      '普段お使いのWebアプリのURL(.../execで終わるもの)を貼り付けてください\n' +
      '(一度アプリをブラウザで開けば、次回からは自動になります)',
      ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    base = res.getResponseText().trim();
  }
  base = base.split('?')[0].replace(/\/+$/, '');
  if (base.indexOf('/exec') === -1) {
    ui.alert('「.../exec」で終わるURLではないようです。もう一度確認してください: ' + base);
    return;
  }
  _writeWebAppUrlToSettings(base);
  ui.alert(
    '設定シートの「WebアプリURL」「印刷ページ」の行にリンクを表示しました。\n' +
    'セルをクリックすると開きます。');
}

