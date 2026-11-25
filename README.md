# 提出物みえるくん

**提出物を QR コードでさっとチェック。生徒は自分の提出状況をスマホや Chromebook で確かめられます。**

ワーク・プリント・ノートなどの提出を、教材の QR → 評価 → 生徒の QR の順に読むだけで記録します。未提出・再提出の生徒には Google Classroom で個別に知らせることができます。Google スプレッドシートと Google Apps Script だけで動きます。

## 見る・試す

- 紹介ページ：https://yoshifumi-tsubomoto.github.io/teishutsu-mierukun/
- **体験版**（見本のデータで画面を触れます。Google アカウントでログインして開きます）：https://script.google.com/macros/s/AKfycbz6YF99JsB9BQO6xfZroeYpfpmcBjX06nl-o3dx4AWTjxSlw22aQcdmsZfvcIAiilMc/exec
- 説明書（導入と使い方）：https://docs.google.com/document/d/1mZe16dx0zZkBDOiZnLO7_HBp_gGcrWJZSoC4k-TVN3o/edit?usp=sharing

## しくみ

```
GitHub（プログラムと版の記録を配るだけ）
   ├─→ A校がテンプレートをコピー → A校の Google アカウントで動く（データは A校のドライブ）
   └─→ B校がテンプレートをコピー → B校の Google アカウントで動く（データは B校のドライブ）
```

- 各学校が、**自分の学校の Google アカウント**にテンプレートのスプレッドシートをコピーして使います。作者のサーバーはなく、データが作者に届くこともありません
- 公開の設定は「組織内のみ」です。生徒は学校のアカウントのまま開けます
- QR コードの中身は生徒 ID だけです。生徒は自分の分しか見られません

## できること

- QR コードでの提出チェック（USB・Bluetooth のハンディリーダー、カメラ、写真での読み取り）、取り消し（UNDO/REDO）
- 教材ごとの評価ラベル（合格/再提出、5〜1 など）、再提出の回数の記録
- 未提出・再提出リスト（学年・組・教科・教材名・範囲・締切での絞り込み、条件の記憶）
- 未提出・再提出の生徒への Classroom 通知
- 生徒のマイページ（やること・提出済み、提出をあきらめる、自分用のタスク）
- 教員画面からの教材の登録・編集（教材コードの自動採番、登録者の記録）
- 毎日集める教材（デイリー）、受付終了、集計シート、QR ラベルの印刷

## 使えるもの

- **学校の Google アカウント（Google Workspace for Education など）**
- 学校で使うときは、**学校の情報管理のきまりに従い、管理職などに確認してから**使ってください

## はじめかた

1. [テンプレートのスプレッドシート](https://docs.google.com/spreadsheets/d/1VLvvcstELPIuUoD_LZRFwEKtJFTsCyOsLVecWB8v8tU/copy)を、学校のアカウントでコピーします
2. [説明書](https://docs.google.com/document/d/1mZe16dx0zZkBDOiZnLO7_HBp_gGcrWJZSoC4k-TVN3o/edit?usp=sharing)のとおり、名簿・教材を入れて、Web アプリとしてデプロイします（30 分ほど）

テンプレートには、このリポジトリの [gas/](gas/) と同じファイルが入っています。

## 新しい版が出たら

[Releases](https://github.com/yoshifumi-tsubomoto/teishutsu-mierukun/releases) に、版ごとの変更点を書いています。新しい版にするときは、[gas/](gas/) の `Code.gs`・`index.html`・`student.html`・`print.html` を、自分の学校の Apps Script に貼り替えて、「デプロイを管理」から新しいバージョンでデプロイしてください（Google Apps Script のしくみ上、自動では更新されません）。`appsscript.json` は貼り替えなくてかまいません。

## 要望・不具合の連絡

- [Issues](https://github.com/yoshifumi-tsubomoto/teishutsu-mierukun/issues)

## ライセンス

Apache License 2.0（[LICENSE](LICENSE)・[NOTICE](NOTICE)）。無料で使え、改変して配ることもできます。保証や保守はありません。

Copyright 2026 坪本 吉史
