#!/usr/bin/env bash
# 配布用ビルド:開発版の作業フォルダ(clasp)から gas/ に配布用ファイルを作る。
#  - コード.js → gas/Code.gs(IS_DEV_BUILD=false にする。IS_DEMO_BUILD=false を確認)
#  - index.html / student.html / print.html をそのまま
#  - gas/appsscript.json はリポジトリで管理している公開用(組織内のみ・Classroom)を使う
#  - 最後に check-public.sh を通す
set -eu
cd "$(dirname "$0")/.."
SRC="${SRC:-../提出物チェック/dev}"
[ -f "$SRC/コード.js" ] || { echo "中止: $SRC/コード.js がありません"; exit 1; }

# GAS が付ける先頭の空行と CR を落とす
clean() { sed 's/\r$//' "$1" | sed '/./,$!d'; }

clean "$SRC/コード.js" | sed 's/^const IS_DEV_BUILD = true;$/const IS_DEV_BUILD = false;/' > gas/Code.gs
for f in index.html student.html print.html; do clean "$SRC/$f" > "gas/$f"; done

grep -Eq "^const IS_DEV_BUILD = false;" gas/Code.gs || { echo "中止: IS_DEV_BUILD が false になっていません"; exit 1; }
grep -Eq "^const IS_DEMO_BUILD = false;" gas/Code.gs || { echo "中止: IS_DEMO_BUILD が false ではありません"; exit 1; }
grep -E "^const APP_VERSION" gas/Code.gs

./scripts/check-public.sh
