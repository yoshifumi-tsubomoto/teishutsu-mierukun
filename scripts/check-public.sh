#!/usr/bin/env bash
# 公開前の点検:入れてはいけない文字がリポジトリに無いかを調べる。
# 1つでも見つかれば一覧を出して終了コード1で止まる(push の前に必ず通す)。
#
# ・ここに書くのは「一般的な形」だけ(学校のドメイン、個人のアカウントなど)。
# ・本番・開発版の ID のような「それ自体が秘密」の文字列は、リポジトリの外の
#   ファイル(PRIVATE_LIST、1行1語)に置き、このスクリプトはそれを読むだけにする。
#   そうしないと、点検のためのリストを公開してしまうため。
set -u
cd "$(dirname "$0")/.."
PRIVATE_LIST="${PRIVATE_LIST:-../提出物チェック/公開チェック_非公開リスト.txt}"

# 調べる対象:git に入る(入れる予定の)ファイルすべて。scripts/ 自身も含める
FILES=$(git ls-files --cached --others --exclude-standard 2>/dev/null || find . -type f -not -path './.git/*')

ng=0
hit() { echo "NG: $1"; echo "$2" | sed 's/^/    /'; ng=1; }

check() {  # check <説明> <正規表現> [除外するファイルの正規表現]
  local out
  out=$(echo "$FILES" | grep -v -E "${3:-^$}" | xargs -d '\n' grep -n -i -E -- "$2" 2>/dev/null)
  [ -n "$out" ] && hit "$1" "$out"
}

check "学校のドメイン"            'tonami|ed\.jp'                 '^scripts/check-public\.sh$'
check "個人のGoogleアカウント"     '@gmail\.com|tsubotto|tsubomoto-yoshifumi' '^scripts/check-public\.sh$'
check "生徒のアカウントらしき文字" '\bst[0-9]{6}\b'
check "校名"                      '庄川'                           '^scripts/check-public\.sh$'
# 作者名は LICENSE・NOTICE・README と紹介ページの著作権表示にだけ書く
check "作者名(LICENSE/NOTICE/README/紹介ページ 以外)" '坪本' '^(LICENSE|NOTICE|README\.md|docs/index\.html|scripts/check-public\.sh)$'

if [ -f "$PRIVATE_LIST" ]; then
  while IFS= read -r w; do
    w="${w%$'\r'}"
    [ -z "$w" ] || [ "${w#\#}" != "$w" ] && continue
    out=$(echo "$FILES" | xargs -d '\n' grep -n -F -- "$w" 2>/dev/null)
    [ -n "$out" ] && hit "非公開リストの語(${w:0:6}…)" "$out"
  done < "$PRIVATE_LIST"
else
  echo "注意: 非公開リスト($PRIVATE_LIST)が見つからないため、ID の点検をしていません"
  ng=1
fi

if [ "$ng" = 0 ]; then echo "OK: 入れてはいけない文字は見つかりませんでした"; fi
exit "$ng"
