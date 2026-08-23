#!/bin/bash
# 新しい商品登録提案ファイルを生成するスクリプト
# 使い方: bash templates/new-proposal.sh "商品名"

PRODUCT_NAME="${1:-新商品}"
DATE=$(date +%Y%m%d)
OUTPUT_DIR="proposals"
TEMPLATE="templates/satofuru-proposal-template.md"
OUTPUT_FILE="${OUTPUT_DIR}/${DATE}_${PRODUCT_NAME}.md"

mkdir -p "$OUTPUT_DIR"

if [ ! -f "$TEMPLATE" ]; then
  echo "エラー: テンプレートが見つかりません: $TEMPLATE"
  exit 1
fi

sed "s/作成日 | /作成日 | ${DATE}/" "$TEMPLATE" > "$OUTPUT_FILE"

echo "✅ 提案ファイルを作成しました: $OUTPUT_FILE"
echo "   次のステップ: ファイルを開いて各項目を記入してください"
