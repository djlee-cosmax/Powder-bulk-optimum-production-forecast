#!/bin/bash
# 매주 월요일 새 xlsx 파일을 받아 data/ 폴더에 배치 + manifest 갱신
# 사용법: bash update_data.sh <표준대비실적_xlsx> <폐기데이터_xlsx>
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DATA_DIR="$SCRIPT_DIR/data"
mkdir -p "$DATA_DIR"

if [ -z "$1" ] || [ -z "$2" ]; then
    echo "사용법: bash update_data.sh <표준대비실적_xlsx> <폐기데이터_xlsx>"
    echo ""
    echo "예시:"
    echo "  bash update_data.sh \\"
    echo "    \"/mnt/c/Users/djlee/OneDrive - COSMAX/바탕 화면/표준 대비 실적 데이터_260427.xlsx\" \\"
    echo "    \"/mnt/c/Users/djlee/OneDrive - COSMAX/바탕 화면/파우더 벌크 폐기 데이터_260427.xlsx\""
    exit 1
fi

if [ ! -f "$1" ]; then echo "에러: 파일 없음 - $1"; exit 1; fi
if [ ! -f "$2" ]; then echo "에러: 파일 없음 - $2"; exit 1; fi

cp "$1" "$DATA_DIR/standard_perf.xlsx"
cp "$2" "$DATA_DIR/disposal.xlsx"

DATE=$(date +%Y-%m-%d)
STD_NAME="$(basename "$1")"
DISP_NAME="$(basename "$2")"

cat > "$DATA_DIR/manifest.json" <<EOF
{
  "standardPerf": {
    "updated": "$DATE",
    "originalFilename": "$STD_NAME"
  },
  "disposal": {
    "updated": "$DATE",
    "originalFilename": "$DISP_NAME"
  }
}
EOF

echo "데이터 업데이트 완료 ($DATE)"
echo "  표준대비실적: $STD_NAME → standard_perf.xlsx"
echo "  폐기데이터:   $DISP_NAME → disposal.xlsx"
echo ""
echo "다음 단계 (깃 배포):"
echo "  cd $SCRIPT_DIR"
echo "  git add data/"
echo "  git commit -m \"데이터 업데이트 $DATE\""
echo "  git push"
