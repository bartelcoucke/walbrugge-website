#!/bin/bash
set -e
export PATH="/home/linuxbrew/.linuxbrew/bin:$PATH"
cd ~/walbrugge-repo

echo "=== WebP Conversie voor walbrugge.be ==="
echo ""

STATS="/tmp/webp_stats_$$.txt"
> "$STATS"
COUNT=0

convert_one() {
  local input="$1"
  local quality="$2"
  local output="${input%.*}.webp"

  local orig_size=$(stat -c%s "$input" 2>/dev/null) || return
  echo "Converting: $input (q=$quality)"
  cwebp -quiet -q "$quality" "$input" -o "$output"

  local webp_size=$(stat -c%s "$output" 2>/dev/null) || return
  local saved=$((orig_size - webp_size))
  local pct=$((100 * saved / orig_size))
  echo "  -> $(basename "$output") ($(numfmt --to=iec $orig_size) -> $(numfmt --to=iec $webp_size), bespaard ${pct}%)"
  echo "$orig_size $webp_size" >> "$STATS"
}

echo "--- Logo's/icons (kwaliteit 90) ---"
for f in \
  public/assets/img/booking-logo.png \
  public/assets/img/logo-gold.png \
  public/assets/img/logo-white.png \
  public/assets/img/logo.png \
  public/assets/img/walbrugge_logo.png \
  public/assets/img/walbrugge_logo_full_white.png \
  public/assets/img/walbrugge_logo_white.png; do
  if [ -f "$f" ]; then
    convert_one "$f" 90
    COUNT=$((COUNT + 1))
  fi
done

echo ""
echo "--- Foto's (kwaliteit 82) ---"

# Build list of files to convert (excluding skip list)
CONVERT_LIST=$(mktemp)

find public/assets/img/ -type f \( -iname '*.jpg' -o -iname '*.jpeg' -o -iname '*.png' \) | while IFS= read -r f; do
  base=$(basename "$f")
  case "$base" in
    favicon.png|apple-touch-icon.png|og-image.jpg|og-image-original.jpg|walbrugge_logo_round.png|walbrugge_logo_round.webp|booking-logo.png|logo-gold.png|logo-white.png|logo.png|walbrugge_logo.png|walbrugge_logo_full_white.png|walbrugge_logo_white.png)
      echo "SKIPPING: $f (op skip lijst)"
      ;;
    *)
      echo "$f" >> "$CONVERT_LIST"
      ;;
  esac
done

while IFS= read -r f; do
  [ -z "$f" ] && continue
  convert_one "$f" 82
  COUNT=$((COUNT + 1))
done < "$CONVERT_LIST"
rm -f "$CONVERT_LIST"

echo ""
echo "=== Conversie voltooid ==="

TOTAL_ORIG=0
TOTAL_WEBP=0
while read o w; do
  TOTAL_ORIG=$((TOTAL_ORIG + o))
  TOTAL_WEBP=$((TOTAL_WEBP + w))
done < "$STATS"
rm -f "$STATS"

TOTAL_SAVED=$((TOTAL_ORIG - TOTAL_WEBP))
echo "Bestanden geconverteerd: $COUNT"
echo "Totale originele grootte: $(numfmt --to=iec $TOTAL_ORIG)"
echo "Totale WebP grootte:      $(numfmt --to=iec $TOTAL_WEBP)"
echo "Totale besparing:         $(numfmt --to=iec $TOTAL_SAVED)"
