#!/usr/bin/env bash
set -euo pipefail

source_image="${1:-assets/icon.png}"
target_icon="${2:-assets/icon.icns}"
iconset_dir="${target_icon%.icns}.iconset"

if [[ ! -f "$source_image" ]]; then
  echo "Simge kaynağı bulunamadı: $source_image" >&2
  exit 1
fi

rm -rf "$iconset_dir"
mkdir -p "$iconset_dir"

for size in 16 32 128 256 512; do
  double_size=$((size * 2))
  sips -z "$size" "$size" "$source_image" --out "$iconset_dir/icon_${size}x${size}.png" >/dev/null
  sips -z "$double_size" "$double_size" "$source_image" --out "$iconset_dir/icon_${size}x${size}@2x.png" >/dev/null
done

if ! iconutil -c icns "$iconset_dir" -o "$target_icon"; then
  # Some macOS releases reject otherwise valid generated iconsets. Sips can
  # produce the same native ICNS container directly from the RGBA source.
  sips -s format icns "$source_image" --out "$target_icon" >/dev/null
fi
rm -rf "$iconset_dir"
echo "$target_icon"
