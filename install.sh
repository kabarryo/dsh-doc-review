#!/usr/bin/env bash
# Install dsh-doc-review into a DSH profile by link + profile bundle entry.
set -euo pipefail
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROFILE="${DSH_PROFILE_DIR:-$HOME/.dsh/profiles/web}"

python3 - "$PROFILE/package.json" "$SRC" <<'PY'
import json
import shutil
import sys
import time
from pathlib import Path

manifest = Path(sys.argv[1])
source = sys.argv[2]
backup = manifest.with_name(manifest.name + ".bak-" + time.strftime("%Y%m%d-%H%M%S"))
shutil.copy2(manifest, backup)
data = json.loads(manifest.read_text("utf-8"))
deps = data.setdefault("dependencies", {})
deps["dsh-doc-review"] = f"link:{source}"
bundles = data["dsh"]["profile"].setdefault("bundles", [])
if "dsh-doc-review" not in bundles:
    bundles.append("dsh-doc-review")
manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", "utf-8")
print(f"manifest updated (backup: {backup})")
PY

ln -sfn "$SRC" "$PROFILE/node_modules/dsh-doc-review"
echo "linked: $PROFILE/node_modules/dsh-doc-review -> $SRC"

BIN="$HOME/.dsh/doc-review/bin"
mkdir -p "$BIN" "$HOME/.dsh/doc-review/logs"
cp "$SRC/scripts/"*.py "$BIN/"
chmod +x "$BIN/"*.py
echo "helpers: $BIN"
