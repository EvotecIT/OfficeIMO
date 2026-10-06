"""Download only the pinned, hash-verified test corpus; never used by product code."""
import hashlib
import json
from pathlib import Path
import sys
from urllib.request import urlopen

manifest = json.loads(Path(__file__).with_name("corpus.json").read_text())
if len(sys.argv) != 2:
    raise SystemExit("Usage: python fetch-corpus.py <test-corpus-directory>")
root = Path(sys.argv[1])
root.mkdir(parents=True, exist_ok=True)
base = "https://raw.githubusercontent.com/c2pa-org/public-testfiles/" + manifest["revision"] + "/"
for case in manifest["cases"]:
    name = case["file"]
    if Path(name).name != name or "/" in name or "\\" in name:
        raise SystemExit("Corpus entries must be filenames")
    target = root / name
    if target.exists():
        data = target.read_bytes()
    else:
        with urlopen(base + manifest["sourceDirectory"] + "/" + name, timeout=30) as response:
            data = response.read(8 * 1024 * 1024 + 1)
    if len(data) > 8 * 1024 * 1024 or hashlib.sha256(data).hexdigest() != case["sha256"]:
        raise SystemExit("Corpus hash or size check failed: " + name)
    if not target.exists():
        with target.open("xb") as output:
            output.write(data)
    print(name, case["sha256"])
for name in ("LICENSE", "README.md"):
    target = root / ("UPSTREAM-" + name)
    if not target.exists():
        with urlopen(base + name, timeout=30) as response:
            data = response.read(1024 * 1024 + 1)
        if len(data) > 1024 * 1024:
            raise SystemExit("Oversized attribution file")
        with target.open("xb") as output:
            output.write(data)
