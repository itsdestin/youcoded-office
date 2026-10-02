#!/usr/bin/env bash
# Fails if any program or library in a Windows converter folder needs a DLL that is neither in
# that folder nor part of every Windows 10/11 install.
# Usage: build/check-win-imports.sh <converter folder>   (needs objdump, from binutils)
#
# WHY (2026-10-02, clean Windows 11 VM): x2t needs Microsoft's Visual C++ runtime
# (VCRUNTIME140.dll, VCRUNTIME140_1.dll, MSVCP140.dll). GitHub's Windows machines have it
# installed, so the smoke test passed; a fresh PC does not, x2t exited 0xC0000135 ("DLL not
# found") and every document sat on "Opening…" forever. This check reads what the files ask for
# instead of trusting the build machine, so a missing library fails the build, not a user.
set -euo pipefail
DIR="${1:?usage: check-win-imports.sh <converter folder>}"
# Libraries every supported Windows ships in System32 (the UCRT api-ms-win-crt-* set is part of
# Windows 10 and later). Anything else must be in DIR.
SYSTEM='^(api-ms-win-.*|advapi32|bcrypt|crypt32|cryptui|dbghelp|gdi32|gdiplus|kernel32|ole32|oleaut32|rpcrt4|shell32|shlwapi|urlmon|user32|winmm|ws2_32|version|secur32|ncrypt|iphlpapi|winhttp|wininet|comdlg32|imm32|setupapi|usp10|normaliz|msimg32)\.dll$'
missing=0
shopt -s nullglob nocaseglob
for f in "$DIR"/*.exe "$DIR"/*.dll; do
  for dll in $(objdump -p "$f" | awk '/DLL Name/{print $3}'); do
    low=$(printf '%s' "$dll" | tr 'A-Z' 'a-z')
    [[ "$low" =~ $SYSTEM ]] && continue
    if ! ls "$DIR" | grep -qix -- "$dll"; then
      echo "$(basename "$f") needs $dll, which is not in $DIR and not part of Windows" >&2
      missing=1
    fi
  done
done
[ "$missing" = 0 ] && echo "every DLL the Windows converter needs is shipped or part of Windows"
exit "$missing"
