#!/bin/bash

cd "$(dirname "$0")" || exit 1
LAUNCHER="thermofisher_invoice_app_simplex/launcher.py"

pause_then_exit() {
  echo
  read -r -p "Press Return to close this window. " _
  exit "$1"
}

if [ ! -f "$LAUNCHER" ]; then
  echo "Could not find $LAUNCHER next to this file."
  echo "Keep this file inside the unzipped Invoice-Generation folder and try again."
  pause_then_exit 1
fi

export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

supported() {
  "$1" -c 'import platform, sys, sysconfig
ok = ((3, 10) <= sys.version_info[:2] <= (3, 14) and sys.maxsize > 2**32
      and platform.python_implementation() == "CPython"
      and not sysconfig.get_config_var("Py_GIL_DISABLED"))
sys.exit(0 if ok else 1)' </dev/null >/dev/null 2>&1
}

find_python() {
  local versions="3.12 3.13 3.11 3.14 3.10"
  [ "$(uname -m)" = "x86_64" ] && versions="3.12 3.11 3.10 3.13 3.14"
  local v p
  for v in $versions; do
    for p in "/Library/Frameworks/Python.framework/Versions/$v/bin/python$v" \
             "/opt/homebrew/bin/python$v" "/usr/local/bin/python$v" \
             "$(command -v "python$v" 2>/dev/null)"; do
      if [ -n "$p" ] && [ -x "$p" ] && supported "$p"; then echo "$p"; return 0; fi
    done
  done
  p="$(command -v python3 2>/dev/null)"
  if [ -n "$p" ] && ! { [ "$(uname)" = "Darwin" ] && [ "$p" = "/usr/bin/python3" ]; } && supported "$p"; then
    echo "$p"; return 0
  fi
  return 1
}

if ! PY="$(find_python)"; then
  echo "Python 3.10 to 3.14 was not found on this computer."
  echo
  if [ "$(uname)" = "Darwin" ]; then
    echo "Install Python 3.13 from https://www.python.org/downloads/macos/"
    echo "(the page opens now): choose the \"macOS 64-bit universal2 installer\","
    echo "run it, then double click this file again."
    [ "$(uname -m)" = "x86_64" ] && echo "On an Intel Mac, Python 3.12.10 works best."
    open "https://www.python.org/downloads/macos/"
  else
    echo "Install Python 3.12 or 3.13 with your package manager (for example the"
    echo "python3.12 and python3.12-venv packages), then run this again."
  fi
  pause_then_exit 1
fi

"$PY" "$LAUNCHER" "$@"
status=$?
[ "$status" -ne 0 ] && pause_then_exit "$status"
exit 0
