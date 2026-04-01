#!/bin/zsh
cd "/Users/tiantian/Documents/tt_compare" || exit 1

if [[ -x ".venv/bin/python" ]]; then
  ".venv/bin/python" "launch_compare_app.py"
else
  python3 "launch_compare_app.py"
fi
