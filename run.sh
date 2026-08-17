#!/bin/bash
# Activa el venv y corre actualizar.py desde donde sea que estés parado.
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$DIR/venv/bin/activate"
python3 "$DIR/actualizar.py"
