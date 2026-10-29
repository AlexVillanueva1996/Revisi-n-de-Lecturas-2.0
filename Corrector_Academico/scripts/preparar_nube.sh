#!/usr/bin/env bash
set -euo pipefail
cd /workspace/Revisi-n-de-Lecturas-2.0/Corrector_Academico
python3 -m venv .venv
.venv/bin/python -m pip install --disable-pip-version-check -r requirements.txt
.venv/bin/python -m pip check
command -v dot
command -v soffice
.venv/bin/python -m src.main --diagnostico
mkdir -p TEMPORALES
.venv/bin/python -m unittest discover -s tests -v > TEMPORALES/verificacion_entorno.log 2>&1
