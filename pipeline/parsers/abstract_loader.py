"""Read-only loader for pre-built loan/JV abstracts. Zero PDF/Claude dependency
here — abstracts are generated once, offline, by tools/extract_loan_abstract.py.
Returns None whenever no abstract exists yet, so views can render a graceful
"not on file" message instead of crashing.
"""

from __future__ import annotations

import json
from typing import Optional

from pipeline.property_config import JVDocumentRef, LoanRef, PropertyConfig


def _load_json(cfg: PropertyConfig, abstract_file: str, data_dir: str = "data") -> Optional[dict]:
    if not abstract_file:
        return None
    path = cfg.abstract_path(abstract_file, data_dir)
    if not path.exists():
        return None
    return json.loads(path.read_text())


def load_loan_abstract(cfg: PropertyConfig, loan_ref: LoanRef, data_dir: str = "data") -> Optional[dict]:
    return _load_json(cfg, loan_ref.abstract_file, data_dir)


def load_jv_abstract(cfg: PropertyConfig, jv_ref: JVDocumentRef, data_dir: str = "data") -> Optional[dict]:
    return _load_json(cfg, jv_ref.abstract_file, data_dir)
