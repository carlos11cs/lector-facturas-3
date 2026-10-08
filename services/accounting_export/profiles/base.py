from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Sequence

from ..models import AccountingExportData


@dataclass(frozen=True)
class ExportTable:
    rows: Sequence[Mapping]
    columns: Sequence[str]
    sheet_name: str
    filename_prefix: str


class ExportProfile:
    profile_id: str

    def table(self, export_data: AccountingExportData, kind: str) -> ExportTable:
        raise NotImplementedError
