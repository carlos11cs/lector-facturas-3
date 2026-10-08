from __future__ import annotations

from .profiles.generic import GenericLedgedExportProfile


class ExportProfileRegistry:
    def __init__(self):
        self._profiles = {}

    def register(self, profile):
        if not profile.profile_id or profile.profile_id in self._profiles:
            raise ValueError("Perfil de exportación duplicado o sin identificador.")
        self._profiles[profile.profile_id] = profile

    def get(self, profile_id):
        try:
            return self._profiles[profile_id]
        except KeyError as exc:
            raise KeyError(f"Perfil de exportación desconocido: {profile_id}") from exc


export_profile_registry = ExportProfileRegistry()
export_profile_registry.register(GenericLedgedExportProfile())
