"""
Bootstrap idempotente da estrutura de trabalho do Gerencie Carteira.

Garante que %USERPROFILE%\\Documents\\Gerencie_Carteira\\{planilhas,html,logs}
existe antes do pipeline rodar. A pasta publica (A:\\PUBLICA\\...) e rota de
rede pre-existente — NUNCA criada pelo app, apenas verificada.
"""

import os
from pathlib import Path
from typing import Optional

WORKSPACE_DIRNAME = "Gerencie_Carteira"
SUBDIRS = ("planilhas", "html", "logs")


def resolve_workspace_root(override: Optional[str] = None) -> Path:
    """Retorna o caminho absoluto da raiz do workspace local do usuario."""
    if override:
        return Path(os.path.expandvars(override)).expanduser().resolve()
    return (Path(os.path.expandvars("%USERPROFILE%")) / "Documents" / WORKSPACE_DIRNAME).resolve()


def ensure_workspace(override: Optional[str] = None) -> dict:
    """Cria subpastas faltantes sob a raiz. Idempotente."""
    root = resolve_workspace_root(override)
    root.mkdir(parents=True, exist_ok=True)

    created: list[str] = []
    for sub in SUBDIRS:
        p = root / sub
        if not p.exists():
            p.mkdir(parents=True, exist_ok=True)
            created.append(str(p))

    return {
        "root": str(root),
        "subdirs": [str(root / s) for s in SUBDIRS],
        "created": created,
        "already_existed": len(created) == 0,
    }


def verify_public_path(public_path: str) -> dict:
    """Testa acessibilidade da rota publica (NAO cria)."""
    expanded = os.path.expandvars(public_path)
    exists = os.path.isdir(expanded)
    return {"path": expanded, "accessible": exists}


def detect_legacy_workspace() -> Optional[str]:
    """Detecta pasta legada da v2.14.1 ('Gerencie Carteira' com espaco)."""
    legacy = Path(os.path.expandvars("%USERPROFILE%")) / "Documents" / "Gerencie Carteira"
    return str(legacy) if legacy.exists() else None


def resolve_legacy_planilhas() -> Optional[Path]:
    """Subpasta 'Diário' da v2.14.1, onde ficavam as planilhas diarias."""
    legacy = Path(os.path.expandvars("%USERPROFILE%")) / "Documents" / "Gerencie Carteira" / "Diário"
    return legacy if legacy.exists() else None


if __name__ == "__main__":
    import json
    result = ensure_workspace()
    print(json.dumps(result, indent=2, ensure_ascii=False))
    print("legacy:", detect_legacy_workspace())
    print("publico A:", verify_public_path(r"A:\PUBLICA\GERENCIE CARTEIRA PUBLICA"))
