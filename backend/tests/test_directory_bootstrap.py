import os
from pathlib import Path

from directory_bootstrap import (
    SUBDIRS,
    ensure_workspace,
    resolve_workspace_root,
    verify_public_path,
)


def test_resolve_workspace_default_is_absolute():
    root = resolve_workspace_root()
    assert root.is_absolute()
    assert root.name == "Gerencie_Carteira"


def test_resolve_workspace_with_override(tmp_path):
    root = resolve_workspace_root(str(tmp_path / "custom"))
    assert root == (tmp_path / "custom").resolve()


def test_ensure_workspace_creates_all_subdirs(tmp_path):
    target = tmp_path / "ws"
    result = ensure_workspace(str(target))
    assert result["already_existed"] is False
    assert len(result["created"]) == len(SUBDIRS)
    for sub in SUBDIRS:
        assert (target / sub).is_dir()


def test_ensure_workspace_is_idempotent(tmp_path):
    target = tmp_path / "ws"
    ensure_workspace(str(target))
    second = ensure_workspace(str(target))
    assert second["already_existed"] is True
    assert second["created"] == []


def test_ensure_workspace_partial(tmp_path):
    target = tmp_path / "ws"
    (target / "planilhas").mkdir(parents=True)
    result = ensure_workspace(str(target))
    assert result["already_existed"] is False
    # Apenas html e logs foram criadas (planilhas ja existia)
    created_basenames = {Path(p).name for p in result["created"]}
    assert "planilhas" not in created_basenames
    assert "html" in created_basenames
    assert "logs" in created_basenames


def test_verify_public_path_inaccessible(tmp_path):
    result = verify_public_path(str(tmp_path / "nope"))
    assert result["accessible"] is False


def test_verify_public_path_accessible(tmp_path):
    result = verify_public_path(str(tmp_path))
    assert result["accessible"] is True


def test_verify_public_path_expands_env(monkeypatch, tmp_path):
    monkeypatch.setenv("FAKEROOT", str(tmp_path))
    result = verify_public_path(r"%FAKEROOT%")
    assert result["path"] == str(tmp_path)
    assert result["accessible"] is True
