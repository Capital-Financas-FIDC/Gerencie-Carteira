import json

from log_emitter import emit, emit_result


def _parse_captured(capsys) -> list[dict]:
    out = capsys.readouterr().out.splitlines()
    return [json.loads(line) for line in out]


def test_emit_minimal_fields(capsys):
    emit("info", "mensagem teste")
    events = _parse_captured(capsys)
    assert len(events) == 1
    ev = events[0]
    assert ev["level"] == "info"
    assert ev["msg"] == "mensagem teste"
    assert "ts" in ev
    assert "step" not in ev
    assert "progress" not in ev
    assert "data" not in ev


def test_emit_with_optional_fields(capsys):
    emit("step", "processando", step="excel.save", progress=80, data={"rows": 42})
    ev = _parse_captured(capsys)[0]
    assert ev["step"] == "excel.save"
    assert ev["progress"] == 80
    assert ev["data"] == {"rows": 42}


def test_emit_preserves_acentos(capsys):
    emit("info", "Razão Social com acento e çedilha")
    ev = _parse_captured(capsys)[0]
    assert ev["msg"] == "Razão Social com acento e çedilha"


def test_emit_one_json_per_line(capsys):
    emit("info", "linha 1")
    emit("warning", "linha 2")
    emit("error", "linha 3")
    events = _parse_captured(capsys)
    assert len(events) == 3
    assert [e["level"] for e in events] == ["info", "warning", "error"]


def test_emit_result_ok(capsys):
    emit_result("ok", r"C:\path\planilha.xlsm")
    ev = _parse_captured(capsys)[0]
    assert ev["level"] == "success"
    assert ev["step"] == "done"
    assert ev["data"]["result"]["status"] == "ok"
    assert ev["data"]["result"]["spreadsheet_path"].endswith("planilha.xlsm")


def test_emit_result_warning_and_error(capsys):
    emit_result("warning", None)
    emit_result("error", None)
    events = _parse_captured(capsys)
    assert events[0]["level"] == "warning"
    assert events[1]["level"] == "error"
