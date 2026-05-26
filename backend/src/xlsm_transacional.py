"""
Escrita transacional do .xlsm (padrao de mercado: atomic write).

Principios:
  - A base nunca e sobrescrita in-place: o pipeline abre a base e salva num
    arquivo de nome diferente. A base ja e, por construcao, o backup (zero-copy).
  - A versao nova e escrita primeiro num arquivo-parcial e so e "promovida"
    para o nome final via `os.replace` (rename atomico) APOS o save.
  - Colisao de nome (rerun no mesmo dia): se o destino ja existir, ele e movido
    para um arquivo-bak antes do replace; removido apos sucesso, restaurado em
    caso de falha. Rename, nunca copia — barato.
  - A recuperacao autoritativa (kill forcado) e do supervisor Electron, que
    varre os artefatos orfaos. `sweep_orfaos` aqui e o equivalente Python.

Duas restricoes do Excel/xlwings moldam este modulo:
  1. O formato e deduzido pela EXTENSAO -> o marcador vai ANTES dela:
       Gerencie Carteira_2026_05_16.xlsm
         -> parcial: Gerencie Carteira_2026_05_16.partial.xlsm
         -> backup : Gerencie Carteira_2026_05_16.bak.xlsm
  2. Enquanto o workbook estiver aberto, o Excel mantem LOCK no arquivo
     salvo. Por isso a API e dividida em duas fases:
       - `escrever_parcial(wb, destino)`  -> com o Excel AINDA ABERTO
       - `promover(parcial, destino)`     -> APOS wb.close()/app.quit()
     O `promover` ainda faz retry curto porque o Windows (Excel/antivirus)
     pode liberar o handle de forma preguicosa.

Nenhuma escrita destrutiva ocorre antes de existir uma versao nova integra.
"""

import glob
import os
import re
import time

PARTIAL_MARK = "partial"
BAK_MARK = "bak"

# Casa nomes como "<algo>.partial.xlsm" / "<algo>.bak.xlsm" (qualquer extensao)
_ORFAO_RE = re.compile(rf"\.(?:{PARTIAL_MARK}|{BAK_MARK})\.[^.]+$", re.IGNORECASE)


def _artefato(destino: str, marca: str) -> str:
    """Insere a marca ANTES da extensao: a.xlsm -> a.<marca>.xlsm."""
    raiz, ext = os.path.splitext(destino)
    return f"{raiz}.{marca}{ext}"


def _eh_orfao(nome: str) -> bool:
    return bool(_ORFAO_RE.search(nome))


def _replace_com_retry(origem: str, destino: str, *, tentativas: int = 12,
                        intervalo: float = 0.5) -> None:
    """os.replace resiliente ao release preguicoso de lock do Windows."""
    ultimo: OSError | None = None
    for i in range(tentativas):
        try:
            os.replace(origem, destino)
            return
        except PermissionError as e:  # [WinError 32] arquivo em uso
            ultimo = e
            time.sleep(intervalo * (1 + i * 0.25))
    raise ultimo if ultimo else OSError(f"Falha ao mover {origem} -> {destino}")


def sweep_orfaos(pasta: str) -> list[str]:
    """Remove artefatos transacionais orfaos (*.partial.* / *.bak.*). Idempotente."""
    removidos: list[str] = []
    if not os.path.isdir(pasta):
        return removidos
    for nome in os.listdir(pasta):
        if _eh_orfao(nome):
            alvo = os.path.join(pasta, nome)
            try:
                os.remove(alvo)
                removidos.append(alvo)
            except OSError:
                pass
    return removidos


def escrever_parcial(wb, destino: str) -> str:
    """
    Fase 1 (Excel AINDA ABERTO): limpa orfaos e salva o workbook no
    arquivo-parcial (extensao preservada p/ o xlwings). Retorna o caminho
    parcial. Falha aqui => base intacta, destino jamais tocado.
    """
    pasta = os.path.dirname(destino)
    os.makedirs(pasta, exist_ok=True)
    sweep_orfaos(pasta)
    parcial = _artefato(destino, PARTIAL_MARK)
    wb.save(parcial)
    return parcial


def promover(parcial: str, destino: str) -> str:
    """
    Fase 2 (APOS wb.close()/app.quit()): promove o parcial para o nome final
    de forma atomica. Se `destino` ja existir, e movido p/ <dest>.bak.<ext>
    antes do replace (restaurado em caso de falha).
    """
    if not os.path.exists(parcial):
        raise FileNotFoundError(f"Arquivo parcial inexistente: {parcial}")

    backup = _artefato(destino, BAK_MARK)
    tinha_destino = os.path.exists(destino)
    if tinha_destino:
        if os.path.exists(backup):
            os.remove(backup)
        os.replace(destino, backup)

    try:
        _replace_com_retry(parcial, destino)
    except OSError:
        if tinha_destino and os.path.exists(backup):
            os.replace(backup, destino)  # restaura estado anterior
        if os.path.exists(parcial):
            try:
                os.remove(parcial)
            except OSError:
                pass
        raise

    if tinha_destino and os.path.exists(backup):
        os.remove(backup)
    return destino


def limpar_publico_antigos(
    pasta_copia: str,
    glob_antigos: str = "Gerencie*.xls*",
    *,
    tentativas: int = 4,
    intervalo: float = 0.5,
) -> list[tuple[str, OSError]]:
    """
    Remove os arquivos publicos antigos com retry curto. Retorna a lista de
    (caminho, erro) das remocoes que falharam apos `tentativas` — o caller deve
    emit warning para que a equipe veja (tipicamente a planilha esta aberta no
    Excel de outro usuario). Nunca remove artefatos transacionais
    (.partial/.bak).
    """
    falhas: list[tuple[str, OSError]] = []
    for f in glob.glob(os.path.join(pasta_copia, glob_antigos)):
        if _eh_orfao(os.path.basename(f)):
            continue
        ultimo: OSError | None = None
        for i in range(tentativas):
            try:
                os.remove(f)
                ultimo = None
                break
            except OSError as e:
                ultimo = e
                if i < tentativas - 1:
                    time.sleep(intervalo * (1 + i * 0.25))
        if ultimo is not None:
            falhas.append((f, ultimo))
    return falhas
