"""Gera o AntonioOliveiraCalc_bundle.py (auto-extrator self-contained) a partir dos
arquivos atuais do add-in. Rode SEMPRE que mudar o .py, apis.py, di.py, config.py,
o .csv ou o .xlam (nova versão):

    python gerar_bundle.py

O bundle embute cada arquivo em base64 e, ao ser rodado, escreve tudo na pasta atual.
Inclui o .xlam VERSIONADO (AntonioOliveiraCalc_v{VERSION}.xlam)."""
import base64
import os
import re

AQUI = os.path.dirname(os.path.abspath(__file__))


def _versao():
    txt = open(os.path.join(AQUI, "AntonioOliveiraCalc.py"), encoding="utf-8").read()
    m = re.search(r'VERSION\s*=\s*"([^"]+)"', txt)
    return m.group(1) if m else "0.0.0"


VERSION = _versao()
XLAM = f"AntonioOliveiraCalc_v{VERSION.split('.')[0]}.xlam"   # ex.: _v1.xlam

# Arquivos que vão pra share (Z:\AntonioOliveira\AntonioOliveiraCalc):
ARQUIVOS = [
    XLAM,
    "AntonioOliveiraCalc.py",
    "apis.py",
    "di.py",
    "config.py",
    "feriados_anbima.csv",
]

CABECALHO = f'''"""
AntonioOliveiraCalc — AUTO-EXTRATOR (bundle self-contained) — v{VERSION}
=======================================================================
Baixe SÓ este arquivo, jogue na pasta destino e rode:

    python AntonioOliveiraCalc_bundle.py

Escreve todos os arquivos do add-in na pasta atual (ou passe outra pasta como arg).
Os .py + o .csv precisam ficar no PYTHONPATH que o .xlam espera:
    Z:/AntonioOliveira/AntonioOliveiraCalc   (baked no .xlam, com barra invertida)
Se a share do banco tiver outro caminho, o .xlam precisa ser re-baked com o path certo.

O .xlam é VERSIONADO ({XLAM}) — instalar/registrar ESTE arquivo. Versões antigas
(_v1, _v2...) ficam na share e nunca são apagadas (não quebram quem já usa).

Todas as UDFs têm prefixo cp (=cpPu, =cpTaxa, =cpFluxo, =cpFluxoCompleto, =cpSelic ...).
Sem dependências externas (só a stdlib). Não contém segredos.
"""
import base64, os, sys

DEST = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))

FILES = {{
'''

RODAPE = '''}

def main():
    os.makedirs(DEST, exist_ok=True)
    for nome, b64 in FILES.items():
        caminho = os.path.join(DEST, nome)
        with open(caminho, "wb") as f:
            f.write(base64.b64decode(b64))
        print("escrito:", caminho)
    print(f"\\nOK — {len(FILES)} arquivos em {DEST}")
    print("Lembre: essa pasta precisa ser Z:\\\\AntonioOliveira\\\\AntonioOliveiraCalc "
          "(ou re-bake o .xlam com o path certo).")

if __name__ == "__main__":
    main()
'''


def main():
    partes = [CABECALHO]
    for nome in ARQUIVOS:
        caminho = os.path.join(AQUI, nome)
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"faltando: {nome}")
        b64 = base64.b64encode(open(caminho, "rb").read()).decode("ascii")
        partes.append(f"    {nome!r}: {b64!r},\n")
    partes.append(RODAPE)
    saida = os.path.join(AQUI, "AntonioOliveiraCalc_bundle.py")
    with open(saida, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join(partes))
    print(f"bundle gerado: {saida}  (v{VERSION}, {len(ARQUIVOS)} arquivos, xlam={XLAM})")


if __name__ == "__main__":
    main()
