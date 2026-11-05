"""Re-bake do .xlam do CalcCP (ver CLAUDE.md, "Adicionar/renomear UDF").

    python rebake.py <xlam-base> <xlam-novo>
    ex.: python rebake.py CalcCP_v8.xlam CalcCP_v9.xlam

1. Abre o CalcCP_DEV.xlam via COM, remove o `xlwings_udfs` e regera as casquinhas (`import_udfs`).
2. Reaplica as 2 alterações de desempenho no VBA (v8.1.0):
   - `XLPy()` monta o comando do Python uma vez por sessão (antes: ~15 leituras de config por célula);
   - as casquinhas passam o VALOR das células (`CpValor`), não o objeto Range.
3. Tira o nome do usuário do Windows do `vbaProject.bin` (o save do Excel o injeta; o repo é público).
4. Transplanta o `vbaProject.bin` para uma cópia do <xlam-base> (que mantém config e docProps limpos).
Precisa de "Confiar no acesso ao modelo de objeto do VBA" ligado e do Excel fechado.
"""
import os
import re
import shutil
import sys
import zipfile

PASTA = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.join(PASTA, "CalcCP_DEV.xlam")

XLPY_ORIGINAL = """Function XLPy()
    XLPyLoadDLL
    If 0 <> XLPyDLLActivateAuto(XLPy, XLPyCommand, 1) Then Err.Raise 1000, Description:=XLPy
End Function"""

XLPY_CACHEADO = """Function XLPy()
    ' CalcCP: o comando do Python (XLPyCommand) e a carga da DLL (XLPyLoadDLL) eram refeitos a CADA
    ' celula calculada -- ~15 GetConfig, cada um relendo a aba xlwings.conf e procurando arquivos .conf
    ' na pasta da planilha e do usuario. O comando nao muda durante a sessao: monta uma vez so.
    ' (ADD_WORKBOOK_TO_PYTHONPATH = false, entao a planilha ativa nao entra no comando.)
    Static comandoPy As String
    If comandoPy = "" Then
        XLPyLoadDLL
        comandoPy = XLPyCommand
    End If
    If 0 <> XLPyDLLActivateAuto(XLPy, comandoPy, 1) Then Err.Raise 1000, Description:=XLPy
End Function"""

CPVALOR = """
' CalcCP: celula unica vai como VALOR (Value2: datas viram o serial numerico, que o CalcCP ja aceita).
' Passar o objeto Range obrigava o Python a voltar ao Excel (chamada COM entre processos, com o Excel
' ocupado calculando) para ler cada argumento de cada celula. Faixa de varias celulas segue como Range.
Private Function CpValor(ByVal v As Variant) As Variant
    If IsObject(v) Then
        If TypeOf v Is Range Then
            If v.Cells.CountLarge = 1 Then
                CpValor = v.Value2
                Exit Function
            End If
        End If
        Set CpValor = v
    Else
        CpValor = v
    End If
End Function
"""

RE_CHAMADA = re.compile(r'(XLPy\.CallUDF\("CalcCP", "(\w+)", )Array\(([^)]*)\)')


def Modulo(comps, nome):
    for i in range(1, comps.Count + 1):
        if comps.Item(i).Name == nome:
            return comps.Item(i)
    return None


def AplicarXlpy(comps):
    cm = Modulo(comps, "xlwings").CodeModule
    codigo = cm.Lines(1, cm.CountOfLines)
    if "Static comandoPy" in codigo:
        return False
    assert codigo.count(XLPY_ORIGINAL.replace("\n", "\r\n")) == 1, "XLPy original do xlwings não encontrado"
    ini = codigo.split("\r\n").index("Function XLPy()") + 1
    cm.DeleteLines(ini, len(XLPY_ORIGINAL.split("\n")))
    cm.InsertLines(ini, XLPY_CACHEADO.replace("\n", "\r\n"))
    return True


def AplicarCpValor(comps):
    cm = Modulo(comps, "xlwings_udfs").CodeModule
    codigo = cm.Lines(1, cm.CountOfLines)

    def Envolver(m):
        args = [a.strip() for a in m.group(3).split(",") if a.strip()]
        return m.group(1) + "Array(" + ", ".join(f"CpValor({a})" for a in args) + ")"

    novo, n = RE_CHAMADA.subn(Envolver, codigo)
    cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(novo.rstrip("\r\n") + "\r\n" + CPVALOR.replace("\n", "\r\n"))
    return n


def Limpar(binario):
    usuario = os.environ.get("USERNAME", "")
    limpo = binario
    if usuario:
        for v in {usuario, usuario.lower(), usuario.upper(), usuario.capitalize()}:
            troca = ("user" + "1" * max(0, len(v) - 4))[:len(v)] if len(v) >= 4 else "x" * len(v)
            limpo = limpo.replace(v.encode(), troca.encode()).replace(v.encode("utf-16-le"), troca.encode("utf-16-le"))
        assert usuario.lower().encode() not in limpo.lower(), "nome do usuário ficou no vbaProject.bin"
    assert len(limpo) == len(binario)
    return limpo


def main(base, novo):
    import xlwings as xw
    from xlwings import udfs

    shutil.copy(DEV, DEV + ".bak")
    sys.path.insert(0, PASTA)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(DEV)
        wb.api.IsAddin = False
        comps = wb.api.VBProject.VBComponents
        antigo = Modulo(comps, "xlwings_udfs")
        if antigo is not None:
            comps.Remove(antigo)
        udfs.import_udfs("CalcCP", wb.api)
        print("XLPy cacheado agora:", AplicarXlpy(comps))
        print("casquinhas com CpValor:", AplicarCpValor(comps))
        cod = Modulo(comps, "xlwings_udfs").CodeModule
        print("UDFs:", cod.Lines(1, cod.CountOfLines).count("Function cp"))
        wb.api.IsAddin = True
        wb.save()
        wb.close()
    finally:
        app.quit()

    binario = Limpar(zipfile.ZipFile(DEV).read("xl/vbaProject.bin"))
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(novo, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, binario if item.filename == "xl/vbaProject.bin" else zin.read(item.filename))
    print("gerado:", novo)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(os.path.join(PASTA, sys.argv[1]), os.path.join(PASTA, sys.argv[2]))
