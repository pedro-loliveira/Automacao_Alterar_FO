# -*- coding: utf-8 -*-
"""
Automação Folha de Operação (ERP-SAI) - PyAutoGUI + Excel

Uso:
    python automacao_folha.py --ler     -> só lê a planilha e mostra o que seria feito (não toca no ERP)
    python automacao_folha.py           -> executa a automação (respeita MODO_TESTE / LIMITE_ITENS)

Dependências:
    pip install pyautogui pygetwindow openpyxl
"""

# =============================================================================
# 1. IMPORTS
# =============================================================================
import os
import re
import sys
import time
import shutil
import tempfile
import logging
from datetime import datetime

import openpyxl
import pyautogui
import pygetwindow as gw

import json
import warnings

from PIL import Image, ImageChops, ImageStat
warnings.filterwarnings("ignore", message="Data Validation extension")

# =============================================================================
# 2. CONFIGURAÇÃO GERAL  (edite aqui)
# =============================================================================
# False desliga toda a checagem por imagem
CONFERENCIA_VISUAL = True
REGIAO_BOTAO_INCLUIR = (424, 397, 76, 25)
# cole a tupla impressa pelo teste: (x, y, largura, altura)
REGIAO_CAMPO_OPERACAO = None
LIMIAR_BOTAO = 6.0      # cinza x preto = 12,35 (ruído 0,00)
# "igual à referência" = diferença <= 3 (ruído 0,00; menor distância entre estados 18,55)
LIMIAR_CAMPO = 3.0
# segundos esperando o estado esperado (sem multiplicador)
TIMEOUT_ESTADO = 40
ARQ_REF_CAMPO_AZUL = "ref_op_A_azul.png"        # Operação azul, janela com foco
ARQ_REF_CAMPO_SEMFOCO = "ref_op_B_semfoco.png"  # Operação vazia, janela sem foco
ARQ_REF_BOTAO_CINZA = "ref_botao_A_azul.png"    # Incluir cinza

CAMINHO_PLANILHA = (
    r"C:\Users\p.deoliveira\OneDrive - Hyva Global\Área de Trabalho\asdasd\Folha_Operação\Analise_Folha_Operação_2s_2026.xlsx"
)
ABA = "Ordem_mais_Produzida"
LINHA_CABECALHO = 2          # linha dos títulos (A..N)
COLUNA_ITEM = "Rótulos de Linha"

# Colunas de subcentros NA ORDEM de cadastro (D -> N). CCD é tratado à parte.
COLUNAS_SUBCENTROS = [
    "CDSE", "CDTE01", "CDPL", "CDPL01", "CDPL02",
    "CDVD01", "CDVD02", "CDCA01", "CDCA02", "CDOX", "CDFR",
]
COLUNAS_MAPEAMENTO = ["KANBAN", "CCD"]
# =============================================================================
# CONFIGURAÇÃO GERAL
# =============================================================================

# ----- Arquivos -----
# cache da planilha; apague para forçar nova leitura
ARQ_CACHE = "cache_planilha.json"

# ----- Títulos das janelas do ERP -----
# tela com Produto e linha_1
TITULO_PRINCIPAL = "Atualiza Folha de Operação"
# tela de Incluir/Excluir operação
TITULO_OPERACOES = "Atualiza Operações da Folha de Operação"
# aviso após Excluir (None = não espera a janela, usa só ESPERA_AVISO_FIXA)
TITULO_AVISO = "Aviso"

# ----- Modo teste -----
# True: pede ENTER antes de cada item, dobra as esperas, usa a seta p/ baixo
MODO_TESTE = False
# nº máximo de itens por execução (só vale em MODO_TESTE; None = todos)
LIMITE_ITENS = 1
# setas p/ baixo antes do Incluir da tela principal (só em MODO_TESTE)
TESTE_SETA_BAIXO = 4
# multiplica TODAS as esperas fixas e o PAUSE
MULT_ESPERA = 1.5 if MODO_TESTE else 0.5

# ----- Segurança -----
pyautogui.FAILSAFE = True       # mouse no canto superior esquerdo aborta o script
CONTAGEM_REGRESSIVA = 5       # segundos de espera antes de começar a clicar
MAX_EXCLUSOES_SEGURANCA = 30    # nunca excluir mais que isso por item

# exclusões extras permitidas se o Incluir não abrir (sobrou linha)
MAX_EXCLUSOES_EXTRAS = 3
# segundos esperando a tela de operações abrir em cada tentativa
TIMEOUT_TENTATIVA_INCLUIR = 8
# folga antes de clicar na linha_1 (a grade precisa terminar de atualizar)
ESPERA_ANTES_LINHA1 = 2.0

# ----- Velocidade: ações do PyAutoGUI -----
# pausa automática depois de CADA ação (clique, tecla, etc.)
pyautogui.PAUSE = 0.8 * MULT_ESPERA
INTERVALO_DIGITACAO = 0.06            # segundos entre cada tecla digitada

# ----- Velocidade: preenchimento dos campos -----
ESPERA_APOS_CAMPO = 0.8         # folga depois de clicar/preencher cada campo
ESPERA_APOS_TAB_CAMPO = 0.8     # folga depois do TAB que libera os campos seguintes
TAB_APOS_CAMPO = {"operacao"}   # campos que exigem TAB depois de digitar
ESPERA_APOS_TAB = 0.5           # folga entre os 3 TABs depois de digitar o Produto

# ----- Velocidade: exclusão -----
# pausa antes de confirmar o aviso (só usada se TITULO_AVISO = None)
ESPERA_AVISO_FIXA = 2.0
# folga depois que a janela principal voltou da exclusão
ESPERA_POS_EXCLUIR = 1.4

# ----- Velocidade: inclusão -----
# tempo do ERP "travar" e limpar os campos depois do Incluir
ESPERA_APOS_INCLUIR = 2.0

# ----- Timeouts e estados das janelas -----
# máximo esperando uma janela abrir/fechar (antes de parar com erro)
TIMEOUT_JANELA = 20
# a tela principal continua aberta depois do Cancelar?
PRINCIPAL_ABERTA_APOS_CANCELAR = True

# cabeçalho da coluna com a REV.FO correta (máximo da consulta)
COLUNA_REV_FO = "1.AFOSC_REV_FO"
# Produto -> TAB (REV) -> TAB (chega na REV.FO)
TABS_ANTES_REV_FO = 2
# depois de digitar a REV.FO, 1 TAB para sair do campo
TABS_APOS_REV_FO = 1
# tempo p/ o ERP validar a REV.FO e mostrar erro, se houver
ESPERA_VALIDACAO_REV = 1.0
# tecla que fecha o aviso de erro (CONFIRMAR)
TECLA_FECHAR_AVISO = "enter"
# False: itens de itens_com_erro.txt são pulados na próxima execução
REPROCESSAR_ERROS = False
ARQ_ERROS = "itens_com_erro.txt"
TITULO_ERRO = "ERRO"

ARQ_CHECKPOINT = "itens_concluidos.txt"
ARQ_LOG = "automacao_folha.log"

PENDENTE = "PREENCHER"

# segundos esperando a tela de exclusão abrir; se não abrir, Excluir está desabilitado
TIMEOUT_ABRIR_EXCLUIR = 5

# =============================================================================
# 3. COORDENADAS  (None = ainda não preenchido; NÃO inventei nenhuma)
#    Dica: rode `python -c "import pyautogui; pyautogui.displayMousePosition()"`
# =============================================================================

# Monitor 2
# COORDENADAS = {
#     # janela principal
#     "produto": (2090, 336),
#     "linha_1": (2173, 455),
#     "excluir_folha": (2514, 832),       # Excluir da tela principal
#     "incluir_folha": (2258, 832),       # Incluir da tela principal
#     "cancelar": (2678, 832),            # Cancelar da tela principal
#     # janela de operações
#     "excluir_operacao": (2547, 613),
#     "ok": (2534, 617),
#     "operacao": (2082, 400),
#     "subcentro": (2182, 468),
#     "qualificacao": (2434, 497),
#     "tempo_execucao": (2487, 526),
#     "numero_operadores": (2628, 555),
#     "incluir_operacao": (2381, 619),
#     "sair": (2670, 276),
# }
# Monitor 1
COORDENADAS = {
    # janela principal
    "produto": (218, 130),
    "linha_1": (203, 249),
    "excluir_folha": (579, 624),       # Excluir da tela principal
    "incluir_folha": (353, 624),       # Incluir da tela principal
    "cancelar": (761, 624),            # Cancelar da tela principal
    # janela de operações
    "excluir_operacao": (627, 410),
    "ok": (932, 586),
    "operacao": (156, 191),
    "subcentro": (243, 263),
    "qualificacao": (513, 291),
    "tempo_execucao": (521, 323),
    "numero_operadores": (715, 347),
    "incluir_operacao": (461, 409),
    "sair": (750, 66),
}

# =============================================================================
# 4. REGRAS DOS SUBCENTROS  (edite aqui)
#    nome_erp  : nome digitado no campo Subcentro (diferente do nome da planilha)
#    operadores: nº de operadores específico de cada subcentro
# =============================================================================
REGRAS_SUBCENTROS = {
    "CDSE":   {"nome_erp": "CDSE", "qualificacao": "05", "operadores": 1},
    "CDTE01": {"nome_erp": "CDTE01", "qualificacao": "05", "operadores": 2},
    "CDPL":   {"nome_erp": "CDPL", "qualificacao": "05", "operadores": 2},
    "CDPL01": {"nome_erp": "CDPL01", "qualificacao": "05", "operadores": 2},
    "CDPL02": {"nome_erp": "CDPL02", "qualificacao": "05", "operadores": 2},
    "CDVD01": {"nome_erp": "CDVD01", "qualificacao": "05", "operadores": 2},
    "CDVD02": {"nome_erp": "CDVD02", "qualificacao": "05", "operadores": 1},
    "CDCA01": {"nome_erp": "CDCA01", "qualificacao": "05", "operadores": 2},
    "CDCA02": {"nome_erp": "CDCA02", "qualificacao": "05", "operadores": 2},
    "CDOX":   {"nome_erp": "CDOX", "qualificacao": "05", "operadores": 1},
    "CDFR":   {"nome_erp": "CDFR", "qualificacao": "05", "operadores": 1},
}

# Regra fixa do CCD
CCD_OPERACAO = "0.100"
CCD_SUBCENTRO = "CCD"
CCD_QUALIFICACAO = "05"


# =============================================================================
# 5. LOG
# =============================================================================
def _criar_log():
    lg = logging.getLogger("folha")
    lg.setLevel(logging.INFO)
    fmt = logging.Formatter("[%(asctime)s] %(message)s", "%H:%M:%S")
    for h in (logging.StreamHandler(sys.stdout),
              logging.FileHandler(ARQ_LOG, encoding="utf-8")):
        h.setFormatter(fmt)
        lg.addHandler(h)
    return lg


log = _criar_log()


class ErroAutomacao(Exception):
    """Erro com contexto (item / subcentro / etapa)."""

    def __init__(self, item, subcentro, etapa, erro):
        self.item, self.subcentro, self.etapa, self.erro = item, subcentro, etapa, erro
        super().__init__(
            f"Item: {item} | Subcentro: {subcentro} | Etapa: {etapa} | Erro: {erro}")


# =============================================================================
# 6. LEITURA DA PLANILHA
# =============================================================================
PADRAO_ITEM = re.compile(r"^\d+(\.\d+)+$")   # ex.: 1.83.1.02.3342


def converter_tempo(valor):
    """Converte valor da célula em float; retorna None se for '-', vazio, 0 ou inválido."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        n = float(valor)
    else:
        txt = str(valor).strip()
        if txt in ("", "-", "nan", "NaN", "None"):
            return None
        try:
            n = float(txt.replace(",", "."))
        except ValueError:
            return None
    n = round(n, 2)
    return n if n > 0 else None


def formatar_tempo(n):
    """0.02 -> '0,02' (2 casas, vírgula)."""
    return f"{n:.2f}".replace(".", ",")


def _assinatura_arquivo():
    st = os.stat(CAMINHO_PLANILHA)
    return {"mtime": st.st_mtime, "size": st.st_size, "aba": ABA}


def tem_valor(valor):
    """True se a célula tem algum valor (não vazio e diferente de '-'). Sem arredondar."""
    if valor is None:
        return False
    return str(valor).strip() not in ("", "-", "nan", "NaN", "None")


def _assinatura_arquivo():
    st = os.stat(CAMINHO_PLANILHA)
    return {"mtime": st.st_mtime, "size": st.st_size, "aba": ABA, "versao": 5}


def ler_planilha():
    """Retorna itens: {'codigo', 'subcentros': [(coluna, tempo)], 'existentes_extras': [KANBAN/CCD presentes]}."""
    if not os.path.exists(CAMINHO_PLANILHA):
        raise FileNotFoundError(f"Planilha não encontrada: {CAMINHO_PLANILHA}")

    # ---- cache ----
    assinatura = _assinatura_arquivo()
    if os.path.exists(ARQ_CACHE):
        try:
            with open(ARQ_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
            if cache.get("assinatura") == assinatura:
                itens = [{"codigo": i["codigo"],
                          "subcentros": [tuple(s) for s in i["subcentros"]],
                          "existentes_extras": i["existentes_extras"],
                          "rev_fo": i.get("rev_fo")}
                         for i in cache["itens"]]
                log.info(f"Planilha carregada do cache. Itens válidos: {len(itens)} "
                         f"(apague {ARQ_CACHE} para reler)")
                return itens
        except Exception as e:
            log.info(f"Cache inválido, relendo a planilha ({e})")

    # ---- leitura real ----
    t0 = time.time()
    tmp = os.path.join(tempfile.gettempdir(), "folha_tmp.xlsx")
    shutil.copy2(CAMINHO_PLANILHA, tmp)
    log.info(f"Planilha copiada em {time.time() - t0:.1f}s")

    t1 = time.time()
    wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
    if ABA not in wb.sheetnames:
        raise ValueError(f"Aba '{ABA}' não encontrada. Abas: {wb.sheetnames}")
    ws = wb[ABA]
    ws.reset_dimensions()
    log.info(f"Aba encontrada: {ABA}")

    linhas = list(ws.iter_rows(values_only=True))
    wb.close()
    log.info(
        f"Planilha lida em {time.time() - t1:.1f}s ({len(linhas)} linhas)")

    cab = {}
    for idx, v in enumerate(linhas[LINHA_CABECALHO - 1]):
        if v is not None:
            cab[str(v).strip()] = idx
    faltando = [n for n in [COLUNA_ITEM, COLUNA_REV_FO] + COLUNAS_SUBCENTROS + COLUNAS_MAPEAMENTO
                if n not in cab]
    if faltando:
        raise ValueError(
            f"Cabeçalhos não encontrados na linha {LINHA_CABECALHO}: {faltando}")

    def celula(linha, nome):
        i = cab[nome]
        return linha[i] if len(linha) > i else None

    itens, ignoradas = [], 0
    for linha in linhas[LINHA_CABECALHO:]:
        bruto = celula(linha, COLUNA_ITEM)
        codigo = str(bruto).strip() if bruto is not None else ""
        if not PADRAO_ITEM.match(codigo):
            if codigo:
                ignoradas += 1
            continue
        subs = []
        for nome in COLUNAS_SUBCENTROS:
            t = converter_tempo(celula(linha, nome))
            if t is not None:
                subs.append((nome, t))
        extras = [n for n in COLUNAS_MAPEAMENTO if tem_valor(celula(linha, n))]
        rev = celula(linha, COLUNA_REV_FO)
        try:
            n_rev = int(float(str(rev).replace(",", "."))
                        ) if tem_valor(rev) else None
        except ValueError:
            n_rev = None
        # None = deixar o ERP preencher; "0" é digitado
        rev_fo = str(n_rev) if n_rev is not None else None
        itens.append({"codigo": codigo, "subcentros": subs,
                      "existentes_extras": extras, "rev_fo": rev_fo})

    with open(ARQ_CACHE, "w", encoding="utf-8") as f:
        json.dump({"assinatura": assinatura, "itens": itens},
                  f, ensure_ascii=False)

    log.info(
        f"Planilha carregada. Itens válidos: {len(itens)} | linhas ignoradas: {ignoradas}")
    return itens


# =============================================================================
# 7. REGRAS DE NEGÓCIO
# =============================================================================
def obter_regras_subcentro(coluna):
    if coluna not in REGRAS_SUBCENTROS:
        raise KeyError(f"Subcentro '{coluna}' sem regra em REGRAS_SUBCENTROS")
    return REGRAS_SUBCENTROS[coluna]


def montar_operacoes_item(item):
    """Gera a lista de operações do item na ordem de cadastro (subcentros + CCD no fim)."""
    subs = item["subcentros"]
    if len(subs) >= 10:
        # 10º subcentro seria 0.100, que é o código fixo do CCD
        raise ErroAutomacao(item["codigo"], "-", "montar operações",
                            f"{len(subs)} subcentros: o 10º colidiria com a operação 0.100 do CCD")
    ops, soma_pond, soma_op = [], 0.0, 0
    for i, (coluna, tempo) in enumerate(subs, start=1):
        r = obter_regras_subcentro(coluna)
        ops.append({
            "coluna": coluna,
            "operacao": f"0.{i * 10:03d}",          # 0.010, 0.020, ...
            "subcentro": r["nome_erp"],
            "qualificacao": r["qualificacao"],
            "tempo": formatar_tempo(tempo),
            "operadores": str(r["operadores"]),
        })
        soma_pond += tempo * int(r["operadores"])
        soma_op += int(r["operadores"])
    if ops:
        ops.append(calcular_ccd(soma_pond, soma_op))
    return ops


def calcular_ccd(soma_tempo_ponderado, soma_operadores):
    """CCD: tempo = soma(tempo x operadores de cada subcentro); operadores = soma dos operadores."""
    return {
        "coluna": "CCD",
        "operacao": CCD_OPERACAO,
        "subcentro": CCD_SUBCENTRO,
        "qualificacao": CCD_QUALIFICACAO,
        "tempo": formatar_tempo(round(soma_tempo_ponderado, 2)),
        "operadores": str(soma_operadores),
    }


def validar_config_para_executar(itens):
    """Para ANTES de mexer no ERP se faltar coordenada ou regra."""
    problemas = [f"coordenada '{k}'" for k,
                 v in COORDENADAS.items() if v is None]
    usados = {c for it in itens for c, _ in it["subcentros"]}
    for c in sorted(usados):
        r = REGRAS_SUBCENTROS.get(c)
        if r is None:
            problemas.append(f"regra do subcentro {c} inexistente")
            continue
        for campo, valor in r.items():
            if valor == PENDENTE:
                problemas.append(f"regra {c}.{campo}")
    if problemas:
        raise SystemExit("Configuração incompleta, preencha antes de executar:\n  - "
                         + "\n  - ".join(problemas))


# =============================================================================
# 8. JANELAS E ESPERAS
# =============================================================================
_REFS = {}


def _dif(a, b):
    return ImageStat.Stat(ImageChops.difference(a, b)).mean[0]


def _visual_ativo():
    return bool(CONFERENCIA_VISUAL and REGIAO_BOTAO_INCLUIR and REGIAO_CAMPO_OPERACAO)


def _carregar_refs():
    if _REFS:
        return
    for chave, arq, regiao in (("campo_azul", ARQ_REF_CAMPO_AZUL, REGIAO_CAMPO_OPERACAO),
                               ("campo_semfoco", ARQ_REF_CAMPO_SEMFOCO,
                                REGIAO_CAMPO_OPERACAO),
                               ("botao_cinza", ARQ_REF_BOTAO_CINZA, REGIAO_BOTAO_INCLUIR)):
        if not os.path.exists(arq):
            raise FileNotFoundError(
                f"Imagem de referência não encontrada: {arq}")
        img = Image.open(arq).convert("L")
        if img.size != (regiao[2], regiao[3]):
            raise ValueError(f"{arq} tem tamanho {img.size}, mas a região configurada é {regiao[2:]}. "
                             f"Use a mesma região com que as imagens foram salvas.")
        _REFS[chave] = img


def _medir_estado():
    """Classifica o que está na tela: PRONTO_FOCO, PRONTO_SEM_FOCO, PREENCHIDO ou DESCONHECIDO."""
    campo = pyautogui.screenshot(region=REGIAO_CAMPO_OPERACAO).convert("L")
    botao = pyautogui.screenshot(region=REGIAO_BOTAO_INCLUIR).convert("L")
    d_azul = _dif(campo, _REFS["campo_azul"])
    d_sem = _dif(campo, _REFS["campo_semfoco"])
    d_bot = _dif(botao, _REFS["botao_cinza"])
    cinza = d_bot <= LIMIAR_BOTAO
    if cinza and d_azul <= LIMIAR_CAMPO:
        estado = "PRONTO_FOCO"          # Operação azul + Incluir cinza
    elif cinza and d_sem <= LIMIAR_CAMPO:
        estado = "PRONTO_SEM_FOCO"      # Operação vazia sem azul + Incluir cinza
    elif not cinza and d_azul > LIMIAR_CAMPO and d_sem > LIMIAR_CAMPO:
        # Incluir preto + Operação diferente dos dois estados vazios
        estado = "PREENCHIDO"
    else:
        estado = "DESCONHECIDO"
    return estado, d_azul, d_sem, d_bot


def _esperar_estado(aceitos, descricao):
    """Espera o estado da tela estar em 'aceitos' por duas leituras seguidas. Estoura com erro se não chegar."""
    if not _visual_ativo():
        return
    _carregar_refs()
    limite = time.time() + TIMEOUT_ESTADO
    seguidas, ultimo = 0, ("-", 0, 0, 0)
    while time.time() < limite:
        if janela_aberta(TITULO_ERRO):
            raise RuntimeError(f"Janela 'ERRO' apareceu {descricao}")
        ultimo = _medir_estado()
        if ultimo[0] in aceitos:
            seguidas += 1
            if seguidas >= 2:           # duas leituras seguidas, para não pegar um repaint no meio
                log.info(f"Estado {ultimo[0]} (campo azul {ultimo[1]:.2f} / sem foco {ultimo[2]:.2f} "
                         f"/ botão {ultimo[3]:.2f})")
                return
        else:
            seguidas = 0
        time.sleep(0.3)
    raise RuntimeError(f"Estado esperado {sorted(aceitos)} não confirmado em {TIMEOUT_ESTADO}s {descricao}. "
                       f"Último: {ultimo[0]} (campo azul {ultimo[1]:.2f} / sem foco {ultimo[2]:.2f} / botão {ultimo[3]:.2f}). "
                       f"ERP travado, clique perdido ou algo cobrindo a tela")


def esperar_pronto(descricao=""):
    _esperar_estado({"PRONTO_FOCO", "PRONTO_SEM_FOCO"}, descricao)


def esperar_preenchido(descricao=""):
    _esperar_estado({"PREENCHIDO"}, descricao)


def janela_aberta(titulo):
    return any(w.title.strip() == titulo for w in gw.getWindowsWithTitle(titulo))


def verificar_janela():
    """Garante que a janela principal existe e tenta trazê-la para a frente."""
    j = [w for w in gw.getWindowsWithTitle(TITULO_PRINCIPAL)
         if w.title.strip() == TITULO_PRINCIPAL]
    if not j:
        raise RuntimeError(f"Janela '{TITULO_PRINCIPAL}' não encontrada")
    try:
        j[0].activate()
    except Exception:
        pass
    return True


def esperar_janela(titulo, aberta=True, timeout=TIMEOUT_JANELA, descricao=""):
    """Espera a janela com esse título abrir (aberta=True) ou fechar (aberta=False)."""
    limite = time.time() + timeout * MULT_ESPERA
    while time.time() < limite:
        if janela_aberta(titulo) == aberta:
            return True
        time.sleep(0.3)
    estado = "abrir" if aberta else "fechar"
    raise RuntimeError(f"Timeout esperando '{titulo}' {estado} {descricao}")


def pausa(seg):
    time.sleep(seg * MULT_ESPERA)


# =============================================================================
# 9. AÇÕES PYAUTOGUI
# =============================================================================
def clicar(nome):
    x, y = COORDENADAS[nome]
    pyautogui.click(x, y)


def duplo_clicar(nome):
    x, y = COORDENADAS[nome]
    pyautogui.doubleClick(x, y)


def escrever_em(nome_campo, texto):
    """Clica no campo, limpa, digita e, se o campo exigir, aperta TAB."""
    clicar(nome_campo)
    pausa(ESPERA_APOS_CAMPO)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(texto), interval=INTERVALO_DIGITACAO)
    if nome_campo in TAB_APOS_CAMPO:
        pyautogui.press("tab")
        pausa(ESPERA_APOS_TAB_CAMPO)
    else:
        pausa(ESPERA_APOS_CAMPO)


# ==========================================0027===================================
# 10. ETAPAS DO PROCESSO
# =============================================================================
def selecionar_produto(item):
    # rev_fo None = deixar o ERP preencher
    codigo, rev_fo = item["codigo"], item.get("rev_fo")

    verificar_janela()
    esperar_janela(TITULO_PRINCIPAL, True,
                   descricao="(antes de selecionar produto)")
    esperar_janela(TITULO_OPERACOES, False,
                   descricao="(tela de operações já aberta antes do produto)")
    # duplo clique seleciona o código antigo inteiro
    duplo_clicar("produto")
    pausa(ESPERA_APOS_CAMPO)
    pyautogui.write(codigo, interval=INTERVALO_DIGITACAO)

    for _ in range(TABS_ANTES_REV_FO):          # ERP preenche REV e o foco chega na REV.FO
        pyautogui.press("tab")
        pausa(ESPERA_APOS_TAB)

    if rev_fo is not None:
        pyautogui.write(rev_fo, interval=INTERVALO_DIGITACAO)   # inclui "0"
    else:
        log.info(
            f"REV.FO vazia na planilha: mantendo a que o ERP preencheu ({codigo})")

    for _ in range(TABS_APOS_REV_FO):
        pyautogui.press("tab")
        pausa(ESPERA_APOS_TAB)

    # Espera até ESPERA_VALIDACAO_REV segundos pela janela de ERRO
    limite = time.time() + ESPERA_VALIDACAO_REV * MULT_ESPERA
    while time.time() < limite:
        if janela_aberta(TITULO_ERRO):
            pyautogui.press(TECLA_FECHAR_AVISO)
            pausa(ESPERA_APOS_CAMPO)
            if janela_aberta(TITULO_ERRO):
                raise RuntimeError(
                    "Janela 'ERRO' não fechou com a tecla configurada")
            clicar("cancelar")
            pausa(ESPERA_POS_EXCLUIR)
            raise ItemComErro(
                codigo, f"REV.FO {rev_fo if rev_fo is not None else '(ERP)'} recusada pelo ERP")
        time.sleep(0.2)


def excluir_operacao():
    """Um ciclo de exclusão. Só retorna quando voltou para a janela principal."""
    pausa(ESPERA_ANTES_LINHA1)       # o ERP é lento para redesenhar a grade depois da exclusão anterior
    clicar("linha_1")
    clicar("excluir_folha")
    esperar_janela(TITULO_OPERACOES, True,
                   descricao="(abrir tela de exclusão)")
    clicar("excluir_operacao")

    # Espera o aviso abrir ANTES de confirmar, senão o ERP trava
    if TITULO_AVISO:
        esperar_janela(TITULO_AVISO, True,
                       descricao="(abrir aviso de confirmação)")
        pausa(0.5)   # folga para o botão ficar clicável
    else:
        pausa(ESPERA_AVISO_FIXA)

    clicar("ok")

    if TITULO_AVISO:
        esperar_janela(TITULO_AVISO, False,
                       descricao="(fechar aviso após confirmar)")
    esperar_janela(TITULO_OPERACOES, False, descricao="(fechar tela após OK)")
    esperar_janela(TITULO_PRINCIPAL, True,
                   descricao="(voltar à janela principal)")
    pausa(ESPERA_POS_EXCLUIR)


def tentar_excluir_operacao():
    """Um ciclo de exclusão. Retorna False se o Excluir não abriu a janela (grade vazia ou ERP ainda atualizando)."""
    pausa(ESPERA_ANTES_LINHA1)       # o ERP é lento para redesenhar a grade
    clicar("linha_1")
    clicar("excluir_folha")
    try:
        esperar_janela(TITULO_OPERACOES, True, timeout=TIMEOUT_ABRIR_EXCLUIR,
                       descricao="(abrir tela de exclusão)")
    except RuntimeError:
        return False

    clicar("excluir_operacao")

    # Espera o aviso abrir ANTES de confirmar, senão o ERP trava
    if TITULO_AVISO:
        esperar_janela(TITULO_AVISO, True,
                       descricao="(abrir aviso de confirmação)")
        pausa(0.5)
    else:
        pausa(ESPERA_AVISO_FIXA)

    clicar("ok")

    if TITULO_AVISO:
        esperar_janela(TITULO_AVISO, False,
                       descricao="(fechar aviso após confirmar)")
    esperar_janela(TITULO_OPERACOES, False, descricao="(fechar tela após OK)")
    esperar_janela(TITULO_PRINCIPAL, True,
                   descricao="(voltar à janela principal)")
    pausa(ESPERA_POS_EXCLUIR)
    return True


def excluir_todas_operacoes(item):
    """Faz as exclusões esperadas (subcentros + KANBAN/CCD). Para antes se o Excluir deixar de abrir."""
    qtd_subs = len(item["subcentros"])
    extras = item["existentes_extras"]
    total = qtd_subs + len(extras)
    if total > MAX_EXCLUSOES_SEGURANCA:
        raise RuntimeError(
            f"{total} exclusões excede o limite de segurança ({MAX_EXCLUSOES_SEGURANCA})")
    log.info(
        f"Iniciando exclusão: {qtd_subs} subcentros + extras {extras or 'nenhum'} = {total}")
    for i in range(1, total + 1):
        log.info(f"Exclusão {i}/{total}")
        if not tentar_excluir_operacao():
            log.warning(f"Excluir não abriu na exclusão {i}/{total}: grade provavelmente já vazia "
                        f"(a planilha esperava mais operações que o ERP tinha)")
            break


def tentar_abrir_inclusao(primeira_tentativa):
    """Clica na linha 1 e em Incluir. True se a tela de operações abriu (grade vazia)."""
    clicar("linha_1")
    if MODO_TESTE and TESTE_SETA_BAIXO and primeira_tentativa:
        log.info(f"TESTE: pressionando seta para baixo {TESTE_SETA_BAIXO}x")
        for _ in range(TESTE_SETA_BAIXO):
            pyautogui.press("down")
    clicar("incluir_folha")
    try:
        esperar_janela(TITULO_OPERACOES, True, timeout=TIMEOUT_TENTATIVA_INCLUIR,
                       descricao="(abrir tela de operações)")
    except RuntimeError:
        return False
    # deixa a tela terminar de desenhar
    pausa(ESPERA_APOS_CAMPO * 2)
    esperar_pronto("(tela de operações recém-aberta)")
    return True


def iniciar_inclusao():
    # folga: o ERP trava depois da última exclusão
    pausa(ESPERA_POS_EXCLUIR * 2)
    for extra in range(MAX_EXCLUSOES_EXTRAS + 1):
        if tentar_abrir_inclusao(primeira_tentativa=(extra == 0)):
            if extra:
                log.warning(
                    f"Incluir só abriu depois de {extra} tentativa(s) extra(s)")
            return
        if extra == MAX_EXCLUSOES_EXTRAS:
            break
        log.warning(
            f"Incluir não abriu (sobrou linha ou ERP lento). Tentativa extra {extra + 1}/{MAX_EXCLUSOES_EXTRAS}")
        if tentar_excluir_operacao():
            log.warning(
                "Uma exclusão extra foi feita: a contagem da planilha ficou menor que o ERP")
        else:
            # Excluir não abriu: ERP ainda atualizando, dá mais tempo
            pausa(ESPERA_POS_EXCLUIR * 2)
    raise RuntimeError(
        f"Incluir não abriu mesmo após {MAX_EXCLUSOES_EXTRAS} tentativas extras")


def preencher_operacao_generica(op):
    """Preenchimento comum a subcentros e CCD (a tela é a mesma)."""
    escrever_em("operacao", op["operacao"])
    esperar_preenchido(
        "(depois de preencher Operação: o ERP não liberou os demais campos)")
    escrever_em("subcentro", op["subcentro"])
    escrever_em("qualificacao", op["qualificacao"])
    escrever_em("tempo_execucao", op["tempo"])
    escrever_em("numero_operadores", op["operadores"])


def incluir_operacao():
    clicar("incluir_operacao")
    pausa(ESPERA_APOS_INCLUIR)
    esperar_janela(TITULO_OPERACOES, True,
                   descricao="(após Incluir, a tela deveria continuar aberta)")
    esperar_pronto("(depois do Incluir: o ERP não terminou de limpar)")


def processar_subcentro(op):
    preencher_operacao_generica(op)
    incluir_operacao()


def processar_ccd(op):
    # Lógica própria do CCD fica em calcular_ccd(); aqui só a execução.
    preencher_operacao_generica(op)
    incluir_operacao()


def finalizar_item():
    clicar("sair")
    esperar_janela(TITULO_OPERACOES, False, descricao="(após Sair)")
    esperar_janela(TITULO_PRINCIPAL, True,
                   descricao="(voltar à principal após Sair)")
    clicar("cancelar")
    pausa(ESPERA_POS_EXCLUIR)
    if not PRINCIPAL_ABERTA_APOS_CANCELAR:
        esperar_janela(TITULO_PRINCIPAL, False, descricao="(após Cancelar)")


# =============================================================================
# 11. CHECKPOINT
# =============================================================================
def carregar_concluidos():
    if not os.path.exists(ARQ_CHECKPOINT):
        return set()
    with open(ARQ_CHECKPOINT, encoding="utf-8") as f:
        return {l.strip() for l in f if l.strip()}


def registrar_concluido(codigo):
    with open(ARQ_CHECKPOINT, "a", encoding="utf-8") as f:
        f.write(codigo + "\n")


class ItemComErro(Exception):
    """Erro recuperável: o item não foi alterado e o script pode seguir para o próximo."""

    def __init__(self, item, tipo):
        self.item, self.tipo = item, tipo
        super().__init__(f"{item} - {tipo}")


def registrar_erro(codigo, tipo):
    with open(ARQ_ERROS, "a", encoding="utf-8") as f:
        f.write(f"{codigo} - {tipo}\n")


def carregar_com_erro():
    if REPROCESSAR_ERROS or not os.path.exists(ARQ_ERROS):
        return set()
    with open(ARQ_ERROS, encoding="utf-8") as f:
        return {l.split(" - ")[0].strip() for l in f if l.strip()}
# =============================================================================
# 12. PROCESSAMENTO DE UM ITEM
# =============================================================================


def processar_item(item):
    codigo = item["codigo"]
    etapa, atual = "início", "-"
    try:
        ops = montar_operacoes_item(item)
        log.info(f"Processando item: {codigo} (REV.FO {item.get('rev_fo')})")
        log.info("Subcentros válidos: " +
                 ", ".join(c for c, _ in item["subcentros"]))

        etapa = "selecionar produto"
        # recebe o item inteiro (precisa da REV.FO)
        selecionar_produto(item)

        etapa = "exclusão"
        # recebe o item inteiro (conta KANBAN/CCD)
        excluir_todas_operacoes(item)

        etapa = "iniciar inclusão"
        iniciar_inclusao()

        log.info("Iniciando inclusão")
        etapa = "inclusão"
        for op in ops:
            atual = op["coluna"]
            log.info(f"Incluindo {op['coluna']} (op {op['operacao']}, sub {op['subcentro']}, "
                     f"tempo {op['tempo']}, operadores {op['operadores']})")
            if op["coluna"] == "CCD":
                processar_ccd(op)
            else:
                processar_subcentro(op)

        atual, etapa = "-", "finalizar item"
        finalizar_item()
        registrar_concluido(codigo)
        log.info("Item finalizado\n")
    except (ErroAutomacao, ItemComErro):
        raise                              # ItemComErro é tratado no main, sem parar o script
    except Exception as e:
        raise ErroAutomacao(codigo, atual, etapa, repr(e)) from e


# =============================================================================
# 13. MAIN
# =============================================================================
def modo_leitura(itens):
    log.info(f"Total de itens: {len(itens)}")
    for it in itens[: (LIMITE_ITENS or len(itens)) if MODO_TESTE else len(itens)]:
        total_exc = len(it["subcentros"]) + len(it["existentes_extras"])
        print(f"\nItem {it['codigo']}: {len(it['subcentros'])} subcentros, "
              f"extras {it['existentes_extras'] or 'nenhum'} -> {total_exc} exclusões")
        try:
            # Para pré-visualizar sem exigir regras, usa operadores 1 se pendente
            subs = it["subcentros"]
            soma_t = sum(
                t * REGRAS_SUBCENTROS[c]["operadores"] for c, t in subs)
            for i, (c, t) in enumerate(subs, start=1):
                print(f"   0.{i * 10:03d}  {c:7s} tempo {formatar_tempo(t)}")
            ops_def = [REGRAS_SUBCENTROS[c]["operadores"] for c, _ in subs]
            if all(isinstance(o, int) for o in ops_def):
                ccd = calcular_ccd(soma_t, sum(ops_def))
                print(
                    f"   {ccd['operacao']}  CCD     tempo {ccd['tempo']}  operadores {ccd['operadores']}")
            else:
                print("   0.100  CCD     (operadores pendentes nas regras)")
        except Exception as e:
            print("   erro na pré-visualização:", e)


def main():
    log.info("Iniciando automação")
    itens = ler_planilha()

    if "--ler" in sys.argv:
        modo_leitura(itens)
        return

    validar_config_para_executar(itens)

    concluidos = carregar_concluidos()
    com_erro = carregar_com_erro()
    pendentes = []
    for i in itens:
        if i["codigo"] in concluidos:
            continue
        if i["codigo"] in com_erro:
            log.warning(f"PULADO (erro em execução anterior): {i['codigo']}")
            continue
        if not i["subcentros"]:
            log.warning(
                f"PULADO (sem subcentro válido): {i['codigo']} | extras: {i['existentes_extras']}")
            continue
        pendentes.append(i)

    if MODO_TESTE and LIMITE_ITENS:
        pendentes = pendentes[:LIMITE_ITENS]
    log.info(
        f"Itens a processar: {len(pendentes)} (já concluídos antes: {len(concluidos)})")
    log.info("Próximos itens: " +
             ", ".join(i["codigo"] for i in pendentes[:5]))

    log.info(
        f"Começando em {CONTAGEM_REGRESSIVA}s. Mouse no canto superior esquerdo = ABORTAR.")
    time.sleep(CONTAGEM_REGRESSIVA)

    try:
        for item in pendentes:
            if MODO_TESTE:
                input(
                    f"\n[TESTE] Pronto para processar {item['codigo']}. ENTER para continuar (Ctrl+C cancela)...")
            try:
                processar_item(item)
            except ItemComErro as e:
                registrar_erro(e.item, e.tipo)
                log.warning(f"ERRO RECUPERÁVEL, item pulado: {e}\n")
        log.info("Automação concluída")
    except pyautogui.FailSafeException:
        log.error(
            "ABORTADO pelo fail-safe (mouse no canto). Verifique o ERP: o item pode estar incompleto.")
    except KeyboardInterrupt:
        log.error(
            "Interrompido pelo usuário. Verifique o ERP: o item pode estar incompleto.")
    except ErroAutomacao as e:
        log.error(
            f"\nItem: {e.item}\nSubcentro: {e.subcentro}\nEtapa: {e.etapa}\nErro: {e.erro}")
        log.error(
            "PARADO com segurança. O item acima pode estar incompleto no ERP, confira manualmente.")


if __name__ == "__main__":
    main()
