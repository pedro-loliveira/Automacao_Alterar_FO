import time
import pyautogui
import pygetwindow as gw
from PIL import ImageChops, ImageStat

CONTAGEM = 6
BOTAO = (424, 397, 76, 25)      # região do Incluir que você já mediu


def contagem(msg, segundos=CONTAGEM):
    input(f"\n{msg}\n   ENTER para iniciar a contagem...")
    for s in range(segundos, 0, -1):
        print(f"   capturando em {s}s...", end="\r")
        time.sleep(1)
    print(" " * 30, end="\r")


def esperar_ponto(msg):
    input(f"{msg}  Posicione o mouse e pressione ENTER aqui...")
    time.sleep(0.2)
    p = pyautogui.position()
    print("   ->", (p.x, p.y))
    return p.x, p.y


def cap(r): return pyautogui.screenshot(region=r).convert("L")
def dif(a, b): return ImageStat.Stat(ImageChops.difference(a, b)).mean[0]


def ativa():
    w = gw.getActiveWindow()
    return repr(w.title) if w else "None"


print("Abra a tela de operações com os campos vazios.\n")
x1, y1 = esperar_ponto("Canto SUPERIOR ESQUERDO da caixa do campo OPERAÇÃO.")
x2, y2 = esperar_ponto("Canto INFERIOR DIREITO da caixa do campo OPERAÇÃO.")
R = (x1, y1, x2 - x1, y2 - y1)
print("\nREGIAO_CAMPO_OPERACAO =", R)
neutro = (x1 + 3, y2 + 60)


def foto(nome):
    pyautogui.moveTo(*neutro)
    time.sleep(1.5)
    c, b = cap(R), cap(BOTAO)
    c.save(f"ref_op_{nome}.png")
    b.save(f"ref_botao_{nome}.png")
    print(f"   janela ativa: {ativa()}")
    return c, b


contagem("A) Campos VAZIOS, Operação AZUL. Clique na BARRA DE TÍTULO da janela de operações e espere.")
campoA, botaoA = foto("A_azul")

print("\nB) Tire o foco da janela clicando FORA do ERP (barra de tarefas ou área de trabalho).")
contagem("   Faça isso agora e espere (campo vazio, janela sem foco).", 4)
campoB, botaoB = foto("B_semfoco")

print("\nC) Volte o foco (barra de título), escreva a Operação 0.300 e dê TAB.")
contagem("   Depois de escrever e dar o TAB, pressione ENTER e espere.")
campoC, botaoC = foto("C_preenchido")

print("\nD) Clique em INCLUIR e espere o ERP limpar (Operação azul de novo, com foco).")
contagem("   Depois que limpou, pressione ENTER e espere.")
campoD, botaoD = foto("D_depois")

print("\n--- CAMPO OPERAÇÃO ---")
print(
    f"A azul x D azul depois do Incluir:   {dif(campoA, campoD):.2f}   (deve ser ~0)")
print(
    f"A azul x B sem foco:                 {dif(campoA, campoB):.2f}   (>0 = o estado B é diferente do A)")
print(f"A azul x C preenchido:               {dif(campoA, campoC):.2f}")
print(
    f"B sem foco x C preenchido:           {dif(campoB, campoC):.2f}   (deve ser bem > 0)")
print("--- BOTÃO INCLUIR ---")
print(
    f"A x B (ambos cinza):                 {dif(botaoA, botaoB):.2f}   (deve ser ~0)")
print(
    f"A x D (cinza depois do Incluir):     {dif(botaoA, botaoD):.2f}   (deve ser ~0)")
print(f"A cinza x C preto:                   {dif(botaoA, botaoC):.2f}")
