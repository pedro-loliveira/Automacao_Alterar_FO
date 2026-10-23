import time
import pyautogui
import automacao_folha as a

CAMPO = "tempo_execucao"
VALOR = "0,03"


def m_backspace():
    a.clicar(CAMPO)
    pyautogui.press("end")
    pyautogui.press("backspace", presses=25, interval=0.02)
    pyautogui.press("delete", presses=25, interval=0.02)


def m_home_shift_end():
    a.clicar(CAMPO)
    pyautogui.press("home")
    pyautogui.hotkey("shift", "end")
    pyautogui.press("backspace")


def m_triplo():
    pyautogui.tripleClick(*a.COORDENADAS[CAMPO])
    pyautogui.press("backspace")


def m_duplo():
    a.duplo_clicar(CAMPO)
    pyautogui.press("backspace")


print("Tela de operações aberta, Operação preenchida e Tempo de Execução com algo (ex.: 0,07).")
print("Não mexa no mouse. Começando em 6s; olhe o campo Tempo de Execução a cada pausa.")
time.sleep(6)

for nome, metodo in (("backspace", m_backspace), ("home+shift+end", m_home_shift_end),
                     ("triplo clique", m_triplo), ("duplo clique", m_duplo)):
    print(f"\n>>> Método: {nome}")
    metodo()
    print("    campo DEPOIS DE LIMPAR (3s):")
    time.sleep(3)
    pyautogui.write(VALOR, interval=0.06)
    print(f"    campo DEPOIS DE DIGITAR {VALOR} (4s):")
    time.sleep(4)

print("\nFim. Sem TAB, nada foi validado.")


# 1.123.1.01.0047 - REV.FO 10 recusada pelo ERP
# 1.189.1.02.1024 - REV.FO 10 recusada pelo ERP
# 1.123.1.04.1009 - REV.FO 10 recusada pelo ERP
# 1.204.1.01.2144 - REV.FO 10 recusada pelo ERP
# 1.189.1.02.1013 - REV.FO 10 recusada pelo ERP
# 1.185.1.02.0051 - REV.FO 10 recusada pelo ERP
# 1.183.1.04.0009 - REV.FO 10 recusada pelo ERP
# 1.123.1.11.0642 - REV.FO 0 recusada pelo ERP
# 1.189.1.02.1010 - janela ERRO aberta (REV.FO 10)
