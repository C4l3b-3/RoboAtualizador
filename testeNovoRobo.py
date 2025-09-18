import pandas as pd
import os
from openpyxl import load_workbook
import pyautogui
import time

#Criar o dicionário {Placa:KM}
planilha_principal = "planilha_teste.xlsx" #Caminho da planilha 
df_principal = pd.read_excel(planilha_principal)
km_dict = dict(zip(
    df_principal["PLACA"], 
    df_principal["KM FINAL CALCULADO"].round().astype(int)
))
#Inicia o Robô 
def clicar_imagem(imagem, descricao="", tentativas=5, espera=2, conf=0.9):
    for i in range(tentativas):
        local = pyautogui.locateCenterOnScreen(imagem, confidence=conf)
        if local:
            pyautogui.click(local)
            print(f"Cliquei em {descricao}")
            return True
        else:
            print(f"[Aguardando] Tentativa {i+1}: não encontrou {descricao}")
            time.sleep(espera)
    print(f"[ERRO] Não encontrou {descricao} após {tentativas} tentativas")
    return False

#Abrir o sistema 
arquivo_principal = "C:\\InfoSistemas\\CR_Ribal11\\Locadoras.exe"
os.startfile(arquivo_principal)
time.sleep(10)

# Faz login
clicar_imagem("Botao_login.png", "Botão login")
clicar_imagem("Selecao_login.png", "Seleção do Login")
clicar_imagem("Campo_Senha.png", "Campo Senha")
senha_do_usuario = "2311"
pyautogui.typewrite(senha_do_usuario)
clicar_imagem("Botao_OK_login.png", "Botão do Login")
time.sleep(40)

primeira_vez = True
for placa, km in km_dict.items():
    print(f"Atualizando placa {placa} -> KM {km}")
    
    #Abrir Aba para atualizar as placas 
    if primeira_vez:        
        if not clicar_imagem("aba_cadastro.png","Aba Cadastro", tentativas=20,espera=2): 
            print("[ERRO] Sistema não carregado")
            exit()
        else:
            print("[OK] Sistema carregado")
        time.sleep(1)
        if not clicar_imagem("botao_veiculos.png","Botão Veículos"): continue
        time.sleep(2)
        if not clicar_imagem("lupa_pesquisaPlaca.png","Lupa de pesquisa Placa"): continue
        time.sleep(1)
        if not clicar_imagem("campo_digitarPlaca.png","Campo digitar placa"): continue
        primeira_vez = False
    else:
        if not clicar_imagem("lupa_pesquisaPlaca.png","Lupa de pesquisa Placa"): continue
        time.sleep(1)
        if not clicar_imagem("campo_digitarPlaca.png","Campo digitar placa"): continue

    pyautogui.typewrite(placa)
    for _ in range(4):
        pyautogui.press("enter")
    time.sleep(2)
    if not clicar_imagem("botao_editar.png","Botão Editar"): continue
    time.sleep(1)
    if km is None:
        print(f"[aviso] KM vazio para a placa {placa}, pulando")
        continue
    pyautogui.click(752, 351)
    time.sleep(0.3)
    pyautogui.press("backspace", presses=10, interval=0.05)
    pyautogui.typewrite(str(int(km)))
    print(f"[OK]KM atualizado para {km}")
    if not clicar_imagem("botao_salvar.png","Botão de Salvar edição"): continue
    time.sleep(1)
         
    


#Preciso aprimorar - Ele fechar pop ups sozinho clicando no 'X' e controlar melhor a quilometragem 

