import os
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

load_dotenv()

class Planilha:
    def __init__(self, diretorios, api):
        self.diretorios = diretorios
        self.API = api

        caminho_honorarias = Path(os.getenv("PARCELAS_HONORARIAS"))
        self.wb = load_workbook(caminho_honorarias)
        self.sheet_honorarias = self.wb.active

        caminho_3C = Path(os.getenv("BASE_3C"))
        self.wb_3C = load_workbook(caminho_3C)
        self.sheet_3c = self.wb_3C.active

    def iterar_base(self):
        for linha in range (2, self.sheet_honorarias.max_row+1):
            if self.sheet_honorarias.cell(row=linha, column=1).value is None:
                break
            
            processo = self.sheet_honorarias.cell(row=linha, column=2).value
            reu = self.buscar_nome_do_reu(processo)      

            if reu == None:
                print(f"\n❌ Não foi possível localizar o processo: {processo} na planilha 3C.")
                continue
            
            else:
                try:
                    caminho, arquivo = self.diretorios.pegar_arquivo(reu)
                    novo_caminho, novo_arquivo = self.diretorios.renomear_arquivo(processo, caminho, arquivo)
                    self.API.upload(processo, novo_caminho, novo_arquivo)

                except:
                    print(f"\n❌ Não foi possível localizar arquivo do réu: {reu}")
                
    def buscar_nome_do_reu(self, processo):
        for linha in range (2, self.sheet_3c.max_row+1):
            temp_processo = self.sheet_3c.cell(row=linha, column=2).value

            if str(temp_processo) == str(processo):
                nome_do_reu = self.sheet_3c.cell(row=linha, column=1).value
                return nome_do_reu            
            
        return None

class Diretorios:
    def __init__(self):
        self.rede = Path("P:\\")        
    
    def pegar_arquivo(self, nome_do_reu):
        diretorio_do_reu = Path(os.path.join(self.rede, nome_do_reu))
        if diretorio_do_reu.is_dir():
            for elemento in diretorio_do_reu.rglob('*'):
                if "CONFISSÃO" in elemento.name.upper() or "CONFISSAO" in elemento.name.upper():
                    print(f"\n✅ Arquivo do réu: '{nome_do_reu}' localizado!!")
                    return elemento.parent, elemento.name
            
            return None
        return None
        
    def renomear_arquivo(self, processo, caminho, arquivo):
        caminho_original = os.path.join(caminho, arquivo)
        novo_nome = f"{processo}_{arquivo}".replace(' ', '_')
        novo_caminho = os.path.join(caminho, novo_nome)

        try:
            Path(caminho_original).rename(novo_caminho)
            print(f"✅ CONFISSÃO RENOMEADA PARA -> {novo_caminho}")
            return novo_caminho, novo_nome
        
        except Exception as e:
            print(f"\n❌ Não foi possível renomear a o arquivo CONFISSÃO.\nDetalhes do erro:{e}")

class API:
    def __init__(self):
        self.TOKEN = os.getenv("TOKEN")
        self.URL = os.getenv("URL")

        # Configurar cabeçalhos
        self.headers = {
            "Authorization": f"Token {self.TOKEN}",
            "Accept": "application/json"
        }
    
    def upload(self, processo, caminho, arquivo):        
        if not os.path.exists(caminho):
            print(f"Arquivo não encontrado: {caminho}")
            return

        print(f"🟠 Enviando {arquivo} para o processos {processo} ...")
        
        with open(Path(caminho), "rb") as f:
            files = {
                "document" : (arquivo, f, "application/pdf")
            }
            response = requests.post(self.URL, files=files, headers=self.headers)

        match response.status_code:
            case 200 | 202:
                print(f"✅ Upload realizado com sucesso.             Status: {response.status_code}\n")
            
            case 400 | 404:
                print(f"❌ Processos {processo} não foi encontrado no sistema.             Status: {response.status_code}")
            
            case _ :
                print(f"❌ Erro de upload desconhecido: \n\nprocesso: {processo}\ncaminho do arquivo: {os.path.join(caminho, arquivo)}\nStatus: {response.status_code}")

if __name__ == "__main__":
    diretorios = Diretorios()
    api = API()
    
    base = Planilha(diretorios, api)
    base.iterar_base()