from win32com.client import GetObject, Dispatch
from subprocess import Popen
import time
from os.path import exists
import pandas as pd
import os
import sys


# Conexão com SAP GUI - Versão PHS

class SapSSO:
    def __init__(self, sapEnv='PHS [sapphsas01.pharma.aventis.com]', sapFile=''):
        self.sapEnv = sapEnv
        self.sapFile = sapFile
        self.session = self.getSapSession()

    def getSapGui(self):
        try:
            # Tenta conectar com SAP GUI
            sapGui = GetObject('SAPGUI')
            if sapGui is None:
                raise Exception('SAP GUI não encontrado')
            return sapGui.GetScriptingEngine
        except Exception as e:
            # Tenta uma abordagem alternativa
            try:
                sapGui = Dispatch('SAPGUI.ScriptingCtrl.1')
                return sapGui.GetScriptingEngine
            except Exception as e2:
                raise Exception(f'SAP GUI não encontrado. Erro: {str(e)}. Erro alternativo: {str(e2)}')

    def verificar_sap_gui_instalado(self):
        """Verifica se o SAP GUI está instalado no sistema"""
        caminhos_possiveis = [
            r'C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe',
            r'C:\Program Files\SAP\FrontEnd\SapGui\saplogon.exe',
            r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe',
            r'C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe'
        ]
        
        for caminho in caminhos_possiveis:
            if exists(caminho):
                return caminho
        return None

    def getSapSession(self):
        # Verifica se o SAP GUI está instalado
        sap_path = self.verificar_sap_gui_instalado()
        if not sap_path:
            raise Exception('SAP GUI não está instalado no sistema. Verifique se o SAP GUI está instalado corretamente.')
        
        self.sapFile = sap_path
        
        try:
            # Tenta abrir o SAP GUI
            Popen(self.sapFile)
            time.sleep(8)  # Aumenta o tempo de espera para garantir que o SAP GUI abra completamente
        except Exception as e:
            raise Exception(f'Erro ao abrir SAP GUI: {str(e)}')
        
        try:
            sapGui = self.getSapGui()
        except Exception as e:
            raise Exception(f'Erro ao conectar com SAP GUI: {str(e)}')
        
        session = None

        # Aguarda um pouco mais para as conexões ficarem disponíveis
        time.sleep(3)
        
        try:
            if sapGui.Connections.Count > 0:
                for conn in sapGui.Connections:
                    if conn.Description == self.sapEnv:
                        session = conn.Sessions(0)
                        break

            if session is None:
                connection = sapGui.OpenConnection(self.sapEnv, True)
                session = connection.Sessions(0)

            session.findById("wnd[0]").maximize()
            return session
        except Exception as e:
            raise Exception(f'Erro ao estabelecer sessão SAP: {str(e)}')

    def executeTransaction(self, transacao):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + transacao
        self.session.findById("wnd[0]").sendVKey(0)

    def converter_mhtml_para_excel(self, arquivo_mhtml, arquivo_excel=None):
        """
        Converte um arquivo MHTML para Excel
        """
        try:
            if not os.path.exists(arquivo_mhtml):
                print(f"Arquivo {arquivo_mhtml} não encontrado")
                return False
            
            if arquivo_excel is None:
                arquivo_excel = os.path.splitext(arquivo_mhtml)[0] + ".xlsx"
            
            with open(arquivo_mhtml, "r", encoding="utf-8") as file:
                conteudo = file.read()
            
            from io import StringIO
            tabelas = pd.read_html(StringIO(conteudo), header=0)
            
            if len(tabelas) > 0:
                df = tabelas[0]
                df.to_excel(arquivo_excel, index=False)
                print(f"Arquivo Excel gerado: {arquivo_excel}")
                return True
            else:
                print("Nenhuma tabela encontrada no arquivo MHTML")
                return False
                
        except Exception as e:
            print(f"Erro ao converter MHTML para Excel: {str(e)}")
            return False

    def extrair_coluna_material_mb52(self):
        """
        Extrai a coluna Material do arquivo MB52.xlsx e copia para área de transferência
        """
        try:
            # Converte MB52.MHTML para MB52.xlsx se necessário
            if not os.path.exists("MB52.xlsx"):
                if os.path.exists("MB52.MHTML"):
                    self.converter_mhtml_para_excel("MB52.MHTML", "MB52.xlsx")
                else:
                    print("Arquivo MB52.MHTML não encontrado")
                    return False
            
            # Lê o arquivo Excel
            df = pd.read_excel("MB52.xlsx")
            
            # Verifica se a coluna Material existe
            if "Material" not in df.columns:
                print("Coluna 'Material' não encontrada no arquivo MB52.xlsx")
                return False
            
            # Copia a coluna Material para a área de transferência
            df['Material'].to_clipboard(header=False, index=False)
            print(f"Coluna Material copiada para área de transferência ({len(df)} materiais)")
            return True
            
        except Exception as e:
            print(f"Erro ao extrair coluna Material: {str(e)}")
            return False

# Script SAP para execução do relatório SAP - Transação MB52

    def executar_transacao_MB52(self, session):

        # Maximiza a janela principal do SAP

        session.findById("wnd[0]").maximize()

        # Navegação e exportação conforme estrutura fornecida

        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = "CONTROLE_INV"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        # Move o arquivo para a pasta de destino

        session.findById("wnd[0]/tbar[1]/btn[43]").press() 
        session.findById("wnd[1]/usr/ctxtDY_PATH").text =r"C:\Users\i0336821\OneDrive - Sanofi\Área de Trabalho\Projetos\Inventário"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MB52.MHTML"
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

# Script SAP para execução do relatório SAP - Transação BMBC

    def executar_transacao_BMBC(self, session, colar_material=True):
        """
        Executa a transação BMBC com opção de colar dados da coluna Material
        
        Args:
            session: Sessão SAP
            colar_material (bool): Se True, cola os dados da área de transferência
        """
        # Maximiza a janela principal do SAP

        session.findById("wnd[0]").maximize()

        # Navegação e exportação conforme estrutura fornecida

        session.findById("wnd[0]/usr/subLIMITATIONS:RVBBINCO:0101/subSUB:RVBBINCO:0105/tabsTABSTRIP/tabpPUSH01/ssubLIMITATIONS:RVBBINCO:0106/subSELECTION:RVBBINCO:0110/subSUB:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press()
        
        # Se deve colar os dados da área de transferência

        if colar_material:
            print("Colando dados da coluna Material na tela de seleção...")
            # Aguarda um pouco para a tela carregar
            time.sleep(2)
            # Cola os dados da área de transferência (Ctrl+V)
            session.findById("wnd[1]").sendVKey(2)  # VKey 2 = Ctrl+V
            time.sleep(1)
        
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[0]").pressButton("MAST_GRID")
        session.findById("wnd[0]/shellcont").dockerPixelSize = 464
        session.findById("wnd[0]/shellcont").dockerPixelSize = 972
        session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BMBC.MHTML"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

    def executar_mb52_e_bmbc_com_material(self):
        """
        Executa MB52, extrai coluna Material e executa BMBC com os dados colados
        """
        try:
            print("=== Executando MB52 e BMBC com transferência de dados ===")
            
            # 1. Executa MB52
            print("1. Executando transação MB52...")
            self.executeTransaction("MB52")
            time.sleep(2)
            self.executar_transacao_MB52(self.session)
            time.sleep(3)
            
            # 2. Converte MB52.MHTML para Excel e extrai coluna Material
            print("2. Convertendo MB52.MHTML e extraindo coluna Material...")
            if not self.extrair_coluna_material_mb52():
                print("Erro ao extrair coluna Material do MB52")
                return False
            
            # 3. Executa BMBC
            print("3. Executando transação BMBC...")
            self.executeTransaction("BMBC")
            time.sleep(2)
            
            # 4. Cola os dados na tela de seleção do BMBC
            print("4. Colando dados na tela de seleção do BMBC...")
            self.executar_transacao_BMBC(self.session, colar_material=True)
            
            print("Processo concluído com sucesso!")
            return True
            
        except Exception as e:
            print(f"Erro durante execução: {str(e)}")
            return False


# Exemplo de uso
if __name__ == "__main__":
    sap = SapSSO()
    sap.executeTransaction("MB52")
