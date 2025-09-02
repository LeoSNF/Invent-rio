# pip install pywin32

from modulos.formPw import FormPw
from win32com.client import GetObject
from subprocess import Popen
from time import sleep
from os.path import exists
from enum import Enum

import pandas as pd

class Keys(Enum):
    enter = 0
    voltar = 3
    executar = 8
    expandirCabecalho = 26
    expandirSintese = 27
    expandirItem = 28
    comprimirCabecalho = 29
    comprimirSintese = 30
    comprimirItem = 31

class Sap:
    def __init__(self, sapEnv, userId=None, userPW=None, language='EN', connectBy = 1, sapFile = '', newSession=False):
        '''
        Use connectBy:
        1: Description
        2: Connection String
        3: Browser
        '''
        self.sapEnv = sapEnv
        self.userId = userId
        self.userPW = userPW
        self.language = language
        self.connectBy = connectBy
        self.sapFile = sapFile
        self.newSession = newSession
        self.__dictColumns = dict()
        self.session = self.__getSapSession__()
        self.sessions = {0: self.session}
        self.numCtrlSessions = 0

    def __getSapSession__(self):
        if not exists(self.sapFile): self.sapFile = r'C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe'
        if not exists(self.sapFile): self.sapFile.replace(' (x86)', '')
        if not exists(self.sapFile): raise Exception('SAP Gui (saplogon.exe) not founded')

        if self.connectBy == 3:
            Popen(self.sapFile + ' ' + self.sapEnv)
            sapGui = self.__getSapGui__()
            conn = sapGui.Connections(0)
            timeout = 10
            while conn.Sessions.Count == 0 and timeout:
                sleep(1)
                timeout -= 1
            if timeout == 0: raise Exception('Timeout - Não conseguimos conectar com a sessão do SAP após aguardar 10 segundos')
        else:
            Popen(self.sapFile)
            sapGui = self.__getSapGui__()
            conn = None
            session = None
            
            if self.connectBy == 1:
                if sapGui.Connections.Count > 0:
                    for conn in sapGui.Connections:
                        if conn.Description == self.sapEnv:
                            break
                if not conn: conn = sapGui.OpenConnection(self.sapEnv)

            else:
                if sapGui.Connections.Count > 0:
                    for conn in sapGui.Connections:
                        if self.sapEnv in conn.ConnectionString:
                            break
                if not conn: conn = sapGui.OpenConnectionByConnectionString(self.sapEnv)

        if self.newSession:
            numSessions = conn.Sessions.Count + 1
            conn.Sessions(0).createsession()
            while conn.Sessions.Count < numSessions: pass
            session = conn.Sessions(numSessions-1)
        else:
            session = conn.Sessions(0)

        if session.findById('wnd[0]/sbar').text.startswith('SNC logon'):
            session.findById('wnd[0]/usr/txtRSYST-LANGU').text = self.language
            session.findById('wnd[0]').sendVKey(0)
            session.findById('wnd[0]').sendVKey(0)
        elif session.Info.User == '':
            if not self.userId or not self.userPW:
                fPW = FormPw()
                self.userId = fPW.userId
                self.userPW = fPW.userPw

            session.findById('wnd[0]/usr/txtRSYST-BNAME').text = self.userId
            session.findById('wnd[0]/usr/pwdRSYST-BCODE').text = self.userPW
            session.findById('wnd[0]/usr/txtRSYST-LANGU').text = self.language
            session.findById('wnd[0]').sendVKey(0)
        
        session.findById('wnd[0]').maximize()
        return session

    def __getSapGui__(self, timeout = 30):
        if timeout:
            try:
                return GetObject('SAPGUI').GetScriptingEngine
            except:
                sleep(1)
                timeout -= 1
                return self.__getSapGui__(timeout)
        else:
            raise Exception('Não foi possível conectar com o SAP')

    @property
    def subForm(self):
        return self.session.findById('wnd[0]/usr').Children(1).Name

    @property
    def columnsTable(self):
        return self.__dictColumns

    @columnsTable.setter
    def columnsTable(self, field):
        tbl = self.session.findById(field)
        for i in range(0, tbl.Columns.Count):
            self.__dictColumns[tbl.Columns[i].Title] = i

    @property
    def columnsTableOK(self):
        return len(self.__dictColumns.values()) > 0

    def columnByTitle(self, columnTitle):
        return self.__dictColumns.get(columnTitle)

    def openSap(self):
        self.session = self.__getSapSession__()

    def createNewSession(self):
        conn = self.session.Parent
        numSessions = conn.Sessions.Count + 1
        conn.Sessions(0).createsession()
        while conn.Sessions.Count < numSessions: pass
        
        session2 = conn.Sessions(numSessions-1)
        self.numCtrlSessions += 1
        self.sessions[self.numCtrlSessions] = session2
        
        return session2

    def executeTransaction(self, transacao, session = None):
        '''
        Esta rotina irá colocar o '/n' na frente da transação informada e dará o enter\n\
        Caso queira utilizar outra sessão ao invés da sessão inicial, informar a session
        '''
        if not session: session = self.session
        session.findById('wnd[0]/tbar[0]/okcd').text = '/n' + transacao
        session.findById('wnd[0]').sendVKey(0)

    def exportTxtFile(self, pathName, fileName, button=11):
        '''
        Esta rotina irá acionar os comandos de finalização de exportação de TXT no SAP
        buttons: 0 - Gerar | 11 - Substuir | 7 - Ampliar
        '''
        self.session.findById('wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]').select() #não converter
        self.session.findById('wnd[1]/tbar[0]/btn[0]').press() #confirmar
        self.session.findById('wnd[1]/usr/ctxtDY_PATH').text = pathName
        self.session.findById('wnd[1]/usr/ctxtDY_FILENAME').text = fileName
        self.session.findById(f'wnd[1]/tbar[0]/btn[{button}]').press()        

    def fieldExists(self, field, session = None):
        if not session: session = self.session
        try:
            var = session.findById(field)
            return var
        except:
            return False

    def getStatusBar(self, session = None):
        if not session: session = self.session
        return session.findById('wnd[0]/sbar').Text

    def removeStatusBar(self, sendEnter=False, session = None):
        if not session: session = self.session

        if sendEnter: session.findById('wnd[0]').sendVKey(0)
        while session.findById('wnd[0]/sbar').Text != '' and session.findById('wnd[0]/sbar').MessageType != 'S':
            session.findById('wnd[0]').sendVKey(0)
            if session.findById('wnd[0]/sbar').MessageType == 'E': return session.findById('wnd[0]/sbar').Text
        
        if self.fieldExists('wnd[1]', session):
            iRow = 3
            statusError = ''
            try:
                while True:
                    if session.findById(f'wnd[1]/usr/lbl[3,{iRow}]').IconName == 'S_LEDR':
                        statusError += session.findById(f'wnd[1]/usr/lbl[7,{iRow}]').text
                    iRow += 1
            finally:
                session.findById('wnd[1]/tbar[0]/btn[0]').press()
                session.findById('wnd[0]/tbar[1]/btn[6]').press()
                session.findById('wnd[1]/usr/btnSPOP-OPTION2').press()
                return statusError

    def multipleSelection(self, fieldFullId):
        '''
        Copiar dados na área de transferência do Windows antes de chamar este método
        '''
        self.session.findById(fieldFullId).press() # Seleção Múltipla
        self.session.findById('wnd[1]/tbar[0]/btn[16]').press() # Apagar tudo
        self.session.findById('wnd[1]/tbar[0]/btn[24]').press() # Colar
        self.session.findById('wnd[1]/tbar[0]/btn[8]').press() # Transferir           

    def clearMultipleSelection(self, fieldFullId):
        self.session.findById(fieldFullId).press() # Seleção Múltipla
        self.session.findById('wnd[1]/tbar[0]/btn[16]').press() # Apagar tudo
        self.session.findById('wnd[1]/tbar[0]/btn[8]').press() # Transferir  

    def sendKeys(self, key: Keys):
        self.session.findById('wnd[0]').sendVKey(key.value)

    def loadVariant(self, variant):
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/txtV-LOW").text = variant
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()  

    def save(self, session = None):
        if not session: session = self.session
        session.findById('wnd[0]/tbar[0]/btn[11]').press()
        if self.fieldExists("wnd[1]/usr/btnBUTTON_1"): self.fieldExists("wnd[1]/usr/btnBUTTON_1").press() # Adic. Rafa 16/07/2024
        if self.fieldExists('wnd[1]/usr/txtSPOP-TEXTLINE1'):
            if session.findById('wnd[1]/usr/txtSPOP-TEXTLINE1').text == "Document Incomplete":
                session.findById('wnd[1]/usr/btnSPOP-VAROPTION1').press()
        if self.fieldExists('wnd[1]', session): session.findById('wnd[1]').close()

        status = self.removeStatusBar(session=session)
        return status


    def se16n(self, table: str, fields: list, path: str, file: str, filterFields: list = [], filterValues: dict = {}):
        '''
        table - name of SAP Table\n
        fields - list with fields to extract from SAP in key\n
        filterFields - list of fields necessary to filterApply\n
        filterValues - list of dataframe or series with just columns values to filter (in the same index of filterFields)
        '''

        warningsDict = {'PT':{'tableAuth':'sem autorização', 'tableExistence':'verificar o nome', 'valuesNotFound':'Nenhum valor encontrado'},
                        'EN':{'tableAuth':'not authorized', 'tableExistence':'check the name', 'valuesNotFound':'No values found'}}

        session = self.session
        self.executeTransaction("SE16N")
        session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table
        session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
        session.findById("wnd[0]").sendVKey(0)
        
        # if "not authorized" in session.findById("wnd[0]/sbar").Text: raise Exception(f'No access to table {table}')
        # if self.fieldExists("wnd[1]/usr/txtMESSTXT1"):
        #     if "not authorized" in session.findById("wnd[1]/usr/txtMESSTXT1").Text: raise Exception(f'No access to table {table}')
        
        sbarOutput = session.findById("wnd[0]/sbar").Text
        tableAuthWarning = warningsDict.get(self.language).get('tableAuth')
        tableExistenceWarning = warningsDict.get(self.language).get('tableExistence')

        if tableAuthWarning in sbarOutput or tableExistenceWarning in sbarOutput:
            raise Exception(f'No access to table {table}')
        if self.fieldExists("wnd[1]/usr/txtMESSTXT1"):
            if "not authorized" in session.findById("wnd[1]/usr/txtMESSTXT1").Text: raise Exception(f'No access to table {table}')
        
        self.selectFields(fields)
        if len(filterFields): self.filterApply(filterFields, filterValues)

        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        # if session.findById("wnd[0]/sbar").Text == 'No values found':
        #     session.findById('wnd[0]').sendVKey(0)
        #     raise ValueError('No values found')

        sbarOutput = session.findById("wnd[0]/sbar").Text
        valuesNotFoundWarning = warningsDict.get(self.language).get('valuesNotFound')

        if valuesNotFoundWarning in sbarOutput:
            session.findById('wnd[0]').sendVKey(0)
            raise ValueError('No values found')
        
        numColsReturnedSapTable = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").ColumnCount

        if numColsReturnedSapTable != len(fields):
            raise ValueError(f'The number of columns in the table returned by the SAP query does not match the expected number of columns. Expected {len(fields)} but returned {numColsReturnedSapTable}')

        # # Aumenta largura das colunas para tentar trazer as colunas no arquivo de texto correto
        # for field in fields:
        #     session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setColumnWidth(field, 30)

        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        self.exportTxtFile(path, file, button=11)

    def selectFields(self, fields: list):
        session = self.session
        nRows = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").VisibleRowCount
        iScroll = 0
        isToExit = False
        
        while True:
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.Position = nRows * iScroll
            
            rows = session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").rows
            for iRow in range(nRows):
                if rows(iRow).Count == 7:
                    # if not fields.get(rows(iRow)(6).Text): rows(iRow)(5).Selected = 0
                    if not rows(iRow)(6).Text in fields: rows(iRow)(5).Selected = 0

                if rows(iRow).Count == 0:
                    isToExit = True
                    break
            
            iScroll = iScroll + 1
            if isToExit: break    
    
    def filterApply(self, filterFields: list, filterValues: dict):
        session = self.session
        
        for index, field in enumerate(filterFields):
            session.findById("wnd[0]").sendVKey(71)
            session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = field
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").setfocus()
            session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press()
            session.findById("wnd[1]/tbar[0]/btn[34]").press()

            # filterValues[index].to_clipboard(index=False, header=False)

            # filterValues.to_clipboard(index=False, header=False)
            filterValues[field].to_clipboard(index=False, header=False)

            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()


def getConnectionString():
    try:
        sapGui = GetObject('SAPGUI').GetScriptingEngine    
        conn = sapGui.Connections(0)
    except:
        return 'Acessar o ambiente do SAP que deseja pegar a string de conexão'
    else:
        return conn.connectionString.strip().replace('"','""')