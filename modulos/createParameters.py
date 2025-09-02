'''Leve modificação na classe CreateParameters feita por mim para aceitar caracters especiais'''
from crypto import generateKey
from os import getcwd, remove
from subprocess import run

class CreateParameters:
    def __init__(self, dictParameters, key) -> None:
        self.dictParameters = dictParameters
        self.key = key
        self.mountModuleParameters()
        self.writeInFile()
        self.mountModuleWriteParameters()
        self.writeInFile()

    def mountModuleParameters(self):
        self.file = getcwd() + '\\modulos\\parameters.py'

        self.text = 'import json' + chr(10)
        self.text += 'from modulos.crypto import Crypto' + chr(10)
        self.text += 'from os import getcwd' + chr(10) + chr(10)
        self.text += 'class Parameters:' + chr(10)
        self.text += '    def __init__(self, key):' + chr(10)
        self.text += '        self.c = Crypto(key)' + chr(10)

        for key in self.dictParameters:
            self.text += f'        self.__{key} = None' + chr(10)

        for key, item in self.dictParameters.items():
            self.text += '# -------------------------------------------------------' + chr(10)
            self.text += '    @property' + chr(10)
            self.text += f'    def {key}(self):' + chr(10)
            if item[0] == 'c':
                self.text += f"        return self.c.decrypt(self.__{key}).decode('ascii')" + chr(10) + chr(10)
                self.text += f'    @{key}.setter' + chr(10)
                self.text += f'    def {key}(self, value):' + chr(10)
                self.text += f'        self.__{key} = self.c.crypt(value)' + chr(10)
            else:
                self.text += f'        return self.__{key}' + chr(10) + chr(10)
                self.text += f'    @{key}.setter' + chr(10)
                self.text += f'    def {key}(self, value):' + chr(10)
                self.text += f'        self.__{key} = value' + chr(10)

        self.text += '# -------------------------------------------------------' + chr(10)
        self.text += '    def writeParametersFile(self):' + chr(10)
        self.text += '        par = {'

        for key, item in self.dictParameters.items():
            if item[0] == 'c':
                self.text += chr(10) + f"            '{key}': self.__{key}.decode('ascii'),"
            else:
                self.text += chr(10) + f"            '{key}': self.__{key},"

        self.text += self.text[len(self.text):len(self.text)-1] + chr(10) + '        }' + chr(10) + chr(10)
        self.text += '        par = json.dumps(par, ensure_ascii=False)' + chr(10) + chr(10)
        self.text += "        f = open(getcwd() + '\\parameters.json', 'w', encoding='utf-8')" + chr(10)
        self.text += '        f.write(par)' + chr(10)
        self.text += '        f.close' + chr(10)
        self.text += '# -------------------------------------------------------' + chr(10)
        self.text += "    def readParameters(self, file=getcwd() + '\\parameters.json'):" + chr(10)
        self.text += "        f = open(file, 'r', encoding='utf-8')" + chr(10)
        self.text += '        par = json.load(f)' + chr(10)
        self.text += '        f.close()' + chr(10)

        for key, item in self.dictParameters.items():
            if item[0] == 'c':
                self.text += chr(10) + f"        self.__{key} = par['{key}'].encode('ascii')"
            else:
                self.text += chr(10) + f"        self.__{key} = par['{key}']"

    def mountModuleWriteParameters(self):
        self.file = getcwd() + '\\writeParameters.py'

        self.text = 'from modulos.parameters import Parameters' + chr(10) + chr(10)
        self.text += f"p = Parameters({self.key})" + chr(10)

        for key, item in self.dictParameters.items():
            if '\\' in str(item[1]):
                self.text += f"p.{key} = r'{item[1]}'" + chr(10)
            elif type(item[1]) in (int, float, list, tuple, dict):
                self.text += f"p.{key} = {item[1]}" + chr(10)
            else:
                self.text += f"p.{key} = '{item[1]}'" + chr(10)

        self.text += 'p.writeParametersFile()' + chr(10)

    def writeInFile(self):
        with open(self.file, 'w', encoding='utf-8') as f:
            f.write(self.text)
            f.close()



# Create a dictionary with parameters and generate a crypto key
# Obs Rafael 17/04/2025:
# ATENÇÃO! Sempre deixar userId e userPw vazio antes de subir no GIT
# Se precisar atualizar login e senha do SAP, utilize o arquivo updateLogin.py
# sapTablesDictParams são parametros utilizado na extração dos dados no SAP
# txtParams são parametros utilizados para converter os arquivos TXT em dataframe. Deve-se observar que os campos colNames vão renomear as tabelas por posição
# não consegui fazer no formato de dicionario pois há tabelas com o nome repetido (por exemplo VBAP). Mesmo se tentar aumentar a largura da coluna no SAP os nomes vem cortados e 
# alguns repetidos.
# pandasParams são os parametros utilizados para manipular alguns dataframes no pandas.
# Checar o arquivo Tabela dados históricos_analise_mr.xlsx para mais informações

dictPar = dict()
dictPar['version'] = ('nc', '0.0.1')
dictPar['sapEnv'] =  ('nc', 'PHS [sapphsas01.pharma.aventis.com]')
dictPar['sapLang'] =  ('nc', 'EN')
dictPar['sapConnBy'] =  ('nc', 2)
dictPar['sapFile'] =  ('nc', 'C:\\Program Files\\SAP\\FrontEnd\\SapGui\\saplogon.exe')
dictPar['sapNewSession'] = ('nc', 0)
dictPar['sapExportFileDestPath'] = ('nc', '')


# key = generateKey() # Gera a chave criptográfica
key = b'pVOxw0erfHGl4agvXg-nQlu6PySZYn6m7-_kZrlJ3yQ='
print(key) # Anote a chave, pq será necessária para acessar os parâmetros
c = CreateParameters(dictPar, key)

run('python.exe writeParameters.py')
remove('writeParameters.py')