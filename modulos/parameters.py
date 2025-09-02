import json
from modulos.crypto import Crypto
from os import getcwd

class Parameters:
    def __init__(self, key):
        self.c = Crypto(key)
        self.__version = None
        self.__sapEnv = None
        self.__sapLang = None
        self.__sapConnBy = None
        self.__sapFile = None
        self.__sapNewSession = None
        self.__sapExportFileDestPath = None
# -------------------------------------------------------
    @property
    def version(self):
        return self.__version

    @version.setter
    def version(self, value):
        self.__version = value
# -------------------------------------------------------
    @property
    def sapEnv(self):
        return self.__sapEnv

    @sapEnv.setter
    def sapEnv(self, value):
        self.__sapEnv = value
# -------------------------------------------------------
    @property
    def sapLang(self):
        return self.__sapLang

    @sapLang.setter
    def sapLang(self, value):
        self.__sapLang = value
# -------------------------------------------------------
    @property
    def sapConnBy(self):
        return self.__sapConnBy

    @sapConnBy.setter
    def sapConnBy(self, value):
        self.__sapConnBy = value
# -------------------------------------------------------
    @property
    def sapFile(self):
        return self.__sapFile

    @sapFile.setter
    def sapFile(self, value):
        self.__sapFile = value
# -------------------------------------------------------
    @property
    def sapNewSession(self):
        return self.__sapNewSession

    @sapNewSession.setter
    def sapNewSession(self, value):
        self.__sapNewSession = value
# -------------------------------------------------------
    @property
    def sapExportFileDestPath(self):
        return self.__sapExportFileDestPath

    @sapExportFileDestPath.setter
    def sapExportFileDestPath(self, value):
        self.__sapExportFileDestPath = value
# -------------------------------------------------------
    def writeParametersFile(self):
        par = {
            'version': self.__version,
            'sapEnv': self.__sapEnv,
            'sapLang': self.__sapLang,
            'sapConnBy': self.__sapConnBy,
            'sapFile': self.__sapFile,
            'sapNewSession': self.__sapNewSession,
            'sapExportFileDestPath': self.__sapExportFileDestPath,
        }

        par = json.dumps(par, ensure_ascii=False)

        f = open(getcwd() + '\parameters.json', 'w', encoding='utf-8')
        f.write(par)
        f.close
# -------------------------------------------------------
    def readParameters(self, file=getcwd() + '\parameters.json'):
        f = open(file, 'r', encoding='utf-8')
        par = json.load(f)
        f.close()

        self.__version = par['version']
        self.__sapEnv = par['sapEnv']
        self.__sapLang = par['sapLang']
        self.__sapConnBy = par['sapConnBy']
        self.__sapFile = par['sapFile']
        self.__sapNewSession = par['sapNewSession']
        self.__sapExportFileDestPath = par['sapExportFileDestPath']