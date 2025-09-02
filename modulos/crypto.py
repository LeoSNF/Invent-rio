#pip install cryptography

from cryptography.fernet import Fernet

def generateKey():
    return Fernet.generate_key()

class Crypto:
    def __init__(self, key):
        self.f = Fernet(key)
    
    def crypt(self, value):
        if type(value) != bytes:
            try:
                value = value.encode('ascii')
            except:
                raise Exception(f'Não foi possível converter o {value} para ser criptografado')
        return self.f.encrypt(value)
    
    def decrypt(self, value):
        if type(value) != bytes:
            try:
                value = value.enconde('ascii')
            except:
                raise Exception(f'Não foi possível converter o {value} para ser descriptografado')
        if value: return self.f.decrypt(value)
