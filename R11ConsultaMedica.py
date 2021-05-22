from R11 import CONST
from R11RegrasGerais import RegrasGerais

class ConsultaMedica:
    def __init__(self, planilha, indices):
        self.planilha = planilha
        self.indices = indices                
        print('init ConsultaMedica')

    def alterar(self, linha, coluna, valor):
        self.planilha.cell(row=linha, column=coluna).value = valor

    def regra_sem_dados(linha):
        return False
        
    def validar(self):
        self.alterar(33,indicesAusenciaAbonada["RESULTADO"], "POSITIVADO")
        self.alterar(33,indicesAusenciaAbonada["SOLUCAO_RESPOSTA"], "888")        
        