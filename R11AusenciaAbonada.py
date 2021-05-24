from R11 import CONST
from R11RegrasGerais import RegrasGerais

class AusenciaAbonada(RegrasGerais):
    def __init__(self, planilha, indices):
        self.planilha = planilha
        self.indices = indices        
        print('init AusenciaAbonada')

    def alterar(self, linha, coluna, valor):
        self.planilha.cell(row=linha, column=coluna).value = valor

    #def validar(self):
        #self.alterar(32,indicesAusenciaAbonada["RESULTADO"], "POSITIVADO")
        #self.alterar(32,indicesAusenciaAbonada["SOLUCAO_RESPOSTA"], "888")

        