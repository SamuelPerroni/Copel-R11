from R11 import CONST
from R11RegrasGerais import RegrasGerais

class ConsultaMedica:
    def __init__(self, planilha, indices):
        self.planilha = planilha
        self.indices = indices        
        self.regrasGerais = RegrasGerais(planilha, indices)        
        print('init ConsultaMedica')

    def validar(self):
        maxLines = self.planilha.max_row + 1
        maxCols = self.planilha.max_column + 1
        for linha in range(2, maxLines):
            if self.planilha.cell(row=linha, column=self.indices['TICKET']).value is not None:
                if self.regrasGerais.reprovado(linha):
                    self.regrasGerais.reprovar(linha,'10')
                    continue
                if self.regrasGerais.semDados(linha):
                    self.regrasGerais.reprovar(linha, '10')
                    continue
                if self.regrasGerais.horaInicialForaDoPeriodo(linha):
                    self.regrasGerais.reprovar(linha, '10')
                    continue

                #celulas = [ self.planilha.cell(row=linha, column=coluna).value for coluna in range(1, maxCols)]  
                #print(celulas)
            
    def salvar(self):
        self.alterar(33,indicesAusenciaAbonada["RESULTADO"], "POSITIVADO")
        self.alterar(33,indicesAusenciaAbonada["SOLUCAO_RESPOSTA"], "888")        
        