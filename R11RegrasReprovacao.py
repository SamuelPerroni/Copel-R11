class RegrasReprovacao:
    def __init__(self, planilha, indices):
        self.planilha = planilha
        self.indices = indices        
        print('init RegrasReprovacao')

    def alterar(self, linha, coluna, valor):
        self.planilha.cell(row=linha, column=coluna).value = valor

    def validar(self):
        self.alterar(31,indicesAusenciaAbonada["RESULTADO"], "POSITIVADO")
        self.alterar(31,indicesAusenciaAbonada["SOLUCAO_RESPOSTA"], "888")        
        