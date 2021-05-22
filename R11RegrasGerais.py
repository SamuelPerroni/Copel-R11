from R11 import CONST

class RegrasGerais:
    def __init__(self, planilha, indices):
        self.indices = indices
        self.planilha = planilha

    def __getVal(self, linha, indice):
        return self.planilha.cell(row=linha, column=self.indices[indice]).value

    def reprovar(self, linha, valor):
        self.planilha.cell(row=linha, column=self.indices["RESULTADO"]).value = "REPROVADO"
        self.planilha.cell(row=linha, column=self.indices["SOLUCAO_RESPOSTA"]).value = valor
    
    def abonar(self, linha, valor):
        self.planilha.cell(row=linha, column=self.indices["RESULTADO"]).value = CONST.ABONADO
        self.planilha.cell(row=linha, column=self.indices["SOLUCAO_RESPOSTA"]).value = valor

    def reprovado(self, linha):
        return self.__getVal(linha,"RESULTADO") == CONST.REPROVADO

    def semDados(self, linha):
        return (self.__getVal(linha,"AUSENCIA") == CONST.SEM_DADOS 
                and self.__getVal(linha,"PRESENCA") == CONST.SEM_DADOS 
                and self.__getVal(linha,"QTD_MARCACOES") == CONST.SEM_DADOS)
    
    def __entrePeriodo(self, hora, periodos):
        """Retorna True or False se uma determinada hora está entre dois horários.
           Pametros: hora (string, ex: 23:30) - periodos (string de horas, ex: "09:00 - 13:00;14:00 - 16:00")"""
        range_periodos = [tuple(x.split(" - ")) for x in periodos.split(";")]
        for periodo in range_periodos:
            if  periodo[0] <= hora <= periodo[1]:
                return True
        return False

    def horaInicialForaDoPeriodo(self, linha):
        horaInicial = self.__getVal(linha,"HORA_INICIAL")
        horasTeorico = self.__getVal(linha,"HORAS_TEORICO")   
        tempoTeorico = self.__getVal(linha,"TEMPO_TEORICO")        
        periodo1 = self.__getVal(linha,"PERIODO1")
        periodo2 = self.__getVal(linha,"PERIODO2")
        if horasTeorico == "4,00":
            return not self.__entrePeriodo(horaInicial, tempoTeorico)
        else:
            return not (self.__entrePeriodo(horaInicial, periodo1) or self.__entrePeriodo(horaInicial, periodo2))


