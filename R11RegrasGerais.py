from io import SEEK_CUR
from R11 import CONST
from datetime import datetime

class RegrasGerais:
    def __init__(self, planilha, indices):
        self.indices = indices
        self.planilha = planilha

    def is_integer(self, n):
        try:
            float(n)
        except ValueError:
            return False
        else:
            return float(n).is_integer()

    def stringToTime(self, hora):
        return datetime.strptime('19:00','%H:%M')

    def diffHours(self, horaMaior, horaMenor):
        """EX.: horaMaior= "21:00" - horaMenor= "16:30" """
        return (self.stringToTime(horaMaior) - self.stringToTime(horaMenor)).seconds/60/60

    def totalHours(self, horas):
        """horas (string de horas. Ex.: "09:00 - 13:00;14:00 - 16:00")"""
        total = 0.0
        range_periodos = [tuple(x.split(" - ")) for x in horas.split(";")]
        for periodo in range_periodos:
            total += self.diffHours(periodo[1], periodo[0])
        return total

    def getVal(self, linha, indice):
        return self.planilha.cell(row=linha, column=self.indices[indice]).value

    def reprovar(self, linha, valor):
        self.planilha.cell(row=linha, column=self.indices["RESULTADO"]).value = "REPROVADO"
        self.planilha.cell(row=linha, column=self.indices["SOLUCAO_RESPOSTA"]).value = valor
    
    def abonar(self, linha, valor):
        self.planilha.cell(row=linha, column=self.indices["RESULTADO"]).value = CONST.ABONADO
        self.planilha.cell(row=linha, column=self.indices["SOLUCAO_RESPOSTA"]).value = valor

    def reprovado(self, linha):
        return self.getVal(linha,"RESULTADO") == CONST.REPROVADO

    def semDados(self, linha):
        return (self.getVal(linha,"AUSENCIA") == CONST.SEM_DADOS 
                and self.getVal(linha,"PRESENCA") == CONST.SEM_DADOS 
                and self.getVal(linha,"QTD_MARCACOES") == CONST.SEM_DADOS)
    
    def __marcacoesPorPeriodo(self, marcacoes, periodo):
        per = periodo.split(" - ")
        range_marcacoes = [tuple(x.split(" - ")) for x in marcacoes.split(";")]
        retorno_marcacoes = ""
        sep = ""
        for marcacao in range_marcacoes:
            if  per[0] <= marcacao[0] <= per[1] or per[0] <= marcacao[1] <= per[1]:        
                retorno_marcacoes += sep + marcacao[0] + " - " + marcacao[1]
                sep = ";"
        return retorno_marcacoes

    def __horasEntrePeriodos(self, hora, periodos):
        """Retorna True or False se uma determinada hora está entre dois horários.
           Pametros: hora (string, ex: 23:30) - periodos (string de horas, ex: "09:00 - 13:00;14:00 - 16:00")"""
        range_periodos = [tuple(x.split(" - ")) for x in periodos.split(";")]
        for periodo in range_periodos:
            if  periodo[0] <= hora <= periodo[1]:
                return True
        return False
    
    def __marcacoesEntrePeriodos(self, marcacoes, periodos):
        """Retorna True or False se uma determinada hora está entre dois horários.
           Pametros: marcacoes (string, ex: "09:00 - 13:00;14:00 - 16:00") - periodos (string de horas, ex: "09:00 - 13:00;14:00 - 16:00")"""
        range_periodos = [tuple(x.split(" - ")) for x in periodos.split(";")]
        range_marcacoes = [tuple(x.split(" - ")) for x in marcacoes.split(";")]
        for periodo in range_periodos:
            for marcacao in range_marcacoes:
                if  periodo[0] <= marcacao[0] < marcacao[1] <= periodo[1]:
                    return True
        return False

    def horaInicialForaDoPeriodo(self, linha):
        horaInicial = self.getVal(linha,"HORA_INICIAL")
        tempoTeorico = self.getVal(linha,"TEMPO_TEORICO")        
        return not self.__horasEntrePeriodos(horaInicial, tempoTeorico)

    def marcacaoImpar(self, linha):
        return (self.getVal(linha, "QTD_MARCACOES") == CONST.SEM_DADOS 
            and not int(self.getVal(linha, "QTD_MARCACOES")) % 2 == 0)

    def jornadaCumpridaDentroDoTempoTeorico(self, linha):
        # Aqui considerar as marcações de AUSENCIA E PRESENCA
        
        horasTeorico = self.getVal(linha,"HORAS_TEORICO").replace(',','.')  
        tempoTeorico =  self.getVal(linha,"TEMPO_TEORICO")
        marcacoes = self.getVal(linha,"MARCACOES")
        totalHours = self.totalHours(marcacoes)
        dentrodoTempoTeorico = self.__marcacoesEntrePeriodos(self, marcacoes, tempoTeorico)
        return totalHours >= float(horasTeorico) and dentrodoTempoTeorico
    
    def semRegistrosNaEscalaDeIntervalo(self, linha):
        """
        Utilizar o valores entre os periodos: 08:30 - 11:30;14:00 - 17:00 => 11:30 - 14:00        
        """

        tempoTeorico =  self.getVal(linha,"TEMPO_TEORICO")
        marcacoes = self.getVal(linha,"MARCACOES")
        return self.__marcacoesEntrePeriodos(marcacoes, tempoTeorico)

    def validarHoraInicialIntervalo(self, linha):
        horaInicial = self.getVal(linha,"HORA_INICIAL")
        horasTeorico = self.getVal(linha,"HORAS_TEORICO")   
        tempoTeorico = self.getVal(linha,"TEMPO_TEORICO")        
        periodo1 = self.getVal(linha,"PERIODO1")
        periodo2 = self.getVal(linha,"PERIODO2")
        if horasTeorico == "4,00":
            return self.__horasEntrePeriodos(horaInicial, tempoTeorico)
        else:
            return self.__horasEntrePeriodos(horaInicial, periodo1) or self.__horasEntrePeriodos(horaInicial, periodo2)

    def calcularAbono(self, linha):
        horaInicial = self.getVal(linha,"HORA_INICIAL")
        horasTeorico = self.getVal(linha,"HORAS_TEORICO")   
        tempoTeorico = self.getVal(linha,"TEMPO_TEORICO")        
        periodo1 = self.getVal(linha,"PERIODO1")
        periodo2 = self.getVal(linha,"PERIODO2")
        if horasTeorico == "4,00":
            return ""
        else:
            pass

    def validaSolicitacaoJornada(self, tempoTeorico, periodo1, periodo2, solicitacaoInicial, solicitacaoFinal):
        '''Valida se o periodo solicitado é igual a jornada de trabalho.
            Utilizado para Validar se o abono de horas da Ausencia Abonada será do dia todo.'''
            
        tempoTeoricoInicio = tempoTeorico[0:tempoTeorico.find("-")-1]
        periodo1Inicio = periodo1[0:periodo1.find("-")-1]
        tempoTeoricoFinal = tempoTeorico[tempoTeorico.find("-")+2:]
        periodo2Final = periodo2[periodo2.find("-")+2:]

        if (solicitacaoInicial == tempoTeoricoInicio or solicitacaoInicial == periodo1Inicio) and (solicitacaoFinal == tempoTeoricoFinal or solicitacaoFinal == periodo2Final):
            return True
        else:
            return False

    def validaPTBALL(self, registro, ptball, solicitacaoInicial, solicitacaoFinal):
        '''
            Se atingiu o limite do PT_BAL00.
        '''
        tempoTotalSolicitado = solicitacaoFinal - solicitacaoInicial

        if ptball + tempoTotalSolicitado < 40:
            solicitacao = "Reprovado"
            solucao = "13"
        else:
            novoPtball = ptball + tempoTotalSolicitado
            #Chamar função para calcular o abono que irá dar e adicionar o novo PT_ball nos ticket
            #com mesmo registro nas linhas abaixo.
            