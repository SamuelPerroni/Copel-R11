from R11 import CONST
from R11RegrasGerais import RegrasGerais

class ConsultaMedica(RegrasGerais):
    def __init__(self, planilha, indices):
        self.planilha = planilha
        self.indices = indices        
        #self.regrasGerais = RegrasGerais(planilha, indices)        
        print('init ConsultaMedica')



    def validar(self):
        maxLines = self.planilha.max_row + 1
        maxCols = self.planilha.max_column + 1
        for linha in range(2, maxLines):
            if self.getVal(linha, "TICKET") is not None:
                if self.reprovado(linha):
                    #já tem indicação na resposta de REPROVADO
                    continue
                if self.semDados(linha):
                    #colunas de AUSENCIA, PRESENCA e QTD_MARCACOES com SEM DADOS
                    self.reprovar(linha, '03')
                    continue

                if self.is_integer(self.getVal(linha, "QTD_MARCACOES")):
                    #quantidade de marcacoes é um numero
                    if not int(self.getVal(linha, "QTD_MARCACOES")) % 2 == 0:
                        #quantidade de marcacoes é impar
                        self.reprovar(linha,"03")
                        continue
                    else:
                        if self.horaInicialForaDoPeriodo(linha):
                            #valor da coluna HORA_INICIAL fora dos intervalos dos periodos TEMPO_TEORICO e TEMPO_TEORICO
                            if self.jornadaCumpridaDentroDoTempoTeorico(linha):
                                self.reprovar(linha, '09')
                            else:
                                self.reprovar(linha, '07')
                            continue
                        if self.semRegistrosNaEscalaDeIntervalo(linha):
                            self.reprovar(linha,'07')
                            continue
                        #Validar este metodo.....?????????
                        if self.validarHoraInicialIntervalo(linha):
                            retorno = self.calcularAbono(linha)
                            if not retorno is None:
                                if retorno[0] == CONST.REPROVADO:
                                    self.reprovar(linha, retorno[1])
                                else:
                                    self.abonar(linha, retorno[1])
                            continue
                        
                        

