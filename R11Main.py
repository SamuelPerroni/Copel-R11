import sys
import openpyxl
import datetime
from R11AusenciaAbonada import AusenciaAbonada
from R11ConsultaMedica import ConsultaMedica
from R11RegrasGerais import RegrasGerais
from R11RegrasReprovacao import RegrasReprovacao
   

def main():
    print('Iniciando...')
    regrasGerais = RegrasGerais()
    regrasReprovacao = RegrasReprovacao()
    ausenciaAbonada = AusenciaAbonada()
    consultaMedica = ConsultaMedica()
    k = input("Pressione ENTER para encerrar...")

if __name__ == '__main__':
    main()
