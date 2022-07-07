from re_excel import*
import os
import sys

# Main, start of the program
if __name__ == "__main__":
    while True:
        ruta='./BCP SOLES.xlsx'
        wb = load_workbook(ruta)
        leer_libro(wb)
        #imprimir_valores(5,2,wb)
        #imprimir_fila(4,wb)
        #imprimir_columna(1,wb)
        #editar_celda('B11115',wb,ruta,'sddfdgfg')
        #B11118:G11121
        editar_varias_celdas('B11118','G11121',wb,ruta)
        leer_bloques_celdas('B11118','G11121',wb,ruta)
        sys.exit()