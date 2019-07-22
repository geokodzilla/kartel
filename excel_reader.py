# -*- coding: utf-8 -*-

from openpyxl import load_workbook, Workbook
from collections import OrderedDict
from app_blad import *
import os

class ExcelReader(object):

    KOL = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H'}
    
    def __init__(self, filename, numer_arkusza='1', wiersz_naglowka='1', wiersz_danych='2', kolumna_danych='1', id=['3', '4']):
        self.filename = filename
        self.numer_arkusza = numer_arkusza
        self.wiersz_naglowka = wiersz_naglowka
        self.wiersz_danych = wiersz_danych
        self.kolumna_danych = kolumna_danych
        self.id = id
        self._data_validation()
        self._file_validation()
        
    def _data_validation(self):
        try:
            self.numer_arkusza = int(self.numer_arkusza)
            self.wiersz_naglowka = int(self.wiersz_naglowka)
            self.wiersz_danych = int(self.wiersz_danych)
            self.kolumna_danych = int(self.kolumna_danych)
            self.id = [int(i) if i != '' else i for i in self.id]
        except ValueError:
            raise AppBlad('Błędne parametry wczytania pliku')
            
    def _file_validation(self):
        if not os.path.exists(self.filename):
            raise AppBlad('Wskazany plik nie istniej %s' % self.filename)
        
    def read(self):
        wb = load_workbook(self.filename, data_only=True)
        sn = wb.sheetnames # nazwy arkuszy
        arkusz = wb[sn[self.numer_arkusza-1]] # wybieramy z listy arkusz wg kolejności
        last_col = len(arkusz[str(self.wiersz_naglowka)]) + 1 # ustalamy pozycję ostatniej kolumny
        data = OrderedDict()
        try:
            col = self.KOL[self.kolumna_danych]
        except KeyError:
            raise AppBlad('Kolumny poza zakresem')
        else:
            for i in range(self.wiersz_danych, len(arkusz[col])+1):
                if len(self.id) == 2: # ID dla obrebu i budynku w osobnych kolumnach
                    id_obr = str(arkusz.cell(row=i, column=self.id[0]).value)
                    id_bud = str(arkusz.cell(row=i, column=self.id[1]).value)
                    id = '-'.join((id_obr, id_bud))
                elif len(self.id) == 1 and self.id != ['']: # id dla budynku w jednej kolumnie
                    id = arkusz.cell(row=i, column=self.id[0]).value
                else: # gdy ID nie jest określone id zgodne z numerem wiersza
                    id = str(i)
                data[id] = [str(arkusz.cell(row=i, column=k).value) for k in range(self.kolumna_danych, last_col)]
            header = []
            for x in range(self.wiersz_naglowka, self.wiersz_danych):
                header.append([str(arkusz.cell(row=x, column=k).value)  for k in range(self.kolumna_danych, last_col)])
        return data, header
         
if __name__ == "__main__": #pragma: no cover
    er = ExcelReader('./_example_data/raport_oldd.xlsx')
    print(er.read())