# -*- coding: utf-8 -*-

import PySimpleGUI as sg
from excel_reader import ExcelReader
from excel_cmp import ExcelCmp
from excel_writer import ExcelWriter


def main_gui():
    layout = [ [sg.Txt('Program do porównywania plików XLSX')],      
               [sg.Txt('Numer arkusza:    '), sg.In(size=(5,1), key='nr_ark', default_text='1'),
               sg.Txt('Kolumna danych: '), sg.In(size=(5,1), key='k_dan', default_text='1')],        
               [sg.Txt('Wiersz nagłówka: '), sg.In(size=(5,1), key='w_nag', default_text='1'),
               sg.Txt('Wiersz danych:    '), sg.In(size=(5,1), key='w_dan', default_text='2')],      
               [sg.Txt('ID'), sg.In(size=(5,1), key='id', default_text='3,4')],
               [sg.Txt('Ścieżki do plików:')],
               [sg.Text('Plik oryginalny', size=(12, 3)), sg.InputText(size=(75,1)), sg.FileBrowse(button_text='Zmień')],      
               [sg.Text('Plik aktualny', size=(12, 3)), sg.InputText(size=(75,1)), sg.FileBrowse(button_text='Zmień')],      
               [sg.Submit(button_text='Porównaj')]]

    window = sg.Window('XLSX raport', layout)

    while True:      
        event, values = window.Read()

        if event is not None:
            NUMER_ARKUSZA = values['nr_ark']
            NAGLOWEK = values['w_nag']
            WIERSZ_DANYCH = values['w_dan']
            KOLUMNA_DANYCH = values['k_dan']
            ID = values['id'].split(',')
            old_file = values['Zmień']
            new_file = values['Zmień0']
            try:
                old_data, _ = ExcelReader(filename=old_file, numer_arkusza=NUMER_ARKUSZA,
                                        wiersz_naglowka=NAGLOWEK, wiersz_danych=WIERSZ_DANYCH,
                                        kolumna_danych=KOLUMNA_DANYCH, id=ID).read()
                new_data, header = ExcelReader(filename=new_file, numer_arkusza=NUMER_ARKUSZA,
                                        wiersz_naglowka=NAGLOWEK, wiersz_danych=WIERSZ_DANYCH,
                                        kolumna_danych=KOLUMNA_DANYCH, id=ID).read()
                xlsx_cmp = ExcelCmp(old_data, new_data)
                xlsx_cmp.cmp()
                zmiany = xlsx_cmo.mod
                xlsx_writer = ExcelWriter(header, zmiany)
                xlsx_writer.write('./raport.xlsx')
                xlsx_cmp.clear_mod()
            except AppBlad as e:
                sg.PopupError(e.code, title='BŁĄD')
            else:
                sg.Popup('Utworzono raport', title='INFO')
        else:      
            break

if __name__ == "__main__":

  main_gui()