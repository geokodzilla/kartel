# -*- coding: utf-8 -*-

import pytest
from kartel.excel_reader import *
from collections import OrderedDict

@pytest.fixture
def xlsx():
    return ExcelReader('./_example_data/raport_old.xlsx', id=['2'])


def test_data_validation():
    with pytest.raises(AppBlad) as e:
        ExcelReader('./_example_data/raport_old.xlsx', wiersz_naglowka='aa')
    assert e.value.code == 'Błędne parametry wczytania pliku'

def test_file_validation():
    with pytest.raises(AppBlad) as e:
        ExcelReader('./_example_data/raport_older.xlsx')
    assert e.value.code == 'Wskazany plik nie istniej ./_example_data/raport_older.xlsx'

def test_read(xlsx):
    data, header = xlsx.read()
    assert data == OrderedDict([('5-100;1', ['1', '5-100;1', '1110', 'Dj']),
                                                                ('5-100;2', ['2', '5-100;2', '1271', 'In']),
                                                                ('5-100;3', ['3', '5-100;3', '1271', 'In']),
                                                                ('5-101;1', ['4', '5-101;1', '1110', 'Dj'])])
    assert header == [['Lp', 'ID_BUD', 'PKOB', 'GLFUN']]
    
def test_kolumny_poza_zakresem():
    with pytest.raises(AppBlad) as e:
            ExcelReader('./_example_data/raport_old.xlsx', kolumna_danych='9').read()
    assert e.value.code == 'Kolumny poza zakresem'
            
            
@pytest.mark.parametrize("test_input,expected", [([''], OrderedDict([('1', ['1', '5-100;1', '1110', 'Dj']),
                                                                ('2', ['2', '5-100;2', '1271', 'In']),
                                                                ('3', ['3', '5-100;3', '1271', 'In']),
                                                                ('4', ['4', '5-101;1', '1110', 'Dj'])])),
                                                 (['2'], OrderedDict([('5-100;1', ['1', '5-100;1', '1110', 'Dj']),
                                                                ('5-100;2', ['2', '5-100;2', '1271', 'In']),
                                                                ('5-100;3', ['3', '5-100;3', '1271', 'In']),
                                                                ('5-101;1', ['4', '5-101;1', '1110', 'Dj'])])),
                                                  (['1', '2'], OrderedDict([('1-5-100;1', ['1', '5-100;1', '1110', 'Dj']),
                                                                ('2-5-100;2', ['2', '5-100;2', '1271', 'In']),
                                                                ('3-5-100;3', ['3', '5-100;3', '1271', 'In']),
                                                                ('4-5-101;1', ['4', '5-101;1', '1110', 'Dj'])]))
                                                             ])
def test_id(test_input, expected):
    data, _ = ExcelReader('./_example_data/raport_old.xlsx', id=test_input).read()
    assert data == expected
    