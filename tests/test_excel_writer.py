# -*- coding: utf-8 -*-

import pytest
from kartel.excel_writer import *
from kartel.excel_reader import *
from collections import OrderedDict

@pytest.fixture
def xlsx():
    header = [['Lp', 'ID_BUD', 'PKOB', 'GLFUN']]
    mod = [('M', ['2',	'5-100;2',	'1271',	'In -> Bg']),
                    ('U', ['3',	'5-100;3',	'1271', 'In']),
                    ('D', ['5',	'5-101;2', 	'1271', 'In'])]
    return ExcelWriter(header, mod)


def test_write(xlsx, tmpdir):
    file = tmpdir.join('raport.xlsx')
    xlsx.write(file)
    read_data, read_header = ExcelReader(file).read()
    for k, data in read_data.items():
        assert data in [i[1] for i in xlsx.mod]
    assert read_header == xlsx.header