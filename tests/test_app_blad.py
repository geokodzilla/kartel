# -*- coding: utf-8 -*-

import pytest
from kartel.app_blad import *

@pytest.fixture
def err():
    return AppBlad('Prosty błąd')
    
    
def test_app_blad_code(err):
    assert err.code == 'Prosty błąd'
    err.code = 'Trudny błąd'
    assert err.code == 'Trudny błąd'
    assert str(err) == 'Trudny błąd'