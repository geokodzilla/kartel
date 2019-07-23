# -*- coding: utf-8 -*-


class AppBlad(Exception):
    
    def __init__(self, code):
        self.code = code

    def __str__(self):
        return self.code
