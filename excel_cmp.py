# -*- coding: utf-8 -*-

class ExcelCmp(object):
    
    def __init__(self, old_data, new_data):
        self.old_data = old_data
        self.new_data = new_data
        self.mod = []
        
        
    def cmp(self):
        """
        metoda porównuje dane oryginalne z danymi aktualnymi wyszukjuąc wg klucza
        ustalonego przy wczytywaniu wiersze dodane usunięte oraz zmodyfikowane
        dane do zapisu w postaci listy krotek, pierwszy element krotki informuje
        o rodzaju zmiany
        U - usunięte
        D - dodane
        M - modyfikacja
        """
        for k in self.old_data:
            if k not in self.new_data:
                self.mod.append(('U', self.old_data[k]))
        for k, v in self.new_data.items():
            if k not in self.old_data:
                self.mod.append(('D', self.new_data[k]))
            else:
                c = 0
                for nr, i in enumerate(v):
                    if i != self.old_data[k][nr]:
                        v[nr] = '%s -> %s' % (self.old_data[k][nr], i)
                        v[nr] = v[nr].replace('None' , '')
                        c += 1
                if c != 0:
                    self.mod.append(('M', v))
                    
                    
    def clear_mod(self):
        self.mod = []