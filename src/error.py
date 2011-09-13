#!/usr/bin/env python
# -*- coding: UTF-8 -*-

class Timeout(Exception):
    ''' Method or function has been running too much time '''

class ElementNotFound(Exception):
    ''' Can not find HTML element '''

if __name__ == '__main__':
    import  doctest
    doctest.testmod()
