#!/usr/bin/env python
# -*- coding: UTF-8 -*-

class Timeout(Exception):
    ''' Method or function has been running too much time '''

if __name__ == '__main__':
    import  doctest
    doctest.testmod()
