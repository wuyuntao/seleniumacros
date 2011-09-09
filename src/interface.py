#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import re
import logging
from selenium import webdriver
from bridge import Bridge

logger = logging.getLogger('seleniumacros')

def handle_retcode(func):
    def wrap(*args, **kwargs):
        retcode = func(*args, **kwargs)
        if retcode is None or retcode is True:
            return Bridge.OK
        elif retcode is False:
            return Bridge.FAIL
        else:
            return int(retcode)
    return wrap

class Interface(object):
    '''
    Provide similar interfaces to iMacros
    
    >>> imacros = Interface()

    # >>> imacros.iimInit("-ie")
    # 1
    # >>> imacros.iimPlay("Z:/iMacrosDemo/FillForm.iim")
    # 1

    >>> imacros.iimInit("-fx")
    1
    >>> imacros.iimPlay("/home/yuntao/Downloads/iMacrosDemo/FillForm.iim")
    1

    # >>> imacros.iimExit()
    # 1

    '''

    RE_INIT_COMMAND = re.compile(r'^-(\w+)(?:\s+(.*))?$')

    def __init__(self):
        self.bridge = Bridge()

    @handle_retcode
    def iimInit(self, command, openNewBrowser=True, timeout=False):
        """
        Setup new browser driver.
        See http://wiki.imacros.net/iimInit%28%29 for more info.
        """
        # openNewBrowser is ignored, since we always open new window
        match = re.match(self.RE_INIT_COMMAND, command)
        if not match:
            raise ValueError('Wrong command for iimInit')
        self.bridge.set_browser(match.group(1))
        self.bridge.start_driver()
        if timeout > 0:
            self.bridge.set_builtin_variables({'!TIMEOUT_MACRO': int(timeout)})

    @handle_retcode
    def iimPlay(self, macro, timeout=False):
        """
        Start browser process if it is not started yet.  Replay macro.
        See http://wiki.imacros.net/iimPlay%28%29 for more info.
        """
        if timeout > 0:
            self.bridge.set_builtin_variables({'!TIMEOUT': int(timeout)})
        return self.bridge.execute_script(macro)

    @handle_retcode
    def iimSet(self, name, value):
        """
        Set user variable for next macro replay.
        Should be called before each iimPlay() call to take effect.
        See http://wiki.imacros.net/iimSet%28%29 for more info.
        """
        self.bridge.set_variables({name: value})

    @handle_retcode
    def iimDisplay(self, message, timeout=0):
        """
        Displays a message in the browser
        See http://wiki.imacros.net/iimDisplay%28%29 for more info.
        """
        # This feature is not supported yet
        logger.warn("Not supported yet")

    @handle_retcode
    def iimExit(self, timeout=0):
        """
        Closes browser instance.
        See http://wiki.imacros.net/iimExit%28%29 for more info.
        """
        self.bridge.reset()

    @handle_retcode
    def iimGetLastError(self, index=-1):
        """
        Returns the text associated with the last error.
        See http://wiki.imacros.net/iimGetLastError%28%29 for more info.
        """
        return self.bridge.errors[index]

    @handle_retcode
    def iimGetLastExtract(self, index=-1):
        """
        Returns the contents of the !EXTRACT variable.
        See http://wiki.imacros.net/iimGetLastExtract%28%29 for more info.
        """
        return self.bridge.extracts[index]

    @handle_retcode
    def iimTakeBrowserScreenshot(self, path, img_type, timeout=0):
        """
        Takes screenshot of browser or web page
        See http://wiki.imacros.net/iimTakeBrowserScreenshot%28%29 for more info.
        """
        raise NotImplementedError, "Not implemented yet"

    @handle_retcode
    def iimGetLastPerformance(self, index=0):
        """
        Returns the total runtime and STOPWATCH data for the most recent macro run.
        Returned object is tuple where first value indicates data presence,
        second value is STOPWATCH name, third is STOPWATCH value.
        If index equals 1 then total runtime is returned.

        See http://wiki.imacros.net/iimGetLastPerformance for more info.
        """
        raise NotImplementedError, "Not implemented yet"

if __name__ == '__main__':
    import  doctest
    doctest.testmod()
