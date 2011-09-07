#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import logging
from selenium import webdriver
from bridge import Bridge

logger = logging.getLogger('seleniumacros')

class Interface(object):
    ''' Provide similar interfaces to iMacros '''

    def __init__(self):
        self.bridge = Bridge()

    def iimInit(self, command, openNewBrowser=True, timeout=0):
        """
        Setup new browser driver.
        See http://wiki.imacros.net/iimInit%28%29 for more info.
        """
        # openNewBrowser and timeout paramenters are ignored currently
        #
        # FIXME Should we start browser right after interface is initialized?
        m = re.match(r'^-(\w+)(?:\s+(.*))?$', command)
        if not m:
            raise ValueError('Wrong command for iimInit')
        self.bridge.reset()
        self.bridge.set_browser(m.group(1))
        return True

    def iimPlay(self, macro, timeout=0):
        """
        Start browser process if it is not started yet.  Replay macro.
        See http://wiki.imacros.net/iimPlay%28%29 for more info.
        """
        return self.bridge.execute_script(macro, timeout)

    def iimSet(self, name, value):
        """
        Set user variable for next macro replay.
        Should be called before each iimPlay() call to take effect.
        See http://wiki.imacros.net/iimSet%28%29 for more info.
        """
        return self.bridge.set_variables({name: value})

    def iimDisplay(self, message, timeout=0):
        """
        Displays a message in the browser
        See http://wiki.imacros.net/iimDisplay%28%29 for more info.
        """
        # This feature is not supported yet
        logger.warn("Not supported yet")
        # Still return True since we don't want to break the execution
        return True

    def iimExit(self, timeout=0):
        """
        Closes browser instance.
        See http://wiki.imacros.net/iimExit%28%29 for more info.
        """
        self.bridge.reset()
        return True

    def iimGetLastError(self, index=-1):
        """
        Returns the text associated with the last error.
        See http://wiki.imacros.net/iimGetLastError%28%29 for more info.
        """
        try:
            return self.bridge.errors[index]
        except IndexError:
            return None

    def iimGetLastExtract(self, index=-1):
        """
        Returns the contents of the !EXTRACT variable.
        See http://wiki.imacros.net/iimGetLastExtract%28%29 for more info.
        """
        try:
            self.bridge.extracts[index]
        except IndexError:
            return None

    def iimTakeBrowserScreenshot(self, path, img_type, timeout=0):
        """
        Takes screenshot of browser or web page
        See http://wiki.imacros.net/iimTakeBrowserScreenshot%28%29 for more info.
        """
        raise NotImplementedError, "Not implemented yet"

    def iimGetLastPerformance(self, index=0):
        """
        Returns the total runtime and STOPWATCH data for the most recent macro run.
        Returned object is tuple where first value indicates data presence,
        second value is STOPWATCH name, third is STOPWATCH value.
        If index equals 1 then total runtime is returned.

        See http://wiki.imacros.net/iimGetLastPerformance for more info.
        """
        raise NotImplementedError, "Not implemented yet"
