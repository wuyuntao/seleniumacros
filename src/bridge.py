#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import logging
from selenium import webdriver

logger = logging.getLogger('seleniumacros')

class Bridge(object):
    ''' A bridge between iMacros and Selenium '''

    # Browser code
    IE      = 'ie'
    FIREFOX = 'fx'
    CHROME  = 'cr'
    OPERA   = 'op'        # Only supported by selenium

    # Selenium dirvers
    WEB_DRIVERS = {
        IE:      webdriver.Ie,
        FIREFOX: webdriver.Firefox,
        CHROME:  webdriver.Chrome,
        OPERA:   webdriver.Opera,
    }

    # Error codes
    OK      =  1
    FAIL    = -1          # Exception
    TIMEOUT = -3          # Timeout

    DEFAULT_TIMEOUT = 0   # Seconds to raise a timeout error. 0 means unlimited

    SUPPORTED_COMMANDS = (
        # 'ADD',
        # 'BACK',
        # 'CLEAR',
        # 'CLICK',
        'DS',
        # 'EXTRACT',
        # 'FILEDELETE',
        # 'FILTER',
        # 'FRAME',
        # 'IMAGECLICK',
        # 'IMAGESEARCH',
        # 'ONCERTIFICATEDIALOG',
        # 'ONDIALOG',
        # 'ONDOWNLOAD',
        # 'ONERRORDIALOG',
        # 'ONLOGIN',
        # 'ONPRINT',
        # 'ONSECURITYDIALOG',
        # 'ONWEBPAGEDIALOG',
        # 'PAUSE',
        # 'PRINT',
        # 'PROMPT',
        # 'PROXY',
        # 'REFRESH',
        # 'SAVEAS',
        # 'SAVEITEM',
        # 'SEARCH',
        'SET',
        'SIZE',
        # 'STOPWATCH',
        # 'TAB',
        'TAG',
        # 'TRAY',
        'URL',
        # 'VERSION',
        'WAIT',
    )

    SUPPORTED_BUILTIN_VARIABLES = (
        # Unsupported built-in variables
        # '!CLIPBOARD',
        # '!COL1',
        # '!COL2',
        # '!COL3',
        # '!DATASOURCE',
        # '!DATASOURCE_COLUMNS',
        # '!DATASOURCE_LINE',
        # '!ENCRYPTION',
        # '!ENDOFPAGE',
        # '!ERRORIGNORE',
        # '!EXTRACT',
        # '!EXTRACT_TEST_POPUP',
        # '!EXTRACTDIALOG',
        # '!FILELOG',
        # '!FILESTOPWATCH',
        # '!FOLDER_DATASOURCE',
        # '!FOLDER_STOPWATCH',
        # '!IMAGEX',
        # '!IMAGEY',
        # '!LOOP',
        # '!MARKOBJECT',
        # '!NOW',
        # '!POPUP_ALLOWED',
        # '!REPLAYSPEED',
        # '!REGION_BOTTOM',
        # '!REGION_LEFT',
        # '!REGION_RIGHT',
        # '!REGION_TOP',
        # '!SINGLESTEP',
        # '!STOPWATCHTIME',
        # '!STOPWATCH_HEADER',
        # '!TAGSOURCEINDEX',
        # '!TAGX',
        # '!TAGY',
        # '!TIMEOUT',
        # '!TIMEOUT_MACRO',
        # '!TIMEOUT_PAGE',
        # '!TIMEOUT_STEP',
        # '!URLCURRENT',
        # '!USERAGENT',
        # '!VAR1',
        # '!VAR2',
        # '!VAR3',
        # '!WAITPAGECOMPLETE',
    )

    DEFAULT_BUILTIN_VARIABLES = {
    }

    RE_COMMENT = re.compile(r'^\'\"\s*(.*)$')

    def __init__(self):
        self.browser = None
        self.driver = None
        self.built_variables = {}
        self.variables = {}
        self.errors = []
        self.extracts = []
        self.initialized = False

    def set_browser(self, browser=IE):
        if browser not in self.WEB_DRIVERS.keys():
            raise ValueError, "Invalid browser code: %s" % browser
        self.browser = browser
        return True

    def start_driver(self):
        if self.driver is None:
            logger.info("Starting driver")
            self.driver = self.WEB_DRIVERS[self.browser]()
        else:
            logger.info("Driver is already started")
        return True

    def reset(self):
        if self.driver is not None:
            self.driver.close()
            self.driver = None
        self.browser = None
        self.built_variables = {}
        # Set default values for built-in variables
        self.built_variables.update(DEFAULT_BUILTIN_VARIABLES)
        self.variables = {}
        self.errors = []
        self.extracts = []

    def set_builtin_variables(self, variables={}):
        for name, value in variables.items():
            if name in self.SUPPORTED_BUILTIN_VARIABLES:
                self.built_variables[name] = value

    def set_variables(self, variables={})
        self.variables.update(variables)
        return True

    def execute_script(self, script, timeout=DEFAULT_TIMEOUT)
        for command in open(script, 'rb').readlines():
            command = command.strip()
            # Ignore empty string
            if not command:
                continue

            # Handle comment command
            match = self.RE_COMMENT.match(command)
            if match:
                self.execute_comment(match.group(1))
                continue

            command = command.split()
            if command[0] in self.SUPPORTED_COMMANDS:
                if not getattr(self, 'execute_%s_command' % command[0].lower())(command[1:]):
                    return False
            else:
                self.execute_unsupported_command(command[0], command[1:])

    def execute_ds_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_set_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_size_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_tag_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_url_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_wait_command(self):
        raise NotImplementedError, "Not implemented yet."

    def execute_comment(self, comment):
        logger.info("Comment: %s" % comment)
        return True

    def execute_unsupported_command(self, command, arguments):
        logger.warn("This command is not supported yet. %s %s" % (command, ' '.join(arguments)))
        return True
