#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import re
import time
import logging
from selenium import webdriver

logger = logging.getLogger('seleniumacros')

class Bridge(object):
    ''' A bridge between iMacros and Selenium '''

    # Browser code
    IE      = 'ie'
    FIREFOX = 'fx'
    CHROME  = 'cr'

    # Selenium dirvers
    WEB_DRIVERS = {
        IE:      webdriver.Ie,
        FIREFOX: webdriver.Firefox,
        CHROME:  webdriver.Chrome,
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
        '!EXTRACT',
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
        '!TIMEOUT',
        '!TIMEOUT_MACRO',
        '!TIMEOUT_PAGE',
        '!TIMEOUT_STEP',
        # '!URLCURRENT',
        # '!USERAGENT',
        '!VAR1',
        '!VAR2',
        '!VAR3',
        # '!WAITPAGECOMPLETE',
    )

    DEFAULT_BUILTIN_VARIABLES = {
        # Set all timeout to 1 hour
        '!TIMEOUT':       '3600',
        '!TIMEOUT_MACRO': '3600',
        '!TIMEOUT_PAGE':  '3600',
        '!TIMEOUT_STEP':  '3600',
    }

    RE_COMMENT = re.compile(r'^\'\'\s*(.*)$')
    RE_X = re.compile(r'^X=(\d+)$')
    RE_Y = re.compile(r'^Y=(\d+)$')
    RE_VARIABLE = re.compile(r'^\{\{([0-9A-Z_]+)\}\}$')
    RE_BUILTIN_VARIABLE = re.compile(r'^\{\{(![0-9A-Z_]+)\}\}$')
    RE_VARIABLE_NAME = re.compile(r'^([0-9A-Z_]+)$')
    RE_BUILTIN_VARIABLE_NAME = re.compile(r'^(![0-9A-Z_]+)$')
    RE_QUOTED_STRING = re.compile(r'^[\'\"](.*)[\'\"]$')
    RE_SECONDS = re.compile(r'^SECONDS=(\d+)$')

    def __init__(self):
        self.reset()

    def set_browser(self, browser=IE):
        if browser not in self.WEB_DRIVERS.keys():
            raise ValueError, 'Invalid browser code: %s' % browser
        self.browser = browser

    def start_driver(self, force=False):
        if force or self.driver is None:
            logger.info('Starting driver')
            self.driver = self.WEB_DRIVERS[self.browser]()
        else:
            logger.info('Driver is already started')

    def set_builtin_variables(self, variables={}):
        for name, value in variables.items():
            if name in self.SUPPORTED_BUILTIN_VARIABLES:
                self.builtin_variables[name] = value

    def set_variables(self, variables={}):
        self.variables.update(variables)

    def reset(self):
        if getattr(self, 'driver', None):
            self.driver.close()
        self.driver = None
        self.browser = None
        self.builtin_variables = {}
        # Set default values for built-in variables
        self.builtin_variables.update(self.DEFAULT_BUILTIN_VARIABLES)
        self.variables = {}
        self.errors = []
        self.extracts = []

    def execute_script(self, script, timeout=DEFAULT_TIMEOUT):
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

            # FIXME
            # Here is a bug when you run:
            # SET !VAR1 "Hello World"
            command = command.split()
            if command[0] in self.SUPPORTED_COMMANDS:
                getattr(self, 'execute_%s_command' % command[0].lower())(*command[1:])
            else:
                self.execute_unsupported_command(command[0], *command[1:])

    def execute_ds_command(self):
        raise NotImplementedError, 'Not implemented yet.'

    def execute_set_command(self, name, value):
        '''
        SET !VAR1 TEST1
        SET !VAR2 {{TITLE}}
        SET !VAR3 "Hello World"
        SET !TIMEOUT Hello<SP>World 

        >>> bridge = Bridge()
        >>> bridge.execute_set_command('!VAR1', 'TEST1')
        >>> bridge.builtin_variables['!VAR1']
        'TEST1'
        >>> bridge.execute_set_command('!VAR2', '{{!VAR1}}')
        >>> bridge.builtin_variables['!VAR2']
        'TEST1'

        '''
        # TODO
        # support setting Non-built-in variables?
        if not self.RE_BUILTIN_VARIABLE_NAME.match(name):
            raise ValueError, "Wrong name format"
        self.builtin_variables[name] = self._parse_value_string(value)

    def execute_size_command(self, x, y):
        '''
        SIZE X=1024 Y=768

        >>> bridge = Bridge()
        >>> bridge.set_browser(Bridge.FIREFOX)
        >>> bridge.start_driver()
        >>> bridge.execute_size_command("X=300", "Y=600")

        '''
        if not (self.RE_X.match(x) and self.RE_Y.match(y)): 
            raise ValueError, "Wrong argument format"
        x, y = int(x[2:]), int(y[2:])
        # This changes the size of whole firefox window, not only viewport
        self.driver.execute_script('window.resizeTo(%s,%s)' % (x, y))

    def execute_tag_command(self, *args):
        """
        TAG POS=1 FORM=ID:login ATTR=NAME:email

        """
        # TODO
        # Support relative position value
        # Support extract data
        pos, type, form, attrs, content, extract = self._parse_tag_arguments(*args)
        element = self._find_element_by(pos, type, form, attrs)
        if content:
            element.clear()
            element.send_keys(content)
        elif extract:
            logger.warn("Extract is not supported yet. Will trigger a click instead")
            element.click()
        else:
            element.click()

    def execute_url_command(self, goto):
        '''
        URL GOTO=http://www.google.com

        '''
        if not goto.startswith('GOTO='):
            raise ValueError, "Invalid argument format"
        self.driver.get(goto[5:])

    def execute_wait_command(self, seconds):
        match = self.RE_SECONDS.match(seconds)
        if not match:
            raise ValueError, "Invalid argument format"
        time.sleep(int(match.group(1)))

    def execute_comment(self, comment):
        logger.info('Comment: %s' % comment)

    def execute_unsupported_command(self, command, *args):
        logger.warn('This command is not supported yet. %s %s' % (command, ' '.join(args)))

    # Private methods
    def _strip_argument(self, string):
        return string.split('=', 1)[1]

    def _parse_value_string(self, string):
        '''

        >>> bridge = Bridge()
        >>> bridge._parse_value_string('TEST1')
        'TEST1'
        >>> bridge._parse_value_string('"TEST1"')
        'TEST1'
        >>> bridge.set_variables({'TITLE': 'TEST1'})
        >>> bridge._parse_value_string('{{TITLE}}')
        'TEST1'
        >>> bridge.set_builtin_variables({'!VAR1': 'TEST1'})
        >>> bridge._parse_value_string('{{!VAR1}}')
        'TEST1'
        >>> bridge.set_builtin_variables({'!VAR2': 'Hello<SP>World'})
        >>> bridge._parse_value_string('{{!VAR2}}')
        'Hello World'
        >>> bridge.set_builtin_variables({'!TIMEOUT': '360'})
        >>> bridge._parse_value_string('{{!TIMEOUT}}')
        360

        '''
        # TODO
        # Handle string contains variables, e.g. "{{!VAR1}}<SP>Hello World"

        # Check if string is a custom variables
        match = self.RE_VARIABLE.match(string)
        # print match, string
        if match:
            string = self.variables.get(match.group(1), '')
        else:
            # Check if string is a built-in variables
            match = self.RE_BUILTIN_VARIABLE.match(string)
            if match:
                string = self.builtin_variables.get(match.group(1), '')
            else:
                # Check if string is quoted
                match = self.RE_QUOTED_STRING.match(string)
                if match:
                    string = match.group(1)

        # Escape string from iMacros specific chars
        return int(string) if string.isdigit() else self._escape_string(string)

    def _escape_string(self, string):
        return string.replace('<SP>', ' ').replace('<BR>', '\n')

    def _parse_tag_arguments(self, *args):
        pos, type, form, attrs, content, extract = 1, 'div', None, {}, False, False
        args = [arg.split('=', 1) for arg in args]
        for name, value in args:
            if name == 'POS':
                pos = int(value)
            elif name == 'TYPE':
                type = self._parse_html_element_type(value)
            elif name == 'FORM':
                form = self._parse_html_attributes(value)
            elif name == 'ATTR':
                attrs = self._parse_html_attributes(value)
            elif name == 'CONTENT':
                content = self._parse_value_string(value)
            elif name == 'EXTRACT':
                extract = value
            else:
                raise ValueError, "Invalid tag argument"
        return pos, type, form, attrs, content, extract

    def _parse_html_element_type(self, string):
        if string.startswith('INPUT:'):
            return 'input[type=%s]' % string[6:].lower()
        else:
            return string.lower()

    def _parse_html_attributes(self, string):
        attributes = {}
        for attr in string.split('&&'):
            name, value = attr.split(':')
            attributes[name.lower()] = self._parse_value_string(value)
        return attributes

    def _find_element_by(self, pos, type, form, attrs):
        # Assume elements can only have unique id
        if 'id' in attrs:
            return self.driver.find_element_by_id(attrs['id'])

        css_selector = type
        text = attrs.pop('txt', None)
        if text is not None:
            css_selector += ':contains(%s)' % text
        for name, value in attrs.items():
            css_selector += '[%s="%s"'] % (name.lower(), value)

        if form:
            form = self._find_element_by(1, 'form', None, form)
            element = form.find_elements_by_css_selector(css_selector)[pos - 1]
        else:
            element = self.driver.find_elements_by_css_selector(css_selector)[pos - 1]
        return element



if __name__ == '__main__':
    import  doctest
    doctest.testmod()