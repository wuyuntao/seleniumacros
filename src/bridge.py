#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import re
import time
import logging
import platform
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from error import Timeout, ElementNotFound

logger = logging.getLogger('seleniumacros')

try:
    import win32com.client
    import pywintypes
    using_autoit = True
except ImportError:
    if platform.platform().startswith('Windows'):
        raise ImportError, 'This program requires the pywin32 extensions for Python.'
    else:
        # We don't use AutoIt on other platforms
        using_autoit = False
        # TODO Use AutoKey on Linux

class Bridge(object):
    ''' A bridge between iMacros and Selenium '''
    # TODO
    # Handle native dialog from IE or Firefox, e.g. save password dialog

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
        '!REPLAYSPEED',
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
        '!REPLAYSPEED':   'MEDIUM',
    }

    # Seconds to wait between commands
    REPLAYSPEED_FAST = 0
    REPLAYSPEED_MEDIUM = 0.25
    REPLAYSPEED_SLOW = 1

    RE_COMMENT = re.compile(r'^\'\s*(.*)$')
    RE_X = re.compile(r'^X=(\d+)$')
    RE_Y = re.compile(r'^Y=(\d+)$')
    RE_VARIABLE = re.compile(r'^\{\{([0-9A-Z_]+)\}\}$')
    RE_BUILTIN_VARIABLE = re.compile(r'^\{\{(![0-9A-Z_]+)\}\}$')
    RE_VARIABLE_NAME = re.compile(r'^([0-9A-Z_]+)$')
    RE_BUILTIN_VARIABLE_NAME = re.compile(r'^(![0-9A-Z_]+)$')
    RE_QUOTED_STRING = re.compile(r'^[\'\"](.*)[\'\"]$')
    RE_SECONDS = re.compile(r'^SECONDS=(\d+)$')
    # TODO Support more direct screen events
    # RE_DS_CMD = re.compile(r'^CMD=(CLICK|LDBLCLK|LDOWN|LUP|MOVETO|MDOWN|MUP|MDBLCLK|RDOWN|RUP|RDBLCLK|KEY)$')
    RE_DS_CMD = re.compile(r'^CMD=(CLICK|KEY)$')


    def __init__(self):
        self.reset()

    def set_browser(self, browser=IE):
        if browser not in self.WEB_DRIVERS.keys():
            error = 'Invalid browser code: %s' % browser
            logger.error(error)
            raise ValueError, error
        self.browser = browser

    def start_driver(self, force=False):
        if force or self.driver is None:
            logger.info(u'Starting driver')
            self.driver = self.WEB_DRIVERS[self.browser]()

            if using_autoit:
                # Set unique window title to get handle for AutoIT
                self.driver.execute_script(u'document.title = "%s"' % \
                        self.driver.current_window_handle)
                self.autoit = win32com.client.Dispatch('AutoItX3.Control')
                self.autoit_handle = \
                        self.autoit.WinWait(self.driver.current_window_handle)
                self.autoit.WinActivate(self.autoit_handle)
        else:
            logger.info(u'Driver is already started')

    def set_builtin_variables(self, variables={}):
        for name, value in variables.items():
            if name in self.SUPPORTED_BUILTIN_VARIABLES:
                self.builtin_variables[name] = value
            else:
                logger.warn(u'built-in variable %s is not supported yet.')

    def set_variables(self, variables={}):
        self.variables.update(variables)

    def reset(self):
        if getattr(self, 'driver', None):
            self.driver.close()
        self.driver = None
        self.browser = None
        self.autoit = None
        self.autoit_handle = None
        self.builtin_variables = {}
        # Set default values for built-in variables
        self.builtin_variables.update(self.DEFAULT_BUILTIN_VARIABLES)
        self.variables = {}
        self.errors = []
        self.extracts = []

    def execute_script(self, script, timeout=DEFAULT_TIMEOUT):
        for command in open(script, 'rb').readlines():
            command = unicode(command.strip(), 'utf-8')
            # Ignore empty string
            if not command:
                continue

            # Escape command string.
            logger.info(u'Execute command: %s' % self._escape_string(command))
            # Handle comment command
            match = self.RE_COMMENT.match(command)
            if match:
                self.execute_comment(match.group(1))
                continue

            # FIXME
            # Here is a bug when you run:
            # SET !VAR1 'Hello World'
            # So we need a regexp to split tokens in the futrue
            command = command.split()
            if command[0] in self.SUPPORTED_COMMANDS:
                # try:
                #     getattr(self, 'execute_%s_command' % command[0].lower())(*command[1:])
                # except Exception, e:
                #     logger.error(e)
                #     self.errors.append(e)
                #     return False
                getattr(self, 'execute_%s_command' % command[0].lower())(*command[1:])
            else:
                self.execute_unsupported_command(command[0], *command[1:])
            self._replay_wait()

    def execute_ds_command(self, cmd, *args):
        '''
        Since Selenium itself does not support Direct Screen Tech used in iMacros,
        We use AutoIt to simulate OS-Level mouse and keyboard events.

        DS CMD=CLICK X=340 Y=410 CONTENT=
        DS CMD=KEY CONTENT=notepad.exe

        '''
        # TODO Use AutoKey API to simulate user inputs under Linux
        if self.autoit_handle is None:
            raise ValueError, 'AutoIt must be installed to use direct screen commands'

        match = self.RE_DS_CMD.match(cmd)
        if not match:
            raise ValueError, 'Wrong arugment format'
        cmd = match.group(1)
        if cmd == 'KEY':
            # It's a keyboard event
            content = self._parse_value_string(args)
            self.autoit.Send(content)
        else:
            # It's a mouse event
            x, y, content = args
            if not (self.RE_X.match(x) and self.RE_Y.match(y)): 
                raise ValueError, 'Wrong argument format'
            x, y = int(x[2:]), int(y[2:])

            # For IE we can send mouse events directly to page frame
            if self.browser == self.IE:
                self.autoit.ControlClick(self.autoit_handle, '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', 1, x, y)
            elif self.browser == self.Chrome:
                self.autoit.ControlClick(self.autoit_handle, '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', 1, x, y)
            else:
                # FIXME
                # Since we have not found a way to get border size of firefox window,
                # so we have to estimate the coordinate of screen to click which 
                # might be not very precise.
                # TODO Hide Add-on Bar to improve precise
                raise NotImplementedError, 'Not implemented yet'

    def execute_set_command(self, name, value):
        '''
        SET !VAR1 TEST1
        SET !VAR2 {{TITLE}}
        SET !VAR3 'Hello World'
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
            raise ValueError, 'Wrong name format'
        self.builtin_variables[name] = self._parse_value_string(value)

    def execute_size_command(self, x, y):
        '''
        SIZE X=1024 Y=768

        >>> bridge = Bridge()
        >>> bridge.set_browser(Bridge.FIREFOX)
        >>> bridge.start_driver()
        >>> bridge.execute_size_command('X=300', 'Y=600')
        >>> bridge.execute_wait_command('SECONDS=5')
        >>> bridge.reset()

        '''
        # TODO Need to add asserts to resized window
        if not (self.RE_X.match(x) and self.RE_Y.match(y)): 
            raise ValueError, u'Wrong argument format'
        x, y = int(x[2:]), int(y[2:])
        # This changes the size of whole firefox window, not only viewport
        self.driver.execute_script('window.resizeTo(%s,%s)' % (x, y))

    def execute_tag_command(self, *args):
        '''
        TAG POS=1 FORM=ID:login ATTR=NAME:email

        >>> bridge = Bridge()
        >>> bridge.set_browser(Bridge.FIREFOX)
        >>> bridge.start_driver()

        >>> bridge.execute_url_command('GOTO=http://www.iopus.com/imacros/support/html2tag.htm')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=A', 'ATTR=HREF:http://www.iopus.com')
        >>> bridge.execute_wait_command('SECONDS=5')

        >>> bridge.execute_url_command('GOTO=http://www.iopus.com/imacros/support/html2tag.htm')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=A', 'ATTR=ID:myLinkID')
        >>> bridge.execute_wait_command('SECONDS=5')

        >>> bridge.execute_url_command('GOTO=http://www.iopus.com/imacros/support/html2tag.htm')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=A', 'ATTR=NAME:myLinkName')
        >>> bridge.execute_wait_command('SECONDS=5')

        >>> bridge.execute_url_command('GOTO=http://www.iopus.com/imacros/support/html2tag.htm')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=STRONG', 'ATTR=TXT:<SP>iMacros<SP>User<SP>Forum')
        >>> bridge.execute_wait_command('SECONDS=5')

        >>> bridge.execute_url_command('GOTO=http://www.iopus.com/imacros/support/html2tag.htm')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=INPUT:TEXT', 'FORM=NAME:F1', 'ATTR=NAME:tf1', 'CONTENT=Hello<SP>World')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=INPUT:CHECKBOX', 'FORM=NAME:F1', 'ATTR=NAME:cb1&&ID:cb1', 'CONTENT=YES')
        >>> bridge.execute_tag_command('POS=1', 'TYPE=INPUT:RADIO', 'FORM=NAME:F1', 'ATTR=ID:r1', 'CONTENT=YES')

        '''
        # TODO
        # Support relative position value
        # Support extract data
        # Provide better test experience
        # Need to add asserts to clicked links

        pos, type, form, attrs, content, extract = self._parse_tag_arguments(*args)
        try:
            element = self._find_element_by(pos, type, form, attrs)
        except IndexError:
            raise ElementNotFound, u'Can not find HTML element by ' + u' '.join(args)
        # TODO Save element coordinate in !TAGX and !TAGY
        if content:
            # Handle form controls
            if element.tag_name in ('input', 'textarea'):
                input_type = element.get_attribute('type')
                if input_type in ('checkbox', 'radio'):
                    has_to_be_selected = content == 'YES'
                    if element.is_selected() != has_to_be_selected:
                        element.click()
                else:
                    # FIXME
                    # Clearing value for file input does not work on IE
                    # Issue: http://code.google.com/p/selenium/issues/detail?id=2370
                    if input_type != 'file':
                        element.clear()
                    value = self._parse_value_string(content)
                    logger.debug(value)
                    element.send_keys(value)

            elif element.tag_name == 'select':
                options = content.split(':')

                if element.get_attribute('multiple') in ('multiple', 'on', 'true', 'yes'):
                    # If element is a multiple select, hold CTRL key and click all options
                    # FIXME
                    # Will raise an 'Unrecognized command' exception under Firefox:
                    # http://code.google.com/p/selenium/issues/detail?id=1427
                    chain = webdriver.ActionChains(self.driver)
                    action = chain.key_down(Keys.CONTROL)
                    for option in options:
                        action = action.click(self._find_option_by(element, option))
                    action.key_up(Keys.CONTROL).perform()

                else:
                    # If element is a dropdown select, just click the last option
                    try:
                        option = self._find_option_by(element, options[-1])
                        option.click()
                    except IndexError:
                        # Do nothing if content is empty
                        pass

        elif extract:
            logger.warn(u'Extract is not supported yet. Will trigger a click instead')
            element.click()
        else:
            element.click()

    def _find_option_by(self, element, option):
        if option.startswith('%%'):
            # Use value attribute to find option
            option = element.find_element_by_css_selector(u'option[value="%s"]' \
                    % self._parse_value_string(option[2:].strip()))
        else:
            # Use text to find option
            text = self._parse_value_string((option[1:] if option.startswith('$') else option).strip())
            option = [option for option in element.find_elements_by_tag_name('option') \
                    if option.text == text][0]
        return option

    def execute_url_command(self, goto):
        '''
        URL GOTO=http://www.google.com

        '''
        if not goto.startswith('GOTO='):
            raise ValueError, 'Invalid argument format'
        url = goto[5:]
        logger.info(u'Go to URL %s' % self._escape_string(url))
        self.driver.get(url)

    def execute_wait_command(self, seconds):
        match = self.RE_SECONDS.match(seconds)
        if not match:
            raise ValueError, 'Invalid argument format'
        seconds = int(match.group(1))
        logger.info(u'Wait for %s seconds' % seconds)
        time.sleep(seconds)

    def execute_comment(self, comment):
        logger.info(u'Comment: %s' % self._escape_string(comment))

    def execute_unsupported_command(self, *args):
        logger.warn(u'This command is not supported yet. %s' \
                % self._escape_string(' '.join(args)))

    # Private methods
    def _strip_argument(self, string):
        return string.split('=', 1)[1]

    def _parse_value_string(self, string):
        '''

        >>> bridge = Bridge()
        >>> bridge._parse_value_string('TEST1')
        'TEST1'
        >>> bridge._parse_value_string(''TEST1'')
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
        # Handle string contains variables, e.g. '{{!VAR1}}<SP>Hello World'

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
        return string.replace('<SP>', ' ').replace('<BR>', '\n').replace('%', '%%')

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
                raise ValueError, 'Invalid tag argument'
        return pos, type, form, attrs, content, extract

    def _parse_html_element_type(self, string):
        if string.startswith('INPUT:'):
            return u'input[type=%s]' % string[6:].lower()
        else:
            return string.lower()

    def _parse_html_attributes(self, string):
        attributes = {}
        if string:
            for attr in string.split('&&'):
                if ':' in attr:
                    # Both name and value, e.g. NAME:email
                    name, value = attr.split(':', 1)
                    attributes[name.lower()] = self._parse_value_string(value)
                elif attr != '*':
                    # Only attribute name, e.g. NAME
                    attributes[attr.lower()] = True
        return attributes

    def _find_element_by(self, pos, type, form, attrs):
        # Assume elements can only have unique id
        if 'id' in attrs:
            return self.driver.find_element_by_id(attrs['id'])

        css_selector = type
        text = attrs.pop('txt', None)
        # Selenium WebDriver does not support Sizzle-style pseudo-selectors like :contains well
        # So we fallback to standard method
        # if text is not None:
        #     css_selector += ':contains("%s")' % text
        for name, value in attrs.items():
            css_selector += '[%s]' % name.lower() \
                    if value is True \
                    else '[%s="%s"]' % (name.lower(), self._escape_string(value))
        logger.debug(u'css selector: %s' % self._escape_string(css_selector))

        if form:
            form = self._find_element_by(1, 'form', None, form)
            elements = form.find_elements_by_css_selector(css_selector)
            # raise ValueError, 'find form? %s\ncss_selector? %s\nfind elements? %s' % (form, css_selector, len(elements))
        else:
            elements = self.driver.find_elements_by_css_selector(css_selector)
        # FIXME
        # Selenium sometimes returns None if can not find any matching elements
        # So we have to make sure it always return a list
        elements = elements or []
        if text is not None:
            # Selenium will strip element text automatically
            text = text.strip()
            elements = [element for element in elements if element.text == text]
        return elements[pos - 1]

    def _replay_wait(self):
        try:
            time.sleep(getattr(self, 'REPLAYSPEED_%s' % \
                    self.builtin_variables['!REPLAYSPEED']))
        except AttributeError:
            raise ValueError, 'Wrong value for !REPLAYSPEED'

if __name__ == '__main__':
    import  doctest
    doctest.testmod()
