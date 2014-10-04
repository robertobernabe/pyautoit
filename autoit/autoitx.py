from __future__ import absolute_import, division, print_function
__author__ = 'florian.schaeffeler'
import win32com.client
import pywintypes
"""
Before you can use the COM interface to AutoItX it needs to be "registered" (This is done automatically when you install
the full version of AutoIt but you may need to do it manually if you are using AutoItX seperately).
To register the COM interface:
    1. Open a command prompt
    2. Change directory (using CD) to the directory that contains AutoItX3.dll
    3. Type regsvr32.exe AutoItX3.dll and press enter

The name of the AutoItX control is AutoItX3.Control
"""


class AutoItX3():
    """A simple wrapper for the AutoItX COM interface

    For Accessing controls by HANDLE you must specify "[HANDLE: <handle>]" as value for title.


    SW_HIDE Hides       the window and activates another window.
    SW_MAXIMIZE         Maximizes the specified window.
    SW_MINIMIZE         Minimizes the specified window and activates the next top-level window in the Z order.
    SW_RESTORE          Activates and displays the window. If the window is minimized or maximized, the system restores
                        it to its original size and position. An application should specify this flag when restoring a
                        minimized window.
    SW_SHOW             Activates the window and displays it in its current size and position.
    SW_SHOWDEFAULT      Sets the show state based on the SW_ value specified by the program that started the
                        application.
    SW_SHOWMAXIMIZED    Activates the window and displays it as a maximized window.
    SW_SHOWMINIMIZED    Activates the window and displays it as a minimized window.
    SW_SHOWMINNOACTIVE  Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the
                        window is not activated.
    SW_SHOWNA           Displays the window in its current size and position. This value is similar to SW_SHOW, except
                        the window is not activated.
    SW_SHOWNOACTIVATE   Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL,
                        except the window is not actived.
    SW_SHOWNORMAL       Activates and displays a window. If the window is minimized or maximized, the system restores
                        it to its original size and position. An application should specify this flag when displaying
                        the window for the first time.
    """
    try:
        _aux3 = win32com.client.Dispatch("AutoItX3.Control")
    except pywintypes.com_error as e:
        print("Could not bind AutoItX, call failed with code %d: %s" % (e.hr, e.msg))
        if e.exc is None:
            print("There is no extended error information")
        else:
             wcode, source, text, helpFile, helpId, scode = e.exc
             print("The source of the error is", source)
             print("The error message is", text)
             print("More info can be found in %s (id=%d)" % (helpFile, helpId))
        raise Exception("Could not bind AutoItX, you may have to register AutoItX.dll\n%s" % 'regsvr32.exe "<path_to>\AutoItX3.dll"')
    SW_HIDE = _aux3.SW_HIDE
    SW_MAXIMIZE = _aux3.SW_MAXIMIZE
    SW_MINIMIZE = _aux3.SW_MINIMIZE
    SW_RESTORE = _aux3.SW_RESTORE
    SW_SHOW = _aux3.SW_SHOW
    SW_SHOWDEFAULT = _aux3.SW_SHOWDEFAULT
    SW_SHOWMAXIMIZED = _aux3.SW_SHOWMAXIMIZED
    SW_SHOWMINIMIZED = _aux3.SW_SHOWMINIMIZED
    SW_SHOWMINNOACTIVE = _aux3.SW_SHOWMINNOACTIVE
    SW_SHOWNA = _aux3.SW_SHOWNA
    SW_SHOWNOACTIVATE = _aux3.SW_SHOWNOACTIVATE
    SW_SHOWNORMAL = _aux3.SW_SHOWNORMAL
    LOWEST_INT = -2147483647
    LEFT = "left"
    RIGHT = "right"
    MIDDLE = "middle"

    def __init__(self):
        pass

    @property
    def error(self):
        """Status of the error flag (equivalent to the @error macro in AutoIt v3)
        :rtype: int
        """
        return self._aux3.error

    @property
    def version(self):
        """autoitX dll version (equivalent to @autoitversion macro in AutoIt v3)
        :rtype: unicode
        """
        return self._aux3.version

    def auto_it_set_option(self, option, param):
        """Changes the operation of various AutoIt functions/parameters."""
        return self._aux3.AutoItSetOption(option, param)

    def block_input(self):
        """BlockInput Disable/enable the mouse and keyboard."""
        return self._aux3.block_input()

    def cd_tray(self, drive, status):
        """Opens or closes the CD tray.
        :rtype: int"""
        return self._aux3.CDTray(drive, status)

    def clip_get(self):
        """Retrieves text from the clipboard.
        sets error to 1 if clipboard is empty or contains a non-text entry.
        :returns: On Success str containing the text on the clipboard
        :rtype: str
        """
        return self._aux3.ClipGet()

    def clip_put(self, value):
        """Writes text to the clipboard.
        Any existing clipboard contents are overwritten.
        :param value: The text to write to the clipboard.
        :type value: str
        :returns: 1 Success 0 Failure
        :rtype: int
        """
        return self._aux3.ClipPut(value)

    def control_click(self, title, text, controlId, button=LEFT, clicks=1, x=LOWEST_INT, y=LOWEST_INT):
        """Sends a mouse click command to a given control.
        Some controls will resist clicking unless they are the active window. Use the WinActive() function to force the
        control's window to the top before using ControlClick().
        Using 2 for the number of clicks will send a double-click message to the control - this can even be used to
        launch programs from an explorer control!
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access.
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :param button: left, right or middle
        :param clicks: The number of times to click the mouse. Default is 1.
        :param x: The x position to click within the control. Default is center.
        :param y: The y position to click within the control. Default is center.
        :return: 1 Success 0 Failure
        :rtype: int
        """
        return self._aux3.ControlClick(title, text, controlId, button, clicks, x, y)

    def control_command(self, title, text, controlId, command, option):
        """Sends a command to a control.
        When using text instead of ClassName# in "Control" commands, be sure to use the entire text of the control.
        Partial text will fail.
        Certain commands that work on normal Combo and ListBoxes do not work on "ComboLBox" controls.
        When using a control name in the Control functions, you need to add a number to the end of the name to indicate
        which control. For example, if there two controls listed called "MDIClient", you would refer to these as
        "MDIClient1" and "MDIClient2". Use AU3_Spy.exe to obtain a control's number.
        :param title: The title of the window to access.
        :param text: The text of the window to access.
        :param controlId: The control to interact with.
        :param command: The command to send to the control.
        :param option: Additional parameter required by some commands; use "" if parameter is not required.
        """
        return self._aux3.ControlCommand(title, text, controlId, command, option)

    def control_disable(self, title, text, controlId):
        """Disables or "grays-out" a control.
        When using a control name in the Control functions, you need to add a number to the end of the name to indicate
        which control. For example, if there two controls listed called "MDIClient", you would refer to these as
        "MDIClient1" and "MDIClient2".
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: 1 Success 0 Failure
        :rtype: int
        """
        return self._aux3.ControlDisable(title, text, controlId)

    def control_enable(self, title, text, controlId):
        """Enables a "grayed-out" control.
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: 1 Success 0 Failure
        :rtype: int
        """
        return self._aux3.ControlEnable(title, text, controlId)

    def control_focus(self, title, text, controlId):
        """Sets input focus to a given control on a window.
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: 1 Success 0 Failure
        :rtype: int
        """
        return self._aux3.ControlFocus(title, text, controlId)

    def control_get_focus(self, title, text=""):
        """Returns the ControlRef# of the control that has keyboard focus within a specified window.
        :param title:
        :type title: str or int
        :param text:
        :type text: str
        :return: Success ControlRef# of the control   Failure  a blank string and sets error to 1 if window is not found
        :rtype: unicode
        """
        return self._aux3.ControlGetFocus(title, text)

    def control_get_handle(self, title, text, controlId):
        """Retrieves the internal handle of a control.
        :param title: The title of the window to read.
        :type title: str or int
        :param text: The text of the window to read.
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: unicode
        :return: Success Returns a string containing the control handle value.
                 Returns "" (blank string) and sets oAutoIt.error to 1 if no window matches the criteria.
        """
        return self._aux3.ControlGetHandle(title, text, controlId)

    def control_get_pos_height(self, title, text, controlId):
        """Retrieves the position and size of a control relative to it's window.
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: Returns the height of the control.  Failure sets error to 1.
        :rtype: int
        """
        return self._aux3.ControlGetPosHeight(title, text, controlId)

    def control_get_pos_width(self, title, text, controlId):
        """Retrieves the position and size of a control relative to it's window.
        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: Returns the width of the control.  Failure sets error to 1.
        :rtype: int
        """
        return self._aux3.ControlGetPosWidth(title, text, controlId)

    def control_get_pos_x(self, title, text, controlId):
        return self._aux3.ControlGetPosX(title, text, controlId)

    def control_get_pos_y(self, title, text, controlId):
        return self._aux3.ControlGetPosY(title, text, controlId)

    def control_get_pos(self, title, text, controlId):
        """
        :param title:
        :param text:
        :param controlId:
        :return: x,y position of control as tuple
        :rtype: tuple
        """
        return (self.control_get_pos_x(title, text, controlId), self.control_get_pos_y(title, text, controlId))

    def control_get_size(self, title, text, controlId):
        """
        :param title:
        :param text:
        :param controlId:
        :return: size of control as tuple
        :rtype: tuple
        """
        return (self.control_get_pos_width(title, text, controlId), self.control_get_pos_height(title, text, controlId))

    def control_get_text(self, title, text, controlId):
        """Retrieves text from a control.

        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: Returns the text from a control.  Failure sets error to 1 and
        :rtype: unicode
        """
        return self._aux3.ControlGetText(title, text, controlId)

    def control_hide(self, title, text, controlId):
        """Hides a control.

        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :return: Success 1  Failure 0 window/control is not found
        :rtype: int
        """
        return self._aux3.ControlHide(title, text, controlId)

    def control_list_view(self, title, text, controlId, command, option1="", option2=""):
        """Sends a command to a ListView32 control.

        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :param command: The command to send to the control
        :param option1: Additional parameter required by some commands; use "" if parameter is not required.
        :type option1: str
        :param option2: Additional parameter required by some commands; use "" if parameter is not required.
        :type option2: str
        :return:
        """
        return self._aux3.ControlListView(title, text, controlId, command, option1, option2)

    def control_move(self, title, text, controlId, x, y, width=LOWEST_INT, height=LOWEST_INT):
        return self._aux3.ControlListView(title, text, controlId, x, y, width, height)

    def control_send(self, title, text, controlId, string, flag=0):
        """Sends a string of characters to a control.
        Note, this function cannot send all the characters that the usual Send function can (notably ALT keys) but it
        can send most of them--even to non-active or hidden windows!
        ControlSend can be quite useful to send capital letters without messing up the state of "Shift."

        :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access.
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :param string: String of characters to send to the control.
        :param flag: Changes how "keys" is processed: flag = 0 (default), Text contains special characters like + to
                     indicate SHIFT and {LEFT} to indicate left arrow. flag = 1, keys are sent raw.
        :return: Success 1  Failure 0
        :rtype: int
        """
        return self._aux3.ControlSend(title, text, controlId, string, flag)

    def control_set_text(self, title, text, controlId, newText):
        """Sets text of a control.

       :param title: The title of the window to access.
        :type title: str or int
        :param text: The text of the window to access.
        :type text: str
        :param controlId: The control to interact with.
        :type controlId: int
        :param newText: The new text to be set into the control.
        :type newText: str
        :return: Success 1  Failure 0
        :rtype: int
        """
        return self._aux3.ControlSetText(title, text, controlId, newText)

    def control_show(self, title, text, controlId):
        return self._aux3.ControlShow(title, text, controlId)

    def control_tree_view(self, title, text, controlId, command, option1="", option2=""):
        return self._aux3.ControlTreeView(title, text, controlId, command, option1, option2)

    def drive_map_add(self, device, remoteShare, flags=0, user="", password=""):
        """Maps a network drive.

        :param device: The device to map, for example "O:" or "LPT1:". If you pass a blank string for this parameter a
                       connection is made but not mapped to a specific drive. If you specify "*" an unused drive letter
                       will be automatically selected.
        :param remoteShare: The remote share to connect to in the form "\\server\share".
        :param flags: A combination of the following: 0 = default, 1 = Persistant mapping,
                      8 = Show authentication dialog if required
        :param user:  The username to use to connect. In the form "username" or "domain\username".
        :param password: Optional: The password to use to connect.
        :return:
        """
        return self._aux3.DriveMapAdd(device, remoteShare, flags, user, password)

    def drive_map_del(self, device):
        """Disconnects a network drive

        :param device: The device to disconnect, e.g. "O:" or "LPT1:".
        :type device: str
        :return: Success 1  Failure 0
        :rtype: int
        """
        return self._aux3.DriveMapDel(device)

    def drive_map_get(self, device):
        """Retreives the details of a mapped drive.

        :param device: The device (drive or printer) letter to query. Eg. "O:" or "LPT1:"
        :return: Success: Returns details of the mapping, e.g. \\server\share
                 Failure: Returns a blank string "" and sets oAutoIt.error to 1.
        :rtype: str
        """
        return self._aux3.DriveMapGet(device)

    def ini_delete(self, filename, section, key=""):
        return self._aux3.IniDelete(filename, section, key)

    def ini_read(self, filename, section, key, default):
        return self._aux3.IniRead(filename, section, key, default)

    def ini_write(self, filename, section, key, value):
        return self._aux3.IniWrite(filename, section, key, value)

    def is_admin(self):
        return self._aux3.IsAdmin()

    def mouse_click(self, button, x=LOWEST_INT, y=LOWEST_INT, clicks=1, speed=10):
        return self._aux3.MouseClick(button, x, y, clicks, speed)

    def mouse_click_drag(self, button, x1, y1, x2, y2, speed=10):
        return self._aux3.MouseClickDrag(button, x1, y1, x2, y2, speed)

    def mouse_down(self, button):
        """Perform a mouse down event at the current mouse position.
        Use responsibly: For every MouseDown there should eventually be a corresponding MouseUp event.

        :param button: The button to click: "left", "right", "middle", "main", "menu", "primary", "secondary".
        :return: None
        :rtype: None
        """
        return self._aux3.MouseDown(button)

    def mouse_get_cursor(self):
        """Returns a cursor ID Number of the current Mouse Cursor.

        :return: Returns a cursor ID Number:
                 0 = UNKNOWN (this includes pointing and grabbing hand icons)
                 1 = APPSTARTING
                 2 = ARROW
                 3 = CROSS
                 4 = HELP
                 5 = IBEAM
                 6 = ICON
                 7 = NO
                 8 = SIZE
                 9 = SIZEALL
                 10 = SIZENESW
                 11 = SIZENS
                 12 = SIZENWSE
                 13 = SIZEWE
                 14 = UPARROW
                 15 = WAIT
        :rtype: int
        """
        return self._aux3.MouseGetCursor()

    def mouse_get_pos_x(self):
        """Retrieves the current X position of the mouse cursor.
        See MouseCoordMode for relative/absolute position settings. If relative positioning, numbers may be negative.

        :return: Returns the current X position of the mouse cursor.
        """
        return self._aux3.MouseGetPosX()

    def mouse_get_pos_y(self):
        return self._aux3.MouseGetPosY()

    def mouse_get_pos(self):
        """Retrieves the current position of the mouse cursor.

        :return: Returns the current position of the mouse cursor.
        :rtype: tuple
        """
        return (self.mouse_get_pos_x(), self.mouse_get_pos_y())

    def mouse_move(self, x, y, speed=10):
        """Moves the mouse pointer.

        :param x: The screen x coordinate to move the mouse to.
        :param y: The screen y coordinate to move the mouse to.
        :param speed: the speed to move the mouse in the range 1 (fastest) to 100 (slowest). A speed of 0 will move the
                      mouse instantly. Default speed is 10.
        :return:
        """
        return self._aux3.MouseMove(x, y, speed)

    def mouse_up(self, button):
        """Perform a mouse up event at the current mouse position.
        Use responsibly: For every MouseDown there should eventually be a corresponding MouseUp event.

        :param button: The button to click: "left", "right", "middle", "main", "menu", "primary", "secondary".
        :return: None
        :rtype: None
        """
        return self._aux3.MouseUp(button)

    def mouse_wheel(self, direction, clicks=1):
        """Moves the mouse wheel up or down. NT/2000/XP ONLY.

        :param direction: "up" or "down"
        :type direction: str
        :param clicks: Optional: The number of times to move the wheel. Default is 1.
        :type clicks: int
        :return: None
        :rtype: None
        """
        return self._aux3.MouseWheel(direction, clicks)

    def pixel_checksum(self, left, top, right, bottom, step=1):
        """Generates a checksum for a region of pixels.
        Performing a checksum of a region is very time consuming, so use the smallest region you are able to reduce CPU
        load. On some machines a checksum of the whole screen could take many seconds!
        A checksum only allows you to see if "something" has changed in a region - it does not tell you exactly what has
        changed. When using a step value greater than 1 you must bear in mind that the checksumming becomes less
        reliable for small changes as not every pixel is checked.

        :param left: left coordinate of rectangle.
        :param top: top coordinate of rectangle.
        :param right: right coordinate of rectangle.
        :param bottom: bottom coordinate of rectangle.
        :param step: Optional: Instead of checksumming each pixel use a value larger than 1 to skip pixels (for speed).
                     E.g. A value of 2 will only check every other pixel. Default is 1.
        :return: Returns the checksum value of the region.
        :rtype: float
        """
        return self._aux3.PixelChecksum(left, top, right, bottom, step)

    def pixel_get_color(self, x, y):
        """Returns a pixel color according to x,y pixel coordinates.

        :param x: x coordinate of pixel.
        :param y: x coordinate of pixel.
        :return: Returns decimal value of pixel's color. Failure Returns -1 if invalid coordinates.
        :rtype: int
        """
        return self._aux3.PixelGetColor(x, y)

    def pixel_search(self, left, top, right, bottom, colour, shadeVariation=0, step=1):
        """Searches a rectangle of pixels for the pixel color provided.
        The search is performed top-to-bottom, left-to-right, and the first match is returned.
        Performing a search of a region can be time consuming, so use the smallest region you are able to reduce CPU load.

        :param left: left coordinate of rectangle.
        :type left: int
        :param top: top coordinate of rectangle.
        :type top: int
        :param right: right coordinate of rectangle.
        :type right: int
        :param bottom: bottom coordinate of rectangle.
        :type bottom: int
        :param colour: Colour value of pixel to find (in decimal or hex).
        :type colour: int
        :param shadeVariation: Optional: A number between 0 and 255 to indicate the allowed number of shades of
                               variation of the red, green, and blue components of the colour.
                               Default is 0 (exact match).
        :type shadeVariation: int
        :param step: Optional: Instead of searching each pixel use a value larger than 1 to skip pixels (for speed).
                     E.g. A value of 2 will only check every other pixel. Default is 1.
        :type step: int
        :return: Success: Returns a 2 element array containing the pixel's coordinates.
                 Failure: Sets oAutoIt.error to 1 if color is not found.
        :rtype: tuple
        """
        return self._aux3.PixelSearch(left, top, right, bottom, colour, shadeVariation, step)

    def process_close(self, process):
        """Terminates a named process.
        Process names are executables without the full path, e.g., "notepad.exe" or "winword.exe"
        If multiple processes have the same name, the one with the highest PID is terminated--regardless of how
        recently the process was spawned. PID is the unique number which identifies a Process. A PID can be obtained
        through the ProcessExists or Run commands. In order to work under Windows NT 4.0, ProcessClose requires the
        file PSAPI.DLL (included in the AutoIt installation directory). The process is polled approximately every 250
        milliseconds.

        :param process: The title or PID of the process to terminate
        :type process: str or int
        :return: None. (Returns 1 regardless of success/failure.)
        :rtype: int
        """
        return self._aux3.ProcessClose(process)

    def process_exists(self, process):
        """Checks to see if a specified process exists.
        Process names are executables without the full path, e.g., "notepad.exe" or "winword.exe"
        PID is the unique number which identifies a Process.
        In order to work under Windows NT 4.0, ProcessExists requires the file PSAPI.DLL (included in the AutoIt installation directory).
        The process is polled approximately every 250 milliseconds.

        :param process: The name or PID of the process to check.
        :type process: str or int
        :return: Success: Returns the PID of the process.
                 Failure: Returns 0 if process does not exist.
        :rtype: int
        """
        return self._aux3.ProcessExists(process)

    def process_set_priority(self, process, priority):
        """Changes the priority of a process
        Above Normal and Below Normal priority classes are not supported on Windows 95/98/ME. If you try to use them on
        those platforms, the function will fail and oAutoIt.error will be set to 2.

        :param process: The name or PID of the process to check.
        :type process: str or int
        :param priority: A flag which determines what priority to set
                         0 - Idle/Low
                         1 - Below Normal (Not supported on Windows 95/98/ME)
                         2 - Normal
                         3 - Above Normal (Not supported on Windows 95/98/ME)
                         4 - High
                         5 - Realtime (Use with caution, may make the system unstable)
        :type priority: int
        :return: Success: Returns 1.
                 Failure: Returns 0 and sets oAutoIt.error to 1. May set oAutoIt.error to 2 if attempting to use an
                          unsupported priority class.
        :rtype: int
        """
        return self._aux3.SetPriority(process, priority)

    def process_wait(self, process, timeout=0):
        """Pauses script execution until a given process exists.

        :param process: The name of the process to check.
        :type process: str or int
        :param timeout: Optional: Specifies how long to wait (default is to wait indefinitely).
        :type timeout: int
        :return: Success: Returns 1.
                 Failure: Returns 0 if the wait timed out.
        :rtype: int
        """
        return self._aux3.ProcessWait(process, timeout)

    def process_wait_close(self, process, timeout=0):
        """Pauses script execution until a given process does not exist.

        :param process: The name or PID of the process to check.
        :type process: str or int
        :param timeout: Optional: Specifies how long to wait (default is to wait indefinitely).
        :type timeout: int
        :return: Success: Returns 1.
                 Failure: Returns 0 if wait timed out.
        :rtype: int
        """
        return self._aux3.ProcessWaitClose(process, timeout)

    def reg_delete_key(self, keyName):
        """Deletes a key from the registry.
        A registry key must start with "HKEY_LOCAL_MACHINE" ("HKLM") or "HKEY_USERS" ("HKU") or
        "HKEY_CURRENT_USER" ("HKCU") or "HKEY_CLASSES_ROOT" ("HKCR") or "HKEY_CURRENT_CONFIG" ("HKCC").
        Deleting from the registry is potentially dangerous--please exercise caution!
        It is possible to access remote registries by using a keyname in the form "\\computername\keyname".
        To use this feature you must have the correct access rights on NT/2000/XP/2003, or if you are using a 9x based
        OS the remote PC must have the remote regsitry service installed first (See Microsoft Knowledge Base Article
        - 141460).

        :param keyName: The registry key to write to.
        :type keyName: str
        :return: Success: Returns 1.
                 Special: Returns 0 if the key does not exist.
                 Failure: Returns 2 if error deleting key.
        :rtype: int
        """
        return self._aux3.RegDeleteKey(keyName)

    def reg_delete_val(self, keyName, valueName):
        """Deletes a value from the registry.
        A registry key must start with "HKEY_LOCAL_MACHINE" ("HKLM") or "HKEY_USERS" ("HKU") or "HKEY_CURRENT_USER"
        ("HKCU") or "HKEY_CLASSES_ROOT" ("HKCR") or "HKEY_CURRENT_CONFIG" ("HKCC").
        To access the (Default) value use "" (a blank string) for the valuename.
        Deleting from the registry is potentially dangerous--please exercise caution!

        :param keyName: The registry key to write to.
        :type keyName: str
        :param valueName: The value name to delete.
        :type valueName: str
        :return: Success: Returns 1.
                 Special: Returns 0 if the key/value does not exist.
                 Failure: Returns 2 if error deleting key/value.
        :rtype: int
        """
        return self._aux3.RegDeleteVal(keyName, valueName)

    def reg_enum_key(self, keyName, instance):
        """Reads the name of a subkey according to it's instance.
        A registry key must start with "HKEY_LOCAL_MACHINE" ("HKLM") or "HKEY_USERS" ("HKU") or "HKEY_CURRENT_USER"
        ("HKCU") or "HKEY_CLASSES_ROOT" ("HKCR") or "HKEY_CURRENT_CONFIG" ("HKCC").

        :param keyName: The registry key to read.
        :type keyName: str
        :param instance: The 1-based key instance to retrieve
        :type instance: int
        :return: Success: Returns the requested subkey name..
                 Failure: Returns "" and sets the @error flag:
                          1 if unable to open requested key
                          -1 if unable to retrieve requested subkey (key instance out of range)
        :rtype: str
        """
        return self._aux3.RegEnumKey(keyName, instance)

    def reg_enum_val(self, keyName, instance):
        """Reads the name of a value according to it's instance.

        :param keyName: The registry key to read.
        :type keyName: str
        :param instance: The 1-based value instance to retrieve.
        :type instance: int
        :return: Success: Returns the requested value name.
                 Failure: Returns "" and sets the @error flag:
                          1 if unable to open requested key
                          -1 if unable to retrieve requested value name (value instance out of range)
        :rtype: str
        """
        return self._aux3.RegEnumVal(keyName, instance)

    def reg_read(self, keyName, valueName):
        """Reads a value from the registry.
        A registry key must start with "HKEY_LOCAL_MACHINE" ("HKLM") or "HKEY_USERS" ("HKU") or "HKEY_CURRENT_USER" ("HKCU") or "HKEY_CLASSES_ROOT" ("HKCR") or "HKEY_CURRENT_CONFIG" ("HKCC").
        AutoIt supports registry keys of type REG_BINARY, REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ, and REG_DWORD.
        To access the (Default) value use "" (a blank string) for the valuename.
        When reading a REG_BINARY key the result is a string of hex characters, e.g. the REG_BINARY value of 01,a9,ff,77 will be read as the string "01A9FF77".
        When reading a REG_MULTI_SZ key the multiple entries are seperated by a linefeed character.

        :param keyName: The registry key to read.
        :type keyName: str
        :param valueName: The value to read.
        :type valueName: str
        :return: Success: Returns the requested registry value value.
                 Failure: Returns numeric 1 and sets the oAutoIt.error flag:
                          1 if unable to open requested key
                          -1 if unable to open requested value
                          -2 if value type not supported
        :rtype: unicode or int
        """
        return self._aux3.RegRead(keyName, valueName)

    def reg_write(self, keyName, valueName, type, value):
        """Creates a key or value in the registry.
        A registry key must start with "HKEY_LOCAL_MACHINE" ("HKLM") or "HKEY_USERS" ("HKU") or "HKEY_CURRENT_USER"
        ("HKCU") or "HKEY_CLASSES_ROOT" ("HKCR") or "HKEY_CURRENT_CONFIG" ("HKCC").
        AutoIt supports registry keys of type REG_BINARY, REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ, and REG_DWORD.
        To access the (Default) value use "" (a blank string) for the valuename.
        When writing a REG_BINARY key use a string of hex characters, e.g. the REG_BINARY value of 01,a9,ff,77
        can be written by using the string "01A9FF77".
        When writing a REG_MULTI_SZ key you must separate each value with @LF. The value must NOT end with @LF and
        no "blank" entries are allowed (see example).

        :param keyName: The registry key to write to. If no other parameters are specified.
        :param valueName: The valuename to write to.
        :param type: Type of key to write: "REG_SZ", "REG_MULTI_SZ", "REG_EXPAND_SZ", "REG_DWORD", or "REG_BINARY".
        :param value: The value to write.
        :return: Success: Returns 1.
                 Failure: Returns 0 if error writing registry key or value.
        :rtype: int
        """
        return self._aux3.RegWrite(keyName, valueName, type, value)

    def run(self, filename, workingDir="", flag=1):
        """Runs an external program.
        After running the requested program the script continues. To pause execution of the script until the spawned
        program has finished use the RunWait function instead. The error property is set to 1 as an indication of
        failure.

        :param filename: The name of the executable (EXE, BAT, COM, or PIF) to run.
        :type filename: str
        :param workingDir: The working directory.
        :type workingDir: str
        :param flag: The "show" flag of the executed program.
        :type flag: int
        :return: Success: The PID of the process that was launched.
                 Failure: see Remarks.
        :rtype: int
        """
        return self._aux3.Run(filename, workingDir, flag)

    def run_as_set(self, user, domain, password, options=1):
        """Initialise a set of user credentials to use during Run and RunWait operations. 2000/XP or later ONLY.
        This function allows subsequent Run and RunWait functions to run as a different user (e.g. Administrator).
        The function only works on the 2000/XP (or later) platforms. NT4 users should install and use the SU command
        from the NT Resource Kit.
        The "Secondary Logon service" or "RunAs service" must not be disabled if you want this function to work.
        To unset the RunAs details, use the function with no parameters: RunAsSet().

        :param user: The user name to use.
        :type user: str
        :param domain: The domain name to use.
        :type domain: str
        :param password: The password to use.
        :type password: str
        :param options: Optional: 0 = do not load the user profile, 1 = (default) load the user profile,
                        2 = use for net credentials only
        :type options: int
        :return: Returns 0 if the operating system does not support this function.
                 Otherwise returns 1 --regardless of success. (If the login information was invalid,
                 subsequent Run/RunWait commands will fail....)
        :rtype: int
        """
        return self._aux3.RunAsSet(user, domain, password, options)

    def run_wait(self, filename, workingDir="", flag=1):
        """Runs an external program and pauses script execution until the program finishes.
        After running the requested program the script pauses until the program terminates.
        To run a program and then immediately continue script execution use the Run function instead.
        Some programs will appear to return immediately even though they are still running; these programs spawn
        another process - you may be able to use the ProcessWaitClose function to handle these cases.
        The error property is set to 1 as an indication of failure.

        :param filename: The name of the executable (EXE, BAT, COM, PIF) to run.
        :param workingDir: Optional: The working directory.
        :param flag: Optional: The "show" flag of the executed program:
                     SW_HIDE = Hidden window
                     SW_MINIMIZE = Minimized window
                     SW_MAXIMIZE = Maximized window
        :return: Success: Returns the exit code of the program that was run.
                 Failure: see Remarks.
        :rtype: int
        """
        return self._aux3.RunWait(filename, workingDir, flag)

    def send(self, keys, flag=0):
        """Sends simulated keystrokes to the active window.
        See AutoItX help file for keys format.

        :param keys: The sequence of keys to send.
        :param flag: Optional: Changes how "keys" is processed:
                     flag = 0 (default), Text contains special characters like + and ! to indicate SHIFT and ALT key presses.
                     flag = 1, keys are sent raw.
        :return: None
        :rtype: None
        """
        return self._aux3.Send(keys, flag)

    def shutdown(self, code):
        """Shuts down the system.
        The shutdown code is a combination of the following values:
        0 = Logoff
        1 = Shutdown
        2 = Reboot
        4 = Force
        8 = Power down

        :param code: A combination of shutdown codes. See "remarks".
        :return: Success: Returns 1.
                 Failure: Returns 0.
        :rtype: int
        """
        return self._aux3.Shutdown(code)

    def sleep(self, delay):
        """Pause script execution.

        :param delay: Amount of time to pause (in milliseconds).
        :type delay: int
        :return: None
        :rtype: None
        """
        return self._aux3.Sleep(delay)

    def statusbar_get_text(self, title, text="", part=1):
        """Retrieves the text from a standard status bar control.
        This functions attempts to read the first standard status bar on a window
        (Microsoft common control: msctls_statusbar32). Some programs use their own status bars or special versions of
        the MS common control which StatusbarGetText cannot read. For example, StatusbarText does not work on the
        program TextPad; however, the first region of TextPad's status bar can be read using
        ControlGetText("TextPad", "", "HSStatusBar1")
        StatusbarGetText can work on windows that are minimized or even hidden.

        :param title: The title of the window to check.
        :type title: str
        :param text: Optional: The text of the window to check.
        :type text: str
        :param part: Optional: The "part" number of the status bar to read - the default is 1. 1 is the first possible
                     part and usually the one that contains the useful messages like "Ready" "Loading...", etc.
        :return: Success: Returns the text read.
                 Failure: Returns empty string and sets oAutoIt.error to 1 if no text could be read.
        :rtype: unicode
        """
        return self._aux3.StatusBarGetText(title, text, part)

    def tool_tip(self, text, x=LOWEST_INT, y=LOWEST_INT):
        """Creates a tooltip anywhere on the screen.
        If the x and y coordinates are omitted the, tip is placed near the mouse cursor.
        If the coords would cause the tooltip to run off screen, it is repositioned to visible.
        Tooltip appears until it is cleared, until script terminates, or sometimes until it is clicked upon.
        You may use a linefeed character to create multi-line tooltips.

        :param text: The text of the tooltip. (An empty string clears a displaying tooltip)
        :type text: str
        :param x: The x position of the tooltip.
        :type x: int
        :param y: The y position of the tooltip.
        :type y: int
        :return: None
        :rtype: None
        """
        return self._aux3.ToolTip(text, x, y)

    def win_activate(self, title, text=""):
        """Activates (gives focus to) a window.
        You can use the WinActive function to check if WinActivate succeeded. If multiple windows match the criteria,
        the window that was most recently active is the one activated.
        WinActivate works on minimized windows. However, a window that is "Always On Top" could still cover up a window
        you Activated.

        :param title: The title of the window to activate.
        :type title: str
        :param text: Optional: The text of the window to activate.
        :type text: str
        :return: None
        :rtype: None
        """
        return self._aux3.WinActivate(title, text)

    def win_active(self, title, text=""):
        """Checks to see if a specified window exists and is currently active.

        :param title: The title of the window to activate.
        :type title: str
        :param text: Optional: The text of the window to activate.
        :type text: str
        :return: Success: Returns 1.
                 Failure: Returns 0 if window is not active.
        :rtype: int
        """
        return self._aux3.WinActive(title, text)

    def win_close(self, title, text=""):
        """Closes a window.
        This function sends a close message to a window, the result depends on the window
        (it may ask to save data, etc.). To force a window to close, use the WinKill function.
        If multiple windows match the criteria, the window that was most recently active is closed.

        :param title: The title of the window to close.
        :param text: Optional: The text of the window to close.
        :return: None
        :rtype: None
        """
        return self._aux3.WinClose(title, text)

    def win_exists(self, title, text):
        """Checks to see if a specified window exists.
        WinExist will return 1 even if a window is hidden.

        :param title: The title of the window to check.
        :type title: str
        :param text: Optional: The text of the window to check.
        :type text: str
        :return: Returns 1 if the window exists, otherwise returns 0.
        """
        return self._aux3.WinExists(title, text)

    def win_get_caret_pos_x(self):
        """Returns the coordinates of the caret in the foreground window
        WinGetCaretPos might not return accurate values for Multiple Document Interface (MDI) applications
        if absolute CaretCoordMode is used.
        See example for a workaround.
        Note: Some applications report static coordinates regardless of caret position!

        :return: Success: Returns the X coordinate of the caret.
                 Failure: Sets oAutoIt.error to 1.
        :rtype: int
        """
        return self._aux3.WinGetCaretPosX()

    def win_get_caret_pos_y(self):
        """Returns the coordinates of the caret in the foreground window

        :return: Success: Returns the Y coordinate of the caret.
                 Failure: Sets oAutoIt.error to 1.
        :rtype: int
        """
        return self._aux3.WinGetCaretPosY()
