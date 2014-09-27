__author__ = 'florian.schaeffeler'
import pytest
from autoit.autoitx import AutoItX3




@pytest.fixture(scope="class")
def calcExecFilePath():
    return "calc.exe"

@pytest.fixture(scope="class")
def runCalcPid(calcExecFilePath, autoit, request):
    pid = autoit.run(calcExecFilePath)
    closeAutoIt = lambda: autoit.process_close(pid)
    request.addfinalizer(closeAutoIt)
    assert autoit.au3.WinWait("Calculator", "", 5), "Calculator GUI didn't appear"
    return pid

@pytest.fixture(scope="class")
def autoit():
    return AutoItX3()



class TestAutoItX3(object):

    def test_run(self, calcExecFilePath, autoit):
        pid = autoit.run(calcExecFilePath)
        assert autoit.error == 0
        assert pid > 0
        print autoit.process_close(pid)

    def test_control_click(self, runCalcPid, autoit):
        """Add 5 + 3 in calc.exe
        Steps to reproduce:
            1. Open calc.exe
            2. Click on Button 5
            3. Click on Button 3
            Expected Result: 8 should appear in TextBox
        """
        assert isinstance(autoit, AutoItX3)
        for controlId in [135, 93, 133, 121]:  # 5 + 3
            assert autoit.control_click("Calculator", "", controlId), "Could not click on Button %s %s" % (
               controlId, self.test_control_click.__doc__)
        assert True

    def test_control_command(self, runCalcPid, autoit):
        assert isinstance(autoit, AutoItX3)
        assert autoit.control_command("Calculator", "", 135, "IsVisible", "")

    def test_control_disable(self, runCalcPid, autoit):
        assert isinstance(autoit, AutoItX3)
        assert autoit.control_disable("Calculator", "", 135)

    def test_mouse_move_to_center(self, autoit):
        assert autoit.mouse_move(0, 0, speed=0)
        assert autoit.mouse_get_pos() == (0, 0)