__author__ = 'florian.schaeffeler'
import sys
from setuptools import setup
from setuptools.command.test import test as TestCommand


class PyTest(TestCommand):
    user_options = [('pytest-args=', 'a', "Arguments to pass to py.test")]

    def initialize_options(self):
        TestCommand.initialize_options(self)
        self.pytest_args = []

    def finalize_options(self):
        TestCommand.finalize_options(self)
        self.test_args = []
        self.test_suite = True

    def run_tests(self):
        #import here, cause outside the eggs aren't loaded
        import pytest
        errno = pytest.main(self.pytest_args)
        sys.exit(errno)

if __name__ == "__main__":
    setup(
        name='pyautoit', version='0.1',
        description='Python wrapper for AutoIt COM interface',
        author='Florian Schaeffeler',
        packages=[
            'autoit'],
        package_dir={'autoit': 'autoit'},
        package_data={'avira': ['AutoItX3.dll']},
        tests_require=['pytest'],
        cmdclass={'test': PyTest}
    )