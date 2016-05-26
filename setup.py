try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

setup(
    name='icalmerge',
    version='0.1dev',
    description='ical generation for mail merge',
    author='Nate Parsons',
    url='https://github.com/nsp/icalmerge',
    py_modules=['icalmerge',],
    license='GPLv3',
    long_description=open('README.md').read(),
    install_requires=['click',
                      'icalendar',
                      'openpyxl',
                      'python-dateutil',
                      'pytz'],
)
