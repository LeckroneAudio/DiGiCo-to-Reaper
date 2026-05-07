"""
Setup script for DiGiCo to Reaper Converter
"""
from setuptools import setup

APP = ['digico_to_reaper.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['rumps'],
    'plist': {
        'CFBundleName': 'DiGiCo to Reaper',
        'CFBundleDisplayName': 'DiGiCo to Reaper',
        'CFBundleIdentifier': 'com.digico.reaper',
        'CFBundleVersion': '2.0.0',
        'CFBundleShortVersionString': '2.0.0',
        'LSUIElement': False,  # Show in Dock
        'NSHighResolutionCapable': True,
    },
}

setup(
    name='DiGiCo to Reaper',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
