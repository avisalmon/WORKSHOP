import importlib
import subprocess
import sys
import socket
import os
import win32com.client
import platform
import serial.tools.list_ports
import re
import serial
import time

library_name = ['pygetwindow', 'psutil', 'shutil', 'GitPython', 'esptool', 'adafruit-ampy']

    for library in library_name:
        try:
            # Try importing the library
            importlib.import_module(library)
            print(f"{library} is already installed.")
        except ImportError:
            # If the library is not found, install it
            print(f"{library} is not installed. Installing...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', library])

            # Import the library after installation
            try:
                importlib.import_module(library)
            except:
                pass

            print(f"{library_name} has been installed.")