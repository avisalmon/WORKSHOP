# This script will prepare the participants workshop on a windows computer. 
# Created by Avi salmon on a sunny weekend 2023. 

# 1. Check if its a new computer with no Latest directory and clone if needed to Lates. 
# 2. Delete the WORKSHOP directory
# 3. Create a new WORKSHOP dirctory for the participant
# 4. Open Quartus
# 5. Check for Latest updates and pull from git
# C: - Latest
#     |_ ..Directories for reference. 
#    - WORKSHOP
#     |_ ... Directories for the student

# make sure all libraries exsists in the system and import it

import importlib
import subprocess
import sys
import socket
import os
import re
from git import Repo
import configparser

library_name = ['pygetwindow', 'psutil', 'shutil', 'GitPython', 'esptool', 'adafruit-ampy', 'pyserial', 'pywin32', 'adafruit-ampy']

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
            print(f'could not import {library}')
            pass

        print(f"{library_name} has been installed.")

import win32com.client
import time
import serial
import platform
import shutil
import psutil


def find_your_esp32():
    wmi = win32com.client.GetObject("winmgmts:")
    drivers = wmi.InstancesOf("Win32_SystemDriver")

    list_of_drivers = []

    CP210 = False

    for driver in drivers:
        list_of_drivers.append(driver.description)
        if "CP210" in driver.description:
            print("Driver CP210X installed. Great!")
            CP210 = True

    if not CP210:
        driverexe = r'CP210xVCPInstaller_x64.exe'

        # Run the executable and wait until it's done
        subprocess.run(driverexe, shell=True, check=True)

    #check if ESP32 is connected
    com_number = 0
    while com_number == 0:
        ports = serial.tools.list_ports.comports()
            
        pattern = r'COM(\d+)'

        if not ports:
            print("No COM ports found.")
        else:
            #print("Available COM ports:")
            for port in ports:
                #print(f"- {port.device}: {port.description}")
                if "CP210" in port.description:
                    match = re.search(pattern, port.description)
                    if match:
                        com_number = int(match.group(1))

        if com_number:
            print(f'your ESP32 is connected to COM{com_number}')
            return com_number
        else:
            input('\nPlease connect the ESP32 device, than press any key...')
            return 0


#check for Internet connection
def check_internet_connection():
    try:
        # Attempt to connect to a reliable server (Google DNS)
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except socket.error:
        return False
    
# function Upload files to device
def upload_files_to_device(port, directory):
    # Walk through the directory and upload each file.
    for root, dirs, file_names in os.walk(directory):
        for file_name in file_names:
            file_path = os.path.join(root, file_name)
            dest_path = file_path.split(directory)[-1].lstrip('\\/')  # Get the relative path to maintain directory structure.

            # Use ampy CLI tool to upload files to the board
            command = [
                'ampy', 
                '--port', port, 
                'put', 
                file_path, 
                dest_path
            ]
            print(f"Uploading {file_path} to {dest_path}...")
            try:
                subprocess.run(command, check=True)
            except subprocess.CalledProcessError as e:
                print(f"Failed to upload {file_path}: {e}")

def main():

    # Check internet connection
    while not check_internet_connection():
        input('You are not connected to the Internet\nPlease connect to the Internet, press enter when done....')

    # if check_internet_connection():
    #     print("Internet connection is available.")
    # else:
    #     print("No internet connection. Please connect to the network")
    #     sys.exit()

    # Check if libraries are installed:
    # Checking for needed libraries. 

    # Kill all Quartus and unneeded windows

    process_names = ['quartus', 'calc', 'thonny'] 

    # Find running processes matching the process name
    for process_name in process_names:
        for proc in psutil.process_iter(['name']):
            if process_name in proc.info['name'].lower():
                print(f"Found {process_name} process (PID: {proc.pid}). Closing...")
                proc.kill()

    # Check if there is any Latest directory. If not that means that it's a fresh PC
    # Need to create Latest folder and icon on the desktop. 

    directories = ['FPGA_design_store', 'upython-esp32', 'VGAstarter_DE10_lite', 'single_button']

    directories_plus = directories + ['WORKSHOP']
    print(directories_plus)

    if os.path.exists('C:/Latest') and os.path.isdir('C:/Latest'):
        # print('This is not a new Laptop. thats great. lets run')
        # lets check if there are new repositories missing from Latest. 
        for dir in directories_plus:
            dir_full = 'C:/Latest/' + dir
            if not os.path.exists(dir_full):
                repository_url = 'https://github.com/avisalmon/' + dir
                local_directory = 'C:/Latest/' + dir
                os.mkdir(local_directory)
                Repo.clone_from(repository_url, local_directory)
                print("Clone completed successfully:  " + dir)
    else:
        os.mkdir('C:/Latest')
        for dir in directories_plus:
            repository_url = 'https://github.com/avisalmon/' + dir
            local_directory = 'C:/Latest/' + dir
            os.mkdir(local_directory)
            Repo.clone_from(repository_url, local_directory)
            print("Clone completed successfully:  " + dir)

    source_file = "C:/Latest/WORKSHOP/prep_workshop.py"

    desktop = os.path.join(os.path.expanduser("~"), "Desktop")


    # Remove the existing WORKSHOP directory. 
    directory_path = "C:/WORKSHOP/"

    if os.path.exists(directory_path) and os.path.isdir(directory_path):
        print(f"The directory '{directory_path}' exists.")
        shutil.rmtree(directory_path, ignore_errors=True)
        print(f"The directory '{directory_path}' and its contents have been deleted.")

    os.mkdir('C:/WORKSHOP')

    # Copy from Lates to a new WORKSHOP directory for the student

    destination_directory = "C:/WORKSHOP"

    def ignore_git_files(directory, filenames):
        return [name for name in filenames if name == '.git']

    for directory in directories:
        destination_directory = f'C:/WORKSHOP/{directory}/'
        src = 'C:/Latest/' + directory + '/'
        shutil.copytree(src, destination_directory, ignore=ignore_git_files)
        # The copy is ignoring the .git directory
    print('copied all fresh files for the workshop in C:\\WORKSHOP')

    #Let's open a Quartus window


    quartus_executable_options = ['C:/intelFPGA_lite/17.0/quartus/bin64/quartus.exe',
                                'C:/intelFPGA_lite/17.1/quartus/bin64/quartus.exe',
                                'C:/intelFPGA_lite/21.1/quartus/bin64/quartus.exe'
                                ]

    for location in quartus_executable_options:
        if os.path.exists(location):
            #subprocess.Popen([location]) # removed for now until FPGA training will come back. 
            print('now usualy would opened Quartus')
            break

    try:
        #get Thonny path. 
        home_directory = os.path.expanduser("~")
        print(f'Home dir: {home_directory}')
        thonny_path = os.path.join(home_directory, "AppData\Local\Programs\Thonny\\thonny.exe")
        print(f'Thonny location: {thonny_path}')

        #setup the comport to Thonny before opening it:
        com_num = 0 
        while not com_num:
            com_num = find_your_esp32()
        
        # lets burn FW.
        port = 'COM' + str(com_num)
        def check_micropython(port):
            with serial.Serial(port, 115200, timeout=1) as ser:
                ser.write(b'\r\n')
                time.sleep(2)
                response = ser.read_all().decode()
                #print(response)
                return ">>>" in response
        
        if check_micropython(port):
            print(f"MicroPython detected on {port}")
        else:
            print(f"No MicroPython prompt detected on {port}")

            # Burn MicroPython 
            # Try to find 'esptool.py'
            esptool_py = shutil.which('esptool.py')

            # If 'esptool.py' was not found, try to find 'esptool'
            esptool = shutil.which('esptool') if esptool_py is None else esptool_py 

            subprocess.run('cls', shell=True, check=True)
            cmd = f'{esptool} --chip esp32 --port {port} --baud 460800 write_flash -z 0x1000 ESP32_GENERIC-20230426-v1.20.0.bin'
            input(f'*********\n\n\nGet Ready! We will now burn the ESP32 to have Micropython\n\nhold the boot button right to the USB port and press enter\n\n*****************')
            subprocess.run(cmd, shell=True, check=True)
            
            upload_files_to_device(port, r'C:\WORKSHOP\single_button\Upload_these_to_device')

        # process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # stdout, stderr = process.communicate()

        # if process.returncode == 0:
        #     print("Flashing successful!")
        #     print(stdout.decode())
        # else:
        #     print("Flashing failed!")
        #     print(stderr.decode())

        # Fix Thonny before lounching it: 
        config_path = os.path.join(os.getenv('APPDATA'), 'Thonny', 'configuration.ini')
        file_location = r'C:\WORKSHOP\single_button'
        # port is now 'COM#' 
        config = configparser.ConfigParser()
        config.read(config_path)
        config['ESP32']['port'] = port
        #config['run'] = {'working_directory': file_location}

        with open(config_path, 'w') as configfile:
            config.write(configfile)
        

        subprocess.Popen(thonny_path)
        #print(thonny_path)
        print('tried to open it')
        print(f'The ESP32 is on COM{find_your_esp32()}')
    except subprocess.CalledProcessError:
        print('Didnt find Thonny')

    # pull new github updates
    print('Pulling:' + str(directories_plus))

    from git import Repo
    import git

    for directory in directories_plus:
        repository_dir = 'C:/Latest/' + directory
        os.chdir(repository_dir)
        print(os.getcwd())
        # Open the repository
        repo = git.Repo(repository_dir)
        # Perform the git pull operation
        repo.git.pull()



if __name__ == "__main__":
    main()

