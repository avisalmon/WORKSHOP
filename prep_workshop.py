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
        else:
            input('\nPlease connect the ESP32 device, than press any key...')

        return com_number

def main():

    #check for Internet connection
    def check_internet_connection():
        try:
            # Attempt to connect to a reliable server (Google DNS)
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            return True
        except socket.error:
            return False

    # Check internet connection
    if check_internet_connection():
        print("Internet connection is available.")
    else:
        print("No internet connection. Please connect to the network")
        sys.exit()

    # Check if libraries are installed:
    # Checking for needed libraries. 

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

    # Kill all Quartus and unneeded windows
    import psutil

    process_names = ['quartus', 'calc', 'thonny'] 

    # Find running processes matching the process name
    for process_name in process_names:
        for proc in psutil.process_iter(['name']):
            if process_name in proc.info['name'].lower():
                print(f"Found {process_name} process (PID: {proc.pid}). Closing...")
                proc.kill()

    # Check if there is any Latest directory. If not that means that it's a fresh PC
    # Need to create Latest folder and icon on the desktop. 

    directories = ['FPGA_design_store', 'upython-esp32', 'VGAstarter_DE10_lite', 'micropython', 'fab_exp', 'single_button']

    directories_plus = directories + ['WORKSHOP']
    print(directories_plus)

    from git import Repo
    if os.path.exists('C:/Latest') and os.path.isdir('C:/Latest'):
        print('This is not a new Laptop. thats great. lets run')
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

    import shutil

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
            #subprocess.Popen([location])
            print('now opened Quartus')
            break

    try:
        #get Thonny path. 
        home_directory = os.path.expanduser("~")
        print(f'Home dir: {home_directory}')
        thonny_path = os.path.join(home_directory, "AppData\Local\Programs\Thonny\\thonny.exe")
        print(f'Thonny location: {thonny_path}')

        #setup the comport to Thonny before opening it: 
        com_num = find_your_esp32()
        
        # lets burn FW.
        port = 'COM' + str(com_num)
        def check_micropython(port):
            with serial.Serial(port, 115200, timeout=1) as ser:
                ser.write(b'\r\n')
                time.sleep(2)
                response = ser.read_all().decode()
                print(response)
                return ">>>" in response
        
        if check_micropython(port):
            print(f"MicroPython detected on {port}")
        else:
            print(f"No MicroPython prompt detected on {port}")

            # Burn MicroPython 
            cmd = f'esptool\esptool.py --chip esp32 --port {port} --baud 460800 write_flash -z 0x1000 ESP32_GENERIC-20230426-v1.20.0.bin'
            input(f'*********\n\n\nGet Ready! We will now burn the ESP32 to have Micropython\n\nhold the boot button right to the USB port and press enter\n\n*****************')
            subprocess.run(cmd, shell=True, check=True)

        # process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # stdout, stderr = process.communicate()

        # if process.returncode == 0:
        #     print("Flashing successful!")
        #     print(stdout.decode())
        # else:
        #     print("Flashing failed!")
        #     print(stderr.decode())

        subprocess.Popen(thonny_path)
        print(thonny_path)
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

