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

    library_name = ['pygetwindow', 'psutil', 'shutil', 'GitPython']

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

    process_names = ['quartus', 'calc'] 

    # Find running processes matching the process name
    for process_name in process_names:
        for proc in psutil.process_iter(['name']):
            if process_name in proc.info['name'].lower():
                print(f"Found {process_name} process (PID: {proc.pid}). Closing...")
                proc.kill()

    # Check if there is any Latest directory. If not that means that it's a fresh PC
    # Need to create Latest folder and icon on the desktop. 
    
    directories = ['FPGA_design_store', 'upython-esp32', 'VGAstarter_DE10_lite', 'micropython']

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
            subprocess.Popen([location])
            break

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
