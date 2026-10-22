#!/usr/bin/python
# -*- coding: utf-8 -*-

# Install pyxls
import subprocess, sys, glob, re, argparse, importlib
from pathlib import Path
requests = None
# Later in code set with : requests = importlib.import_module("requests")
# See also am import within install_paikkain!
_TEST = False
pythonpath = Path(sys.executable).parent

batfiles = glob.glob("batsrc/*.bat")
installdir = Path(__file__).parent # The directory where the installer is
configdir = installdir / Path("config")
sourcefile = installdir / "src" / "paikkain.py"
datafiledir = installdir / Path("paikka-aineistot")

# Data files
latest_finland_file_id = r"1wlivBAY6AAVlIQAeAD7bU4o84JFPe60z"
latest_world_file_id = r"1aQi9u-oFLKEoUzPdtDhpDl_Ml0Gvu03h"
latest_legnames_id = r"1093uscp4XQJ20CiwU9agdm-Ozg80xrQg"
# Version numbers in files
latest_finland_version_id = "1Hd5hRbTG0dDkE3gZzdBhw8LE1fCKhS-x"
latest_world_version_id = r"1Ki9X5Y_PV2yccdCj9v1ewGtSwdh-_j9_"

latestdata = {
    'Paikkain-world.version.txt': latest_world_version_id, 
    'Paikkain-world.xlsx': latest_world_file_id, 
    'Paikkain-fin-sl.version.txt': latest_finland_version_id, 
    'Paikkain-fin-sl.xlsx': latest_finland_file_id, 
    'Paikkain-legnames.xlsx': latest_legnames_id, 
    'Paikkain-legnames.version.txt': latest_finland_version_id, 
    }

conf_to_datafile = {
    "paikkain_world.toml" : 'Paikkain-world.xlsx', 
    "paikkain_finland.toml" : 'Paikkain-fin-sl.xlsx', 
    "paikkain_fullnames.toml" : 'Paikkain-legnames.xlsx',
}    
    

re_datafile_line = re.compile( r'filenames = \[(.*)\]' )
re_pythonexe = re.compile( r'"\S*?\.exe"' )

def download_file_from_google_drive(file_id, destination):
    URL = "https://docs.google.com/uc?export=download&confirm=1"

    session = requests.Session()

    response = session.get(URL, params={"id": file_id}, stream=True)
    token = get_confirm_token(response)

    if token:
        params = {"id": file_id, "confirm": token}
        response = session.get(URL, params=params, stream=True)

    save_response_content(response, destination)

def get_confirm_token(response):
    for key, value in response.cookies.items():
        if key.startswith("download_warning"):
            return value
    return None

def save_response_content(response, destination):
    CHUNK_SIZE = 32768
    with open(destination, "wb") as f:
        for chunk in response.iter_content(CHUNK_SIZE):
            if chunk:  # filter out keep-alive new chunks
                f.write(chunk)

def update_datafiles(latestdata):
    global requests  
    requests = importlib.import_module("requests") # Can import this only after making sure it is present
    for filename, google_id in latestdata.items():
        # Download latest Finland data
        print(f"Trying to download latest data files.")
        targetfile = datafiledir / filename 
        if not _TEST: download_file_from_google_drive(google_id, targetfile)
        print(f"\tdownloading {filename} successfull.")

def install_packages():
    reqfile = "requirements.txt"
    print(f"Trying to install a Python modules from file {reqfile}")
    rv = subprocess.run([sys.executable, "-m", "pip", "install", "-r", reqfile])
    try:
        rv.check_returncode()
        print(f"\tInstalled {reqfile} successfull.")
    except subprocess.CalledProcessError as msg:
        print(f"\tInstall {reqfile} failed with error message. Exiting.")
        sys.exit()

def install_paikkain():
    if not datafiledir.exists():
        datafiledir.mkdir()
    install_packages( )  
    global requests  
    requests = importlib.import_module("requests") # Can import this only after making sure it is present
    update_datafiles( latestdata )
    # TODO: THESE DEPEND ON THE PATHS NOTING HAVING EMPTY SPACES!
    re_sourcefile = re.compile( r'"\S*?\.py"' )
    re_configfile = re.compile( r'"\S*?\.toml"' )
    for fn in [Path(x) for x in batfiles]:
         outfp = installdir / fn.with_suffix(".bat").name
         print("Creating BAT file from template: ", outfp)
         with fn.open("r") as f:
             d = f.read()
         # Path to src/paikkain.py
         newsource = '"' + str( sourcefile) + '"'
         newsource = newsource.replace( "\\", "/" )
         d = re_sourcefile.sub(newsource, d)
         # Path to  python itself
         newpythonpath = '"' + str(sys.executable) + '"'
         newpythonpath = newpythonpath.replace( "pythonw.exe", "python.exe" ) 
         newpythonpath = newpythonpath.replace( "\\", "/" ) 
         # Path to configfile
         m = re_configfile.search(d)
         if m:             
            oldconfpath = Path(m[0][1:-1]) # Strip ""
            newconfpath =  '"' + str( configdir / oldconfpath.name) + '"'
            newconfpath = newconfpath.replace( "\\", "/" ) 
            d = re_configfile.sub(newconfpath, d)
         d,n = re_pythonexe.subn(newpythonpath, d)
         with outfp.open("w") as outfile: outfile.write(d)

    # INSERT PROPER FILENAMES INTO CONFIG FILES
    for conf_fn, sourcefilename in conf_to_datafile.items():
        fp = configdir / Path(conf_fn)
        datafilepath = datafiledir / sourcefilename
        print(f"Inserting {datafilepath.name} into {fp}.")
        with fp.open("r") as f:
            d = f.read()
        newline = rf"filenames = ['{datafilepath}']"
        newline = newline.replace("\\","/")
        d = re_datafile_line.sub(newline , d)
        with fp.open("w") as f: f.write(d) # Write updated data
        print(f"done")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
                    prog='paikkain')
    parser.add_argument("-u", "--update_only", action="store_true",
                        help="Only update data files to latest versions.")
    args = parser.parse_args()
    if args.update_only: update_datafiles( latestdata )
    else: install_paikkain()
