#!/usr/bin/python
# -*- coding: utf-8 -*-

# Install pyxls
import subprocess, requests, sys, glob, re
from pathlib import Path

pythonpath = Path(sys.executable).parent

batfiles = glob.glob("batsrc/*.bat")
installdir = Path(__file__).parent # The directory where the installer is
setupdir = Path("config")
sourcefile = installdir / "src" / "paikkain.py"
datafiledir = installdir / Path("paikka-aineistot")

latest_finland_file_id = "1wlivBAY6AAVlIQAeAD7bU4o84JFPe60z"
latest_finland_version_id = "1Hd5hRbTG0dDkE3gZzdBhw8LE1fCKhS-x"
latest_world_file_id = "1-8ZDv5jDYx6MpoVYZ7AWg4b_XqWU32fQCkebtirJYH4"
latest_world_version_id = "1Ki9X5Y_PV2yccdCj9v1ewGtSwdh-_j9_"
latest_legnames_id = "1kooBZAoItNJreXN8wYTooVjzh7L_DNSD1140Co1pb6Y"

latestdata = {
    'Paikkain-World-latest.version.txt': latest_world_version_id, 
    'Paikkain-World-latest.xlsx': latest_world_file_id, 
    'Paikkain-Fin-sl-latest.version.txt': latest_finland_version_id, 
    'Paikkain-Fin-sl-latest.xlsx': latest_finland_file_id, 
    'Paikkain-legnames-latest.xlsx': latest_legnames_id, 
    'Paikkain-legnames-latest.version.txt': latest_finland_version_id, 
    }

conf_to_datafile = {
    "paikkain_world.toml" : 'Paikkain-World-latest.xlsx', 
    "paikkain_finland.toml" : 'Paikkain-Fin-sl-latest.xlsx', 
    "paikkain_fullnames.toml" : 'Paikkain-legnames-latest.xlsx',
}    
    
packages = ["openpyxl"]

re_datafile_line = re.compile( r'filenames = \[(.*)\]' )

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

def replace_pythonexe_in_string(datastring):
     newpython = '"' + str(sys.executable.replace("\\","/")) + '"'
     return re_python.sub(newpython, datastring)

def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

datapackages = [latest_finland_file_id]
def install_paikkain():
    for package in packages:
        print(f"Trying to install a Python module {package}")
#        install_package(package)
        print(f"\tinstall {package} successfull.")
    print(datafiledir)
    if not datafiledir.exists():
        datafiledir.mkdir()
    for filename, google_id in latestdata.items():
        # Download latest Finland data
        print(f"Trying to download latest data files.")
        targetfile = datafiledir / filename 
#        download_file_from_google_drive(google_id, targetfile)
        print(f"\tdownloading {filename} successfull.")
    # REWRITE BAT FILES TO REFER TO CORREC
    re_pythonexe = re.compile( r'"\S*?exe"' )
    re_sourcefile = re.compile( r'"\S*?py"' )
    for fn in [Path(x) for x in batfiles]:
         outfp = installdir / fn.with_suffix(".bat").name
         print("Creating BAT file from template: ", outfp)
         with fn.open("r") as f:
             d = f.read()
         newsource = '"' + str( sourcefile) + '"'
         newsource = newsource.replace( "\\", "/" )
         d = re_sourcefile.sub(newsource, d)
         newpythonpath = '"' + str(sys.executable) + '"'
         newpythonpath = newpythonpath.replace( "\\", "/" ) 
         d,n = re_pythonexe.subn(newpythonpath, d)
         with outfp.open("w") as outfile: outfile.write(d)

    # INSERT PROPER FILENAMES INTO CONFIG FILES
    for conf_fn, sourcefilename in conf_to_datafile.items():
        fp = setupdir / Path(conf_fn)
        datafilepath = datafiledir / sourcefilename
        print(f"Inserting {datafilepath.name} into {fp}.")
        with fp.open("r") as f:
            d = f.read()
        newline = rf"filenames = ['{datafilepath}']"
        newline = newline.replace("\\","/")
        d = re_datafile_line.sub(newline , d)
        with fp.open("w") as f: f.write(d) # Write updated data
        print(f"done")

install_paikkain()
