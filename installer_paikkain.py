# Install pyxls
import subprocess, requests, sys, glob, re
from pathlib import Path

pythonpath = Path(sys.executable).parent
batfiles = glob.glob("batsrc/*.bat")
installdir = Path(__file__).parent # The directory where the installer is

latest_finland_file_id = "1wlivBAY6AAVlIQAeAD7bU4o84JFPe60z"
latest_finland_version_id = "1Hd5hRbTG0dDkE3gZzdBhw8LE1fCKhS-x"
latest_world_file_id = "1-8ZDv5jDYx6MpoVYZ7AWg4b_XqWU32fQCkebtirJYH4"
latest_world_version_id = "1Ki9X5Y_PV2yccdCj9v1ewGtSwdh-_j9_"
latest_legnames_id = "1kooBZAoItNJreXN8wYTooVjzh7L_DNSD1140Co1pb6Y"

latestdata = {
    'Paikkain_World_sl-latest.version.txt': latest_world_version_id, 
    'Paikkain_World_sl-latest.xlsx': latest_world_file_id, 
    'Paikkain_Fin_sl-latest.version.txt': latest_finland_version_id, 
    'Paikkain_Fin_sl-latest.xlsx': latest_finland_file_id, 
    'Paikkain_leg_names-latest.xlsx': latest_legnames_id, 
    }
datafiledir = installdir / Path("paikka-aineistot_TEST")
packages = ["openpyxl"]

re_python = re.compile(r'".*?exe"')

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
        install_package(package)
        print(f"\tinstall {package} successfull.")
    print(datafiledir)
    if not datafiledir.exists():
        datafiledir.mkdir()
    for filename, google_id in latestdata.items():
        # Download latest Finland data
        print(f"Trying to download latest data files.")
        targetfile = datafiledir / filename 
        download_file_from_google_drive(google_id, targetfile)
        print(f"\tdownloading {filename} successfull.")
    # REWRITE BAT FILES TO REFER TO CORREC
    for fn in [Path(x) for x in batfiles]:
         with fn.open("r") as f:
             d = f.read()
             d = replace_pythonexe_in_string(d)
             outfp = installdir / fn.with_suffix(".bat").name
             print("Creating BAT file from template: ", outfp)
             with outfp.open("w") as outfile: outfile.write(d)
    
install_paikkain()
