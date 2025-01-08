import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "uninstall", package])

# List of packages to install
packages = [
    "openpyxl",
    "Pillow",
    "docxtpl",
    "docx2pdf",
    "reportlab",
    "PyPDF2",
    "PyMuPDF"
]

for package in packages:
    try:
        install(package)
        print(f"{package} installed successfully.")
    except Exception as e:
        print(f"Failed to install {package}: {e}")
