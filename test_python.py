import sys
import subprocess

print("Python version:", sys.version)
pip_version = subprocess.check_output(["pip", "--version"]).decode().strip()
print("pip version:", pip_version)