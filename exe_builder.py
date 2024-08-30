import os
import subprocess
import sys

# Get the path to the current script's directory
script_dir = os.path.dirname(os.path.realpath(__file__))

# Get the path to the user's desktop
desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")

# Construct the correct path for Template.docx
template_path = os.path.join(script_dir, 'Template.docx')

# Construct the correct --add-data argument
if sys.platform.startswith('win'):
    add_data_arg = f"{template_path};."
else:
    add_data_arg = f"{template_path}:."

command = [
    'pyinstaller',
    '--noconfirm',
    '--onefile',
    '--noconsole',
    '--icon', os.path.join(script_dir, 'fsrg.ico'),
    '--add-data', add_data_arg,
    os.path.join(script_dir, 'gui.py'),
    '--distpath', desktop_dir
]

print("Running PyInstaller with the following command:")
print(" ".join(command))

result = subprocess.run(command, capture_output=True, text=True)

print("PyInstaller Output:")
print(result.stdout)

if result.returncode != 0:
    print("Error occurred:")
    print(result.stderr)
else:
    print("PyInstaller completed successfully.")