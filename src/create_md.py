import subprocess

# Call PowerShell script to get open Excel file addresses
cmd = 'powershell.exe -ExecutionPolicy RemoteSigned -File addres_output.ps1'
result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)

# Extract addresses from PowerShell output
addresses = result.stdout.decode('shift-jis').split()

# Create markdown text with file names and addresses
markdown_text = ''
for address in addresses:
    file_name = address.split('\\')[-1]
    markdown_text += f"[{file_name}](file:///{address})\n"

# Write markdown text to file
with open('file_addresses.md', 'w', encoding='shift-jis') as file:
    file.write(markdown_text)
