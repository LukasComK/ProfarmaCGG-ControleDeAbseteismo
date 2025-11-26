#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to indent all code from line 320+ into the if modo == "Processar" block
"""

file_path = r'c:\Users\Lukas\Desktop\Controle de Abseteismo\app.py'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the line with mode check - "if modo == "..." and "if files_encarregado:"
mode_check_line = None
files_check_line = None

for i, line in enumerate(lines):
    if 'if modo == "📊 Processar Absenteísmo":' in line:
        mode_check_line = i
    if 'if files_encarregado:' in line and i > mode_check_line if mode_check_line else False:
        files_check_line = i
        break

print(f"Mode check at line: {mode_check_line + 1 if mode_check_line else 'NOT FOUND'}")
print(f"Files check at line: {files_check_line + 1 if files_check_line else 'NOT FOUND'}")

# We need to indent from files_check_line to the end of the file
# Get current indentation level of the "if files_encarregado:" line
if files_check_line:
    current_indent = len(lines[files_check_line]) - len(lines[files_check_line].lstrip())
    print(f"Current indent of 'if files_encarregado:': {current_indent} spaces")
    
    # We need to add 4 more spaces to all lines from files_check_line to EOF
    new_lines = lines[:files_check_line]  # Keep everything before
    
    for i in range(files_check_line, len(lines)):
        line = lines[i]
        # Don't indent empty lines or lines that are only whitespace
        if line.strip():
            new_lines.append('    ' + line)  # Add 4 spaces
        else:
            new_lines.append(line)  # Keep empty lines as-is
    
    # Write back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    
    print(f"\nIndented {len(lines) - files_check_line} lines")
    print("Done!")
else:
    print("Could not find the files_encarregado check line")
