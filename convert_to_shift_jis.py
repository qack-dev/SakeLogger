import sys
import os

def convert_to_shift_jis(filepath):
    try:
        # Read content as UTF-8
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Write content as Shift-JIS
        with open(filepath, 'w', encoding='shift_jis') as f:
            f.write(content)
        print(f"Successfully converted {filepath} to Shift-JIS.")
    except Exception as e:
        print(f"Error converting {filepath}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_to_shift_jis.py <file1> <file2> ...")
        sys.exit(1)
    
    for filepath in sys.argv[1:]:
        if os.path.exists(filepath):
            convert_to_shift_jis(filepath)
        else:
            print(f"File not found: {filepath}")
