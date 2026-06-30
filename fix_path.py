import sys

def apply_fix():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Corrigir o path em authenticate_gmail_readonly
    old_code = 'base_path = os.path.dirname(os.path.abspath(__file__))'
    new_code = 'base_path = get_base_path()'

    if old_code in content:
        content = content.replace(old_code, new_code)
        print("Path fix applied!")
    else:
        print("Path fix not found, maybe already applied.")

    with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
        f.write(content)

if __name__ == '__main__':
    apply_fix()
