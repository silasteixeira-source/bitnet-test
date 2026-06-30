import sys
import os

def apply_fix():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    old_func = """def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))"""

    new_func = """def get_base_path():
    import os
    import sys
    # 1. Tentar pegar o diretório de instalação do Inno Setup
    install_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'FeramentaEACEBitnet')
    if os.path.exists(install_dir) and os.path.exists(os.path.join(install_dir, 'AuthNOC.json')):
        return install_dir
        
    # 2. Tentar o diretório do executável
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        if "_MEI" not in exe_dir:
            return exe_dir
            
    # 3. Diretório atual de trabalho (Start In do atalho)
    cwd = os.getcwd()
    if os.path.exists(os.path.join(cwd, 'AuthNOC.json')):
        return cwd
        
    return os.path.dirname(os.path.abspath(__file__))"""

    if old_func in content:
        content = content.replace(old_func, new_func)
        print("get_base_path corrigido com sucesso!")
    else:
        print("Função original não encontrada. Talvez já tenha sido alterada?")

    with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
        f.write(content)

if __name__ == '__main__':
    apply_fix()
