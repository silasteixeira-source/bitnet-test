import sys
import re

def apply_splash_patch():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    old_main = """    gs = authenticate_google_sheets()
    root.deiconify()"""

    new_main = """    gs = authenticate_google_sheets()
    
    # Fecha o Splash Screen nativo do PyInstaller antes de renderizar a janela principal
    try:
        import pyi_splash
        pyi_splash.update_text("Quase lá...")
        pyi_splash.close()
    except ImportError:
        pass
        
    root.deiconify()"""

    if old_main in content:
        content = content.replace(old_main, new_main)
        print("Splash Screen logic injetada em FerramentaEACEv6.py com sucesso!")
    else:
        print("Trecho de código não encontrado. Talvez já injetado?")
        
    with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
        f.write(content)

if __name__ == "__main__":
    apply_splash_patch()
