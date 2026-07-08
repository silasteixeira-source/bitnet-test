import sys

def apply_fix():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Corrigir loop para Grid lado a lado
    old_loop = """        for text, cmd, color in buttons:
            ModernButton(f_botoes, text=text, command=cmd, bg_type=color, font=("Segoe UI", 11, "bold"), radius=16, height=46, width=380).pack(pady=10)"""

    new_loop = """        for i, (text, cmd, color) in enumerate(buttons):
            row = i // 2
            col = i % 2
            btn = ModernButton(f_botoes, text=text, command=cmd, bg_type=color, font=("Segoe UI", 11, "bold"), radius=16, height=46, width=360)
            btn.grid(row=row, column=col, padx=10, pady=10)"""

    if old_loop in content:
        content = content.replace(old_loop, new_loop)
    else:
        print("Loop antigo não encontrado. Pode já estar modificado.")

    # Corrigir caracteres perdidos no patch anterior
    content = content.replace('"Rob Leitor', '"Robô Leitor')
    content = content.replace('"Rob Leitor', '"Robô Leitor')
    content = content.replace('Unificador de Evidncias', 'Unificador de Evidências')
    content = content.replace('Conversor Bidirecional PDF  Word', 'Conversor Bidirecional PDF ⇄ Word')
    
    # Garantir que a assinatura do RoboLeitor aceite gs (gspread) caso não tenha sido ajustada
    sig_old = "def __init__(self, parent_frame, main_app_callback):"
    sig_new = "def __init__(self, parent_frame, main_app_callback, gspread_client=None):"
    content = content.replace(sig_old, sig_new)

    with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
        f.write(content)
        
    print("Correções aplicadas!")

if __name__ == '__main__':
    apply_fix()
