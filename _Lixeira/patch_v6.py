import sys

def run_patch():
    try:
        with open("Rascunho_Robo_Leitor.py", "r", encoding="utf-8") as f:
            rascunho = f.read()
            
        # Extract components
        auth_start = rascunho.find("def authenticate_gmail_readonly():")
        auth_end = rascunho.find("# --- 2. NOVA CLASSE DO APLICATIVO DO ROBÔ ---")
        auth_code = rascunho[auth_start:auth_end].strip() + "\n\n"
        
        class_start = rascunho.find("class RoboLeitorGmailApp:")
        class_end = rascunho.find("if __name__ == \"__main__\":")
        # Remover o comentário extra antes do if __name__
        class_code = rascunho[class_start:class_end].strip()
        class_code = class_code.split("# (Abaixo dessa classe")[0].strip() + "\n\n"
        
        with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
            v6_lines = f.readlines()
            
        new_lines = []
        for line in v6_lines:
            if line.startswith("import base64"):
                new_lines.append(line)
                new_lines.append("from email.utils import parsedate_to_datetime\n")
            elif line.startswith("class EnviadorPleitoApp:"):
                new_lines.append(auth_code + "\n")
                new_lines.append(line)
            elif line.startswith("class MainApp:"):
                new_lines.append(class_code + "\n")
                new_lines.append(line)
            elif '("Dashboard EACE"' in line:
                new_lines.append(line.rstrip('\r\n') + ",\n")
                new_lines.append('            ("Robô Leitor do Gmail", lambda: self.open(RoboLeitorGmailApp), "#e91e63")\n')
            elif 'tk.Label(footer, text="v3.0' in line or 'tk.Label(footer, text="v5.0' in line:
                new_lines.append(line.replace("v3.0", "v6.0").replace("v5.0", "v6.0"))
            else:
                new_lines.append(line)
                
        with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
            f.writelines(new_lines)
        print("Patch concluído com sucesso!")
        
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == '__main__':
    run_patch()
