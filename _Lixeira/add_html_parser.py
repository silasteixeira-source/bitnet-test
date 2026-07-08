import sys
import re

def apply_html_extraction():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    pattern = r'def extrair_texto_da_mensagem\(self, payload\):.*?return "Nenhum texto puro encontrado na mensagem\."'
    
    nova_funcao = """def extrair_texto_da_mensagem(self, payload):
        try:
            def buscar_partes(p):
                mime = p.get("mimeType")
                
                # 1. Se achou texto puro, retorna imediatamente
                if mime == "text/plain":
                    d = p.get("body", {}).get("data")
                    if d:
                        import base64
                        return base64.urlsafe_b64decode(d).decode("utf-8", errors="ignore")
                
                # 2. Se achou HTML e não tem filhas, extrai e limpa as tags
                if mime == "text/html" and "parts" not in p:
                    d = p.get("body", {}).get("data")
                    if d:
                        import base64
                        import re
                        html = base64.urlsafe_b64decode(d).decode("utf-8", errors="ignore")
                        
                        html = re.sub(r'<style.*?>.*?</style>', '', html, flags=re.IGNORECASE|re.DOTALL)
                        html = re.sub(r'<script.*?>.*?</script>', '', html, flags=re.IGNORECASE|re.DOTALL)
                        html = re.sub(r'<br\s*/?>', '\\n', html, flags=re.IGNORECASE)
                        html = re.sub(r'</p>', '\\n\\n', html, flags=re.IGNORECASE)
                        
                        texto = re.sub(r'<[^>]+>', ' ', html)
                        texto = texto.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
                        texto = re.sub(r'[ \\t]+', ' ', texto)
                        texto = re.sub(r'\\n\\s*\\n', '\\n\\n', texto)
                        
                        return texto.strip()

                # 3. Se for multipart, navega pelas partes (Prioridade para o text/plain)
                if "parts" in p:
                    for part in p["parts"]:
                        if part.get("mimeType") == "text/plain":
                            res = buscar_partes(part)
                            if res: return res
                    
                    for part in p["parts"]:
                        res = buscar_partes(part)
                        if res: return res

                return None

            texto_final = buscar_partes(payload)
            if texto_final:
                return texto_final
                
        except Exception as e:
            return f"[Erro ao extrair corpo da mensagem: {e}]"
            
        return "Conteúdo da mensagem não suportado ou vazio." """

    if re.search(pattern, content, flags=re.DOTALL):
        # Passar lambda para evitar problemas de escape \s \n no re.sub
        content = re.sub(pattern, lambda m: nova_funcao, content, flags=re.DOTALL)
        with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
            f.write(content)
        print("Nova inteligência de HTML injetada com sucesso (Fix aplicado)!")
    else:
        print("ERRO: Não encontrou o padrão para substituir.")

if __name__ == "__main__":
    apply_html_extraction()
