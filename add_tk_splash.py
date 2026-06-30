import sys
import re

def apply_tk_splash():
    with open("FerramentaEACEv6.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Remover a parte do main atual
    main_pattern = r'if __name__ == "__main__":.*'
    
    new_main = """def show_custom_splash_and_load(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.geometry("500x350")
    
    # centralizar
    x = (root.winfo_screenwidth() // 2) - (500 // 2)
    y = (root.winfo_screenheight() // 2) - (350 // 2)
    splash.geometry(f"+{x}+{y}")
    
    splash.config(bg="#1a1a1a")
    
    try:
        from PIL import Image, ImageTk
        gif_path = resource_path("logo2.gif")
        gif = Image.open(gif_path)
        frames = []
        try:
            while True:
                frames.append(ImageTk.PhotoImage(gif.copy()))
                gif.seek(len(frames))
        except EOFError:
            pass
    except Exception as e:
        frames = []
        print("Erro ao carregar GIF:", e)
        
    lbl_img = tk.Label(splash, bg="#1a1a1a")
    lbl_img.pack(pady=(40, 20), expand=True)
    
    lbl_status = tk.Label(splash, text="Iniciando e autenticando (Google)...", bg="#1a1a1a", fg="white", font=("Segoe UI", 11, "bold"))
    lbl_status.pack(pady=5)
    
    from tkinter import ttk
    progress = ttk.Progressbar(splash, orient="horizontal", length=300, mode="indeterminate")
    progress.pack(pady=10)
    progress.start(15)
    
    def animate_gif(idx=0):
        if frames:
            lbl_img.config(image=frames[idx])
            idx = (idx + 1) % len(frames)
            splash.after(60, animate_gif, idx)
            
    animate_gif()
    
    result = [None]
    
    def load_task():
        try:
            result[0] = authenticate_google_sheets()
        except Exception as e:
            result[0] = None
            print("Erro auth:", e)
            
    import threading
    t = threading.Thread(target=load_task)
    t.start()
    
    def check_done():
        if t.is_alive():
            splash.after(100, check_done)
        else:
            progress.stop()
            splash.destroy()
            root.deiconify()
            root.title("Ferramentas de Apoio - EACE")
            root.geometry("900x780")
            app = MainApp(root, result[0])
            
    check_done()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    root.withdraw()
    show_custom_splash_and_load(root)
    root.mainloop()"""

    content = re.sub(main_pattern, new_main, content, flags=re.DOTALL)

    with open("FerramentaEACEv6.py", "w", encoding="utf-8") as f:
        f.write(content)
        
    print("Tkinter Splash adicionado!")

if __name__ == "__main__":
    apply_tk_splash()
