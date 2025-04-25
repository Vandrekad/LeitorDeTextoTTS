import tkinter as tk
from tkinter import scrolledtext, messagebox
import pyperclip
import pyttsx3
import threading
import time
from PIL import Image, ImageGrab
import pytesseract
import io

# --- Configuração Inicial ---
# [Windows Apenas] Descomente e ajuste se Tesseract não estiver no PATH
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\orlando.santos\AppData\Local\Programs\Tesseract-OCR'
except Exception as e:
    print(f"Aviso: Não foi possível definir o caminho do Tesseract: {e}")

# --- Funções de Acesso ao Clipboard ---
def obter_texto_clipboard():
    try:
        return pyperclip.paste()
    except Exception as e:
        messagebox.showerror("Erro Clipboard", f"Não foi possível acessar o texto da área de transferência.\n{e}")
        return ""

def obter_imagem_clipboard():
    try:
        imagem = ImageGrab.grabclipboard()
        if isinstance(imagem, Image.Image):
            return imagem
        else:
            # Poderia tentar ler de lista de arquivos se imagem fosse None
            # if isinstance(imagem, list) and imagem:
            #     try: return Image.open(imagem[0]) # Tenta abrir o primeiro arquivo
            #     except: return None
            return None
    except Exception as e:
        # Pode acontecer se o conteúdo não for imagem ou texto suportado
        print(f"Erro ao obter imagem do clipboard: {e}")
        # Não mostrar messagebox aqui para não ser intrusivo, o botão de OCR tratará a ausência
        return None

# --- Função OCR ---
def extrair_texto_de_imagem(imagem):
    if not isinstance(imagem, Image.Image):
        messagebox.showinfo("OCR", "Nenhuma imagem válida encontrada na área de transferência.")
        return ""
    try:
        # Tenta usar português, fallback para inglês se falhar ou se preferir
        texto_extraido = pytesseract.image_to_string(imagem, lang='por') # ou 'por+eng'
        if not texto_extraido.strip(): # Se OCR não retornar nada
             messagebox.showinfo("OCR", "Não foi possível extrair texto da imagem.")
             return ""
        return texto_extraido.strip()
    except pytesseract.TesseractNotFoundError:
        messagebox.showerror("Erro OCR", "Tesseract não encontrado. Verifique a instalação e a configuração do PATH (veja Passo 1 do guia).")
        return ""
    except Exception as e:
        messagebox.showerror("Erro OCR", f"Ocorreu um erro durante o OCR:\n{e}")
        return ""

# --- Configuração e Função TTS ---
engine = None
tts_engine_initialized = False
try:
    engine = pyttsx3.init()
    tts_engine_initialized = True
    # Configurações opcionais (descomente e ajuste)
    # voices = engine.getProperty('voices')
    # for voice in voices:
    #     if "brazil" in voice.name.lower() or "portuguese" in voice.name.lower():
    #         engine.setProperty('voice', voice.id)
    #         break
    # engine.setProperty('rate', 170)
    # engine.setProperty('volume', 1.0)
except Exception as e:
    print(f"Alerta: Não foi possível inicializar o motor TTS na inicialização: {e}")
    # Tentará inicializar de novo ao clicar em ler

def ler_texto_em_voz_alta(texto, janela_ref):
    global engine, tts_engine_initialized
    if not texto:
        messagebox.showinfo("Leitura", "Nenhum texto fornecido para leitura.", parent=janela_ref)
        return

    if not tts_engine_initialized:
        print("Motor TTS não inicializado. Tentando inicializar agora...")
        try:
            engine = pyttsx3.init()
            tts_engine_initialized = True
            print("Motor TTS inicializado com sucesso.")
            # Reaplicar configurações se necessário
        except Exception as e:
            messagebox.showerror("Erro TTS", f"Falha na inicialização do motor TTS:\n{e}", parent=janela_ref)
            return

    def tarefa_leitura():
        try:
            print(f"Lendo: {texto[:50]}...")
            engine.say(texto)
            engine.runAndWait()
            print("Leitura concluída.")
        except Exception as e:
            print(f"Erro durante a leitura: {e}")
            # Usar 'after' para garantir que messagebox rode na thread principal
            janela_ref.after(0, lambda: messagebox.showerror("Erro de Leitura", f"Ocorreu um erro ao tentar ler o texto:\n{e}", parent=janela_ref))

    thread_leitura = threading.Thread(target=tarefa_leitura)
    thread_leitura.daemon = True
    thread_leitura.start()

# --- Funções de Monitoramento ---
monitorar_clipboard = False
thread_monitoramento = None
ultimo_texto_clipboard = ""

def tarefa_monitoramento_clipboard(text_area_widget, janela_principal):
    global ultimo_texto_clipboard
    print("Monitoramento iniciado.")
    while monitorar_clipboard:
        try:
            texto_atual = pyperclip.paste()
            if isinstance(texto_atual, str) and texto_atual != ultimo_texto_clipboard and texto_atual:
                ultimo_texto_clipboard = texto_atual
                print(f"Novo texto detectado (monitor): {texto_atual[:50]}...")
                janela_principal.after(0, lambda ta=text_area_widget, txt=texto_atual: atualizar_texto_area_monitor(ta, txt))
        except Exception:
             # Ignorar erros silenciosamente no loop para evitar spam
             pass
        time.sleep(1) # Intervalo
    print("Monitoramento parado.")

def atualizar_texto_area_monitor(widget_texto, novo_texto):
    widget_texto.delete("1.0", tk.END)
    widget_texto.insert(tk.INSERT, novo_texto)

def iniciar_parar_monitoramento(text_area_widget, janela_principal, botao_monitor):
    global monitorar_clipboard, thread_monitoramento, ultimo_texto_clipboard
    if monitorar_clipboard:
        monitorar_clipboard = False
        botao_monitor.config(text="Iniciar Monitoramento")
    else:
        monitorar_clipboard = True
        botao_monitor.config(text="Parar Monitoramento")
        ultimo_texto_clipboard = "" # Resetar para pegar o atual
        thread_monitoramento = threading.Thread(target=tarefa_monitoramento_clipboard, args=(text_area_widget, janela_principal), daemon=True)
        thread_monitoramento.start()

# --- Interface Gráfica Principal ---
def criar_janela_preview():
    janela = tk.Tk()
    janela.title("Leitor de Tela Avançado")
    janela.geometry("600x550") # Aumentar tamanho
    janela.attributes('-topmost', True)

    # --- Widgets ---
    frame_instrucoes = tk.Frame(janela, bd=1, relief=tk.SUNKEN)
    frame_instrucoes.pack(pady=5, padx=10, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=(
        "Uso:\n"
        "- Copie texto (Ctrl+C) e clique 'Buscar Texto' OU 'Iniciar Monitoramento'.\n"
        "- Copie uma imagem (PrintScreen/Captura) e clique 'Ler de Imagem'.\n"
        "- Edite o texto na caixa e clique 'Ler em Voz Alta'."
    ), justify=tk.LEFT)
    label_instrucoes.pack(pady=5, padx=5)

    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=70, height=18, font=("Arial", 11))
    text_area.pack(pady=5, padx=10, expand=True, fill='both')

    # Frame para botões de ação
    frame_botoes_acao = tk.Frame(janela)
    frame_botoes_acao.pack(pady=(5, 0))

    # Frame para botões de controle
    frame_botoes_controle = tk.Frame(janela)
    frame_botoes_controle.pack(pady=(0, 10))


    # --- Funções dos Botões ---
    def acao_buscar_texto():
        texto_cb = obter_texto_clipboard()
        if texto_cb: # Só atualiza se houver texto
            text_area.delete("1.0", tk.END)
            text_area.insert(tk.INSERT, texto_cb)

    def acao_ler_imagem():
        imagem_cb = obter_imagem_clipboard()
        if imagem_cb:
            texto_ocr = extrair_texto_de_imagem(imagem_cb)
            if texto_ocr:
                text_area.delete("1.0", tk.END)
                text_area.insert(tk.INSERT, texto_ocr)
                # Opcional: Ler automaticamente após OCR?
                # ler_texto_em_voz_alta(texto_ocr, janela)
        else:
             messagebox.showinfo("OCR", "Nenhuma imagem encontrada na área de transferência.", parent=janela)


    def acao_ler_texto_area():
        texto_para_ler = text_area.get("1.0", tk.END).strip()
        ler_texto_em_voz_alta(texto_para_ler, janela)

    # --- Botões ---
    # Ação
    botao_buscar_texto = tk.Button(frame_botoes_acao, text="Buscar Texto (Clipboard)", command=acao_buscar_texto, width=20, height=2)
    botao_buscar_texto.pack(side=tk.LEFT, padx=5, pady=5)

    botao_ler_imagem = tk.Button(frame_botoes_acao, text="Ler de Imagem (Clipboard)", command=acao_ler_imagem, width=20, height=2)
    botao_ler_imagem.pack(side=tk.LEFT, padx=5, pady=5)

    botao_ler_audio = tk.Button(frame_botoes_acao, text="Ler Texto em Voz Alta", command=acao_ler_texto_area, width=20, height=2, fg="blue")
    botao_ler_audio.pack(side=tk.LEFT, padx=5, pady=5)

    # Controle
    botao_monitor = tk.Button(frame_botoes_controle, text="Iniciar Monitoramento", width=20, height=2)
    # Comando precisa ser definido após a criação do botão para passar a si mesmo
    botao_monitor.config(command=lambda: iniciar_parar_monitoramento(text_area, janela, botao_monitor))
    botao_monitor.pack(side=tk.LEFT, padx=5, pady=5)


    botao_fechar = tk.Button(frame_botoes_controle, text="Fechar", command=janela.destroy, width=10, height=2)
    botao_fechar.pack(side=tk.LEFT, padx=5, pady=5)

    # --- Tratamento ao Fechar Janela ---
    def ao_fechar():
        global monitorar_clipboard
        print("Fechando a aplicação...")
        monitorar_clipboard = False # Garante que a thread de monitoramento pare
        # Esperar um pouco se a thread estiver ativa? Opcional.
        # if thread_monitoramento and thread_monitoramento.is_alive():
        #    thread_monitoramento.join(timeout=0.5)
        if tts_engine_initialized and engine._inLoop:
             engine.endLoop() # Tenta parar o TTS se estiver falando
        janela.destroy()

    janela.protocol("WM_DELETE_WINDOW", ao_fechar)

    # --- Iniciar a GUI ---
    janela.mainloop()

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    criar_janela_preview()