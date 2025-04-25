import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk # Adicionado ttk
import pyperclip
# from openai import OpenAI # Removido
# import sounddevice as sd # Removido
# import soundfile as sf # Removido
# import numpy as np # Removido
import threading
import time
from PIL import Image, ImageGrab
import pytesseract
import io
import os
import tempfile
import platform
import sys

# --- Verificação do Sistema Operacional ---
is_windows = platform.system() == "Windows"
if not is_windows:
    print("ERRO: Esta versão do script usa pywin32 e só funciona no Windows.")
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Erro de Plataforma", "Esta aplicação requer Windows (pywin32).")
        root.destroy()
    except Exception: pass
    sys.exit(1)

# Tenta importar pywin32 apenas se for Windows
try:
    import win32com.client as wincl
    try:
        import pythoncom
        pythoncom.CoInitialize()
        print("COM inicializado para thread principal.")
    except ImportError: print("Aviso: pythoncom não encontrado.")
    except Exception as com_err: print(f"Aviso: Erro ao inicializar COM: {com_err}")
    pywin32_ok = True
    print("pywin32 importado com sucesso.")
except ImportError:
    print("ERRO: pywin32 não está instalado. Execute: pip install pywin32")
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Erro Dependência", "'pywin32' necessário.\nExecute 'pip install pywin32'.")
        root.destroy()
    except Exception: pass
    sys.exit(1)
except Exception as e:
     print(f"ERRO inesperado ao importar pywin32: {e}")
     pywin32_ok = False


# --- Configuração Inicial ---
# [Windows Apenas] Tesseract Path (se necessário)
# try:
#     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# except Exception as e:
#     print(f"Aviso: Não definiu caminho Tesseract: {e}")

# --- Variáveis Globais ---
audio_thread = None
stop_audio_flag = threading.Event()
sapi_voice_object = None
available_sapi_voices = {} # Dicionário para vozes {descrição: índice}
# selected_voice_index = tk.IntVar(value=-1) # MOVIDO PARA DENTRO DE criar_janela_preview
# selected_speed = tk.StringVar(value="Normal") # MOVIDO PARA DENTRO DE criar_janela_preview

# --- Funções Clipboard e OCR ---
def obter_texto_clipboard():
    try: return pyperclip.paste()
    except Exception as e: messagebox.showerror("Erro Clipboard", f"Não acesso texto.\n{e}"); return ""
def obter_imagem_clipboard():
    try:
        imagem = ImageGrab.grabclipboard(); return imagem if isinstance(imagem, Image.Image) else None
    except Exception as e: print(f"Erro obter imagem: {e}"); return None
def extrair_texto_de_imagem(imagem):
    if not isinstance(imagem, Image.Image): messagebox.showinfo("OCR", "Nenhuma imagem válida."); return ""
    try:
        lang = 'por'; print(f"Executando OCR: {lang}")
        texto_extraido = pytesseract.image_to_string(imagem, lang=lang)
        if not texto_extraido.strip(): messagebox.showinfo("OCR", "Não extraiu texto."); return ""
        print("Texto extraído."); return texto_extraido.strip()
    except PermissionError as pe: print(f"ERRO PERMISSÃO Tesseract: {pe}"); messagebox.showerror("Erro Permissão OCR", f"Acesso negado Tesseract:\n{pe}"); return ""
    except pytesseract.TesseractNotFoundError: messagebox.showerror("Erro OCR", "Tesseract não encontrado."); return ""
    except Exception as e: messagebox.showerror("Erro OCR", f"Erro inesperado OCR:\n{e}"); return ""

# --- Função TTS com SAPI (pywin32) ---
def inicializar_e_listar_sapi_voices(voice_var_ref): # Recebe a referência da variável Tkinter
    """Inicializa SAPI, lista vozes e preenche o dicionário global."""
    global sapi_voice_object, available_sapi_voices
    if not pywin32_ok: return False # Sai se pywin32 não importou

    if sapi_voice_object is None:
        try:
            print("Inicializando objeto SAPI SpVoice...")
            try: # Inicializa COM para a thread atual (principal ou outra)
                 import pythoncom
                 pythoncom.CoInitialize()
            except ImportError: print("Aviso: pythoncom não encontrado.")
            except Exception as com_err:
                 if hasattr(com_err, 'hresult') and com_err.hresult in [-2147417850, -2147221008]: pass # Ignora se já inicializado
                 else: print(f"Aviso: Erro COM: {com_err}")

            sapi_voice_object = wincl.Dispatch("SAPI.SpVoice")
            print("Objeto SAPI inicializado.")

            # Limpa e preenche o dicionário de vozes
            available_sapi_voices.clear()
            voices = sapi_voice_object.GetVoices()
            print("Vozes SAPI disponíveis:")
            default_voice_desc = sapi_voice_object.Voice.GetDescription() # Pega a descrição da voz padrão
            default_voice_index = 0 # Assume 0 como padrão inicial
            for i, v in enumerate(voices):
                description = v.GetDescription()
                print(f"  {i}: {description}")
                available_sapi_voices[description] = i # Armazena descrição -> índice
                if description == default_voice_desc:
                     default_voice_index = i # Encontra o índice da voz padrão

            # Define a variável Tkinter para a voz padrão (usando a referência)
            voice_var_ref.set(default_voice_index)
            print(f"Voz padrão definida: {default_voice_desc} (Índice: {default_voice_index})")
            return True

        except Exception as e:
            print(f"ERRO CRÍTICO ao inicializar SAPI ou listar vozes: {e}")
            messagebox.showerror("Erro SAPI", f"Não foi possível inicializar SAPI ou listar vozes.\n{e}")
            sapi_voice_object = None
            available_sapi_voices.clear()
            return False
    return True # Já estava inicializado

def ler_texto_pywin32(texto, voice_idx, speed_str, janela_ref, botao_ler, botao_parar):
    """Usa SAPI (pywin32) para ler o texto com voz e velocidade selecionadas."""
    global audio_thread, stop_audio_flag, sapi_voice_object # sapi_voice_object é usado para verificar inicialização

    # Verifica se SAPI foi inicializado com sucesso antes
    # A inicialização principal agora acontece antes da janela ser criada
    if not sapi_voice_object:
        messagebox.showerror("Erro SAPI", "Motor de voz SAPI não inicializado corretamente.", parent=janela_ref)
        return # Sai se falhar

    if audio_thread and audio_thread.is_alive():
        messagebox.showwarning("Leitura", "Leitura em andamento.", parent=janela_ref)
        return

    stop_audio_flag.clear()

    if not texto:
        messagebox.showinfo("Leitura", "Nenhum texto.", parent=janela_ref)
        return

    # Feedback visual e botões
    janela_ref.config(cursor="watch")
    botao_ler.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL)
    janela_ref.update_idletasks()

    def tarefa_leitura_sapi():
        """Função executada na thread para controlar a fala SAPI."""
        nonlocal texto, voice_idx, speed_str, janela_ref, botao_ler, botao_parar
        speak_obj = None # Objeto SAPI específico para esta thread
        com_initialized_thread = False # Flag para controlar CoUninitialize
        try:
            # Inicializa COM e SAPI *dentro* da thread
            try:
                 import pythoncom; pythoncom.CoInitialize(); com_initialized_thread = True
            except ImportError: print("Aviso: pythoncom não na thread.")
            except Exception as com_err:
                 if hasattr(com_err, 'hresult') and com_err.hresult in [-2147417850, -2147221008]: com_initialized_thread = True # Considera inicializado se já estava
                 else: print(f"Aviso: Erro COM thread: {com_err}")

            if not com_initialized_thread: raise Exception("Falha ao inicializar COM na thread.")

            speak_obj = wincl.Dispatch("SAPI.SpVoice") # Cria instância SAPI para a thread

            # --- Define Voz e Velocidade ---
            try:
                 voices = speak_obj.GetVoices()
                 if 0 <= voice_idx < voices.Count:
                      speak_obj.Voice = voices.Item(voice_idx)
                      print(f"Voz definida na thread: {speak_obj.Voice.GetDescription()}")
                 else: print(f"Aviso: Índice de voz inválido ({voice_idx}). Usando padrão da thread.")
                 speed_map = {"Lenta": -3, "Normal": 0, "Rápida": 3}
                 rate_value = speed_map.get(speed_str, 0)
                 speak_obj.Rate = rate_value
                 print(f"Velocidade definida na thread: {rate_value} ({speed_str})")
            except Exception as e_prop: print(f"Aviso: Erro definir voz/velocidade: {e_prop}")
            # -----------------------------

            SVSFlagsAsync = 1; SVSFPurgeBeforeSpeak = 2
            print(f"Falando com SAPI: {texto[:50]}...")
            speak_obj.Speak(texto, SVSFlagsAsync | SVSFPurgeBeforeSpeak)
            time.sleep(0.5)

            while speak_obj.Status.RunningState == 2:
                if stop_audio_flag.is_set():
                    print("Parada solicitada...")
                    speak_obj.Speak("", SVSFlagsAsync | SVSFPurgeBeforeSpeak)
                    break
                time.sleep(0.1)

            if not stop_audio_flag.is_set(): print("Fala SAPI concluída.")
            else: print("Fala SAPI interrompida.")

        except Exception as e:
            print(f"Erro durante a fala SAPI: {e}")
            janela_ref.after(0, lambda err=e: messagebox.showerror("Erro SAPI", f"Erro durante a fala:\n{err}", parent=janela_ref))
        finally:
            def restaurar_gui_thread_safe():
                janela_ref.config(cursor=""); botao_ler.config(state=tk.NORMAL); botao_parar.config(state=tk.DISABLED)
            janela_ref.after(0, restaurar_gui_thread_safe)
            if com_initialized_thread: # Só desinicializa se inicializou com sucesso
                try: import pythoncom; pythoncom.CoUninitialize(); print("COM da thread desinicializado.")
                except ImportError: pass
                except Exception as com_err: print(f"Aviso: Erro desinicializar COM thread: {com_err}")

    audio_thread = threading.Thread(target=tarefa_leitura_sapi); audio_thread.daemon = True; audio_thread.start()

# --- Função para Parar Leitura (SAPI) ---
def acao_parar_leitura():
    """Sinaliza para a thread de áudio SAPI parar a reprodução."""
    global stop_audio_flag, audio_thread
    if audio_thread and audio_thread.is_alive(): print("Sinalizando parada SAPI..."); stop_audio_flag.set()
    else: print("Nenhuma leitura SAPI ativa.")

# --- Funções de Monitoramento ---
monitorar_clipboard = False; thread_monitoramento = None; ultimo_texto_clipboard = ""
def tarefa_monitoramento_clipboard(text_area_widget, janela_principal):
    global ultimo_texto_clipboard; print("Monitoramento iniciado.")
    while monitorar_clipboard:
        try:
            texto_atual = pyperclip.paste()
            if isinstance(texto_atual, str) and texto_atual != ultimo_texto_clipboard and texto_atual:
                ultimo_texto_clipboard = texto_atual; print(f"Novo texto (monitor): {texto_atual[:50]}...")
                janela_principal.after(0, lambda ta=text_area_widget, txt=texto_atual: atualizar_texto_area_monitor(ta, txt))
        except Exception: pass
        time.sleep(1.0)
    print("Monitoramento parado.")
def atualizar_texto_area_monitor(widget_texto, novo_texto): widget_texto.delete("1.0", tk.END); widget_texto.insert(tk.INSERT, novo_texto)
def iniciar_parar_monitoramento(text_area_widget, janela_principal, botao_monitor):
    global monitorar_clipboard, thread_monitoramento, ultimo_texto_clipboard
    if monitorar_clipboard: monitorar_clipboard = False; botao_monitor.config(text="Iniciar Monitoramento", bg="SystemButtonFace", fg="SystemButtonText")
    else:
        monitorar_clipboard = True; botao_monitor.config(text="Parar Monitoramento", bg="red", fg="white")
        try: ultimo_texto_clipboard = pyperclip.paste()
        except: ultimo_texto_clipboard = ""
        thread_monitoramento = threading.Thread(target=tarefa_monitoramento_clipboard, args=(text_area_widget, janela_principal), daemon=True); thread_monitoramento.start()

# --- Interface Gráfica Principal ---
def criar_janela_preview():
    # *** CORREÇÃO APLICADA AQUI ***
    # Move a criação das variáveis Tkinter para depois da criação da janela
    global selected_voice_index, selected_speed # Declara que usaremos as globais

    janela = tk.Tk(); janela.title("Leitor de Tela (OCR + SAPI TTS + Seletores)"); janela.geometry("700x600"); janela.attributes('-topmost', True)

    # Instancia as variáveis Tkinter AQUI
    selected_voice_index = tk.IntVar(value=-1) # Valor inicial -1 indica não definido
    selected_speed = tk.StringVar(value="Normal")

    # Chama a inicialização SAPI que agora também define selected_voice_index
    # Passa a referência da variável para a função poder defini-la
    if not inicializar_e_listar_sapi_voices(selected_voice_index):
         print("Falha ao inicializar SAPI na GUI. Leitura desabilitada.")
         # Poderia desabilitar os botões de leitura aqui

    frame_instrucoes = tk.Frame(janela, bd=1, relief=tk.SUNKEN); frame_instrucoes.pack(pady=5, padx=10, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=("Uso (Windows SAPI):\n- Copie texto/imagem e use 'Buscar'/'Ler Imagem'.\n- Selecione Voz e Velocidade.\n- Clique 'Ler com SAPI' para ouvir. Use 'Parar' para interromper."), justify=tk.LEFT); label_instrucoes.pack(pady=5, padx=5)
    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=70, height=18, font=("Arial", 11)); text_area.pack(pady=5, padx=10, expand=True, fill='both')

    # --- Frame para Seletores ---
    frame_selectors = tk.Frame(janela)
    frame_selectors.pack(pady=5, padx=10, fill=tk.X)

    # Seletor de Voz SAPI
    label_voice = tk.Label(frame_selectors, text="Voz SAPI:")
    label_voice.pack(side=tk.LEFT, padx=(0, 5))
    voice_options = list(available_sapi_voices.keys())
    # Usa a variável global selected_voice_index para o combobox, mas mostra descrições
    voice_combo = ttk.Combobox(frame_selectors, values=voice_options, state='readonly', width=40)
    # Tenta definir o valor inicial baseado no índice salvo em selected_voice_index
    initial_voice_index = selected_voice_index.get()
    initial_voice_desc = ""
    for desc, idx in available_sapi_voices.items():
        if idx == initial_voice_index:
            initial_voice_desc = desc
            break
    if initial_voice_desc:
        voice_combo.set(initial_voice_desc)
    elif voice_options:
        voice_combo.current(0) # Seleciona o primeiro se não achou o padrão

    voice_combo.pack(side=tk.LEFT, padx=5)
    if not available_sapi_voices: voice_combo.config(state=tk.DISABLED)

    # Seletor de Velocidade SAPI
    label_speed = tk.Label(frame_selectors, text="Velocidade:")
    label_speed.pack(side=tk.LEFT, padx=(10, 5))
    speed_options = ["Lenta", "Normal", "Rápida"]
    speed_combo = ttk.Combobox(frame_selectors, textvariable=selected_speed, values=speed_options, state='readonly', width=10)
    speed_combo.pack(side=tk.LEFT, padx=5)

    # --- Frames para Botões ---
    frame_botoes_acao = tk.Frame(janela); frame_botoes_acao.pack(pady=(5, 0))
    frame_botoes_controle = tk.Frame(janela); frame_botoes_controle.pack(pady=(0, 10))

    # --- Botões (Declarados antes) ---
    botao_ler_imagem = tk.Button(frame_botoes_acao, text="Ler Imagem", width=15, height=2)
    botao_ler_audio = tk.Button(frame_botoes_acao, text="Ler com SAPI", width=15, height=2, fg="blue")
    botao_parar_audio = tk.Button(frame_botoes_acao, text="Parar Leitura", command=acao_parar_leitura, width=12, height=2, state=tk.DISABLED)

    # --- Funções dos Botões ---
    def acao_buscar_texto(): texto_cb = obter_texto_clipboard(); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_cb) if texto_cb else None
    def acao_ler_imagem():
        janela.config(cursor="watch"); botao_ler_imagem.config(state=tk.DISABLED); botao_ler_audio.config(state=tk.DISABLED); botao_parar_audio.config(state=tk.DISABLED)
        janela.update_idletasks()
        try:
            imagem_cb = obter_imagem_clipboard()
            if imagem_cb: texto_ocr = extrair_texto_de_imagem(imagem_cb); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_ocr) if texto_ocr else None
            else: messagebox.showinfo("OCR", "Nenhuma imagem.", parent=janela)
        finally: janela.config(cursor=""); botao_ler_imagem.config(state=tk.NORMAL); botao_ler_audio.config(state=tk.NORMAL)

    def acao_ler_texto_area_sapi():
        global selected_voice_index, selected_speed # Acessa as globais
        texto_para_ler = text_area.get("1.0", tk.END).strip()
        selected_voice_desc = voice_combo.get()
        voice_index_to_use = available_sapi_voices.get(selected_voice_desc, -1)
        speed_string_to_use = selected_speed.get() # Usa a variável global

        if voice_index_to_use == -1 and available_sapi_voices:
             print(f"Aviso: Descrição '{selected_voice_desc}' não encontrada. Usando voz padrão.")
             # Tenta pegar o índice da variável global se a busca falhar
             # Isso pode não ser necessário se a inicialização sempre definir um valor válido
             default_idx_from_var = selected_voice_index.get()
             if default_idx_from_var != -1:
                 voice_index_to_use = default_idx_from_var
             else: # Último recurso: usa índice 0
                  voice_index_to_use = 0


        ler_texto_pywin32(texto_para_ler, voice_index_to_use, speed_string_to_use, janela, botao_ler_audio, botao_parar_audio)

    # --- Configuração final e Empacotamento dos Botões ---
    botao_buscar_texto = tk.Button(frame_botoes_acao, text="Buscar Texto", command=acao_buscar_texto, width=15, height=2); botao_buscar_texto.pack(side=tk.LEFT, padx=5, pady=5)
    botao_ler_imagem.config(command=acao_ler_imagem); botao_ler_imagem.pack(side=tk.LEFT, padx=5, pady=5)
    botao_ler_audio.config(command=acao_ler_texto_area_sapi); botao_ler_audio.pack(side=tk.LEFT, padx=5, pady=5)
    botao_parar_audio.pack(side=tk.LEFT, padx=5, pady=5)
    botao_monitor = tk.Button(frame_botoes_controle, text="Iniciar Monitoramento", width=20, height=2); botao_monitor.config(command=lambda: iniciar_parar_monitoramento(text_area, janela, botao_monitor)); botao_monitor.pack(side=tk.LEFT, padx=5, pady=5)
    botao_fechar = tk.Button(frame_botoes_controle, text="Fechar", command=lambda: ao_fechar(janela), width=10, height=2); botao_fechar.pack(side=tk.LEFT, padx=5, pady=5)

    # --- Tratamento ao Fechar Janela ---
    def ao_fechar(janela_a_fechar):
        global monitorar_clipboard, stop_audio_flag, audio_thread, sapi_voice_object
        print("Fechando..."); monitorar_clipboard = False; stop_audio_flag.set()
        if sapi_voice_object:
             try: sapi_voice_object.Speak("", 3)
             except Exception as e: print(f"Info: Erro SAPI ao fechar: {e}")
        if audio_thread and audio_thread.is_alive(): print("Aguardando thread..."); audio_thread.join(timeout=0.5)
        try: import pythoncom; pythoncom.CoUninitialize(); print("COM principal desinicializado.")
        except ImportError: pass
        except Exception as com_err: print(f"Aviso: Erro desinicializar COM: {com_err}")
        janela_a_fechar.destroy()
    janela.protocol("WM_DELETE_WINDOW", lambda: ao_fechar(janela))
    janela.mainloop()

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    if not is_windows or not pywin32_ok: print("Saindo."); sys.exit(1)
    try: print(f"Tesseract: {pytesseract.get_tesseract_version()}")
    except pytesseract.TesseractNotFoundError: print("AVISO: Tesseract não encontrado.")

    # A inicialização SAPI agora acontece dentro de criar_janela_preview
    # após a janela Tk ser criada, então não chamamos mais aqui.
    # if not inicializar_e_listar_sapi_voices():
    #      print("Falha ao inicializar SAPI. A leitura de voz não funcionará.")

    criar_janela_preview()

