import platform
import sys
import threading
import time
import tkinter as tk
import traceback  # Para depuração de erros em threads
from tkinter import scrolledtext, messagebox, ttk

import pyperclip
import pytesseract
from PIL import Image, ImageGrab

# --- Verificação e Importação pywin32 ---
is_windows = platform.system() == "Windows"
if not is_windows:
    print("ERRO: Requer Windows (pywin32).")
    # (Código de erro e saída omitido para brevidade)
    sys.exit(1)
try:
    import win32com.client as wincl
    import pythoncom # Necessário para COM em threads e eventos
    pythoncom.CoInitialize() # Inicializa COM para thread principal
    pywin32_ok = True
    print("pywin32 e COM inicializados.")
except ImportError:
    print("ERRO: pywin32 não instalado.")
    # (Código de erro e saída omitido para brevidade)
    sys.exit(1)
except Exception as e:
     print(f"ERRO importando/inicializando pywin32/COM: {e}")
     pywin32_ok = False


# --- Configuração Tesseract ---
# try:
#     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# except Exception as e:
#     print(f"Aviso: Não definiu caminho Tesseract: {e}")

# --- Variáveis Globais ---
audio_thread = None
stop_audio_flag = threading.Event()
sapi_voice_object = None # Objeto SAPI principal (para listar vozes)
available_sapi_voices = {} # Dicionário para vozes {descrição: índice}
# Variáveis Tkinter movidas para dentro de criar_janela_preview

# --- Global Refs for VoiceEvents ---
global_text_widget = None
global_root_window = None

# --- Classe de Eventos SAPI ---
class VoiceEvents:
    def __init__(self):
        self.last_highlight_range = None

    def OnWordBoundary(self, StreamNumber, StreamPosition, CharacterPosition, Length):
        global global_root_window
        # print(f"WordBoundary Evt: Pos={CharacterPosition}, Len={Length}") # Debug
        if global_root_window:
            global_root_window.after(0, self._update_highlight, CharacterPosition, Length)

    def OnEndStream(self, StreamNumber, StreamPosition):
        global global_root_window
        print("EndStream Event Received.")
        if global_root_window:
            global_root_window.after(0, self._remove_last_highlight)

    def _update_highlight(self, char_pos, length):
        global global_text_widget
        if not global_text_widget: return
        try:
            # print(f"  _update_highlight: Pos={char_pos}, Len={length}") # Debug
            global_text_widget.tag_remove("highlight", "1.0", tk.END)
            start_index = f"1.0 + {char_pos} chars"
            end_index = f"{start_index} + {length} chars"
            # print(f"  Highlighting: {start_index} to {end_index}") # Debug
            global_text_widget.tag_add("highlight", start_index, end_index)
            self.last_highlight_range = (start_index, end_index)
            global_text_widget.see(start_index) # Garante visibilidade
        except tk.TclError as e:
            print(f"Erro Tcl ao atualizar destaque (índice inválido?): {e}")
        except Exception as e:
            print(f"Erro inesperado ao atualizar destaque: {e}")
            # traceback.print_exc()

    def _remove_last_highlight(self):
         global global_text_widget
         if not global_text_widget: return
         try:
             print("Removendo último destaque via _remove_last_highlight.")
             global_text_widget.tag_remove("highlight", "1.0", tk.END)
             self.last_highlight_range = None
         except Exception as e:
             print(f"Erro ao remover último destaque: {e}")


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
def inicializar_e_listar_sapi_voices(voice_var_ref):
    """Inicializa SAPI, lista vozes e preenche o dicionário global."""
    global sapi_voice_object, available_sapi_voices
    if not pywin32_ok: return False

    if sapi_voice_object is None:
        try:
            print("Inicializando objeto SAPI SpVoice...")
            sapi_voice_object = wincl.Dispatch("SAPI.SpVoice")
            print("Objeto SAPI inicializado.")

            available_sapi_voices.clear()
            voices = sapi_voice_object.GetVoices()
            print("Vozes SAPI disponíveis:")
            default_voice_desc = sapi_voice_object.Voice.GetDescription()
            default_voice_index = 0
            for i, v in enumerate(voices):
                description = v.GetDescription()
                print(f"  {i}: {description}")
                available_sapi_voices[description] = i
                if description == default_voice_desc: default_voice_index = i

            voice_var_ref.set(default_voice_index)
            print(f"Voz padrão definida: {default_voice_desc} (Índice: {default_voice_index})")
            return True

        except Exception as e:
            print(f"ERRO CRÍTICO ao inicializar SAPI: {e}")
            messagebox.showerror("Erro SAPI", f"Não foi possível inicializar SAPI.\n{e}")
            sapi_voice_object = None; available_sapi_voices.clear()
            return False
    return True

def ler_texto_pywin32(texto, voice_idx, speed_str, janela_ref, botao_ler, botao_parar):
    """Usa SAPI para ler o texto com voz/velocidade e destaque."""
    global audio_thread, stop_audio_flag, sapi_voice_object, global_text_widget

    if not sapi_voice_object:
        messagebox.showerror("Erro SAPI", "Motor SAPI não inicializado.", parent=janela_ref)
        return
    if not global_text_widget:
         messagebox.showerror("Erro Interno", "Referência ao widget de texto não encontrada.", parent=janela_ref)
         return

    if audio_thread and audio_thread.is_alive():
        # Esta verificação agora só impede iniciar uma nova leitura se outra JÁ ESTIVER rodando
        # Não impede mais o clique durante a leitura (tratado em on_text_click)
        messagebox.showwarning("Leitura", "Leitura já em andamento.", parent=janela_ref)
        return

    stop_audio_flag.clear()
    # Limpa destaque ANTES de iniciar a thread
    global_text_widget.tag_remove("highlight", "1.0", tk.END)

    if not texto:
        messagebox.showinfo("Leitura", "Nenhum texto.", parent=janela_ref)
        return

    janela_ref.config(cursor="watch")
    botao_ler.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL)
    janela_ref.update_idletasks()

    def tarefa_leitura_sapi():
        """Função da thread para controlar a fala SAPI e eventos."""
        nonlocal texto, voice_idx, speed_str, janela_ref, botao_ler, botao_parar
        speak_obj_with_events = None
        com_initialized_thread = False
        event_handler_instance = None # Para manter referência ao handler
        try:
            try:
                 pythoncom.CoInitialize(); com_initialized_thread = True
                 print("COM inicializado para thread de leitura.")
            except Exception as com_err:
                 if hasattr(com_err, 'hresult') and com_err.hresult in [-2147417850, -2147221008]:
                      com_initialized_thread = True
                      print("COM já inicializado para thread de leitura.")
                 else: raise Exception(f"Falha ao inicializar COM na thread: {com_err}")

            speak_obj = wincl.Dispatch("SAPI.SpVoice")
            # Cria instância do handler ANTES de conectar
            event_handler_instance = VoiceEvents()
            speak_obj_with_events = wincl.WithEvents(speak_obj, VoiceEvents)
            # Guarda a referência ao handler criado internamente por WithEvents se necessário,
            # mas geralmente não precisamos interagir com ele diretamente.
            # A instância event_handler_instance é usada implicitamente.

            try:
                 voices = speak_obj.GetVoices()
                 if 0 <= voice_idx < voices.Count: speak_obj.Voice = voices.Item(voice_idx)
                 else: print(f"Aviso: Índice voz inválido ({voice_idx}).")
                 speed_map = {"Lenta": -1, "Normal": 2, "Rápida": 5}
                 rate_value = speed_map.get(speed_str, 0)
                 speak_obj.Rate = rate_value
                 print(f"Voz: {speak_obj.Voice.GetDescription()}, Velocidade: {rate_value}")
            except Exception as e_prop: print(f"Aviso: Erro definir voz/velocidade: {e_prop}")

            speak_obj.EventInterests = 1 + 8 # 1=EndStream, 8=WordBoundary

            SVSFlagsAsync = 1; SVSFPurgeBeforeSpeak = 2
            print(f"Falando com SAPI (com eventos): {texto[:50]}...")
            speak_obj.Speak(texto, SVSFlagsAsync | SVSFPurgeBeforeSpeak)

            time.sleep(0.08)

            # Loop de bombeamento de mensagens COM
            while speak_obj.Status.RunningState == 2:
                if stop_audio_flag.is_set():
                    print("Parada solicitada...")
                    speak_obj.Speak("", SVSFlagsAsync | SVSFPurgeBeforeSpeak) # Tenta parar
                    break
                # Processa eventos COM pendentes de forma mais robusta
                pythoncom.PumpWaitingMessages()
                time.sleep(0.05) # Pequena pausa, essencial para não travar

            # Garante processamento final de mensagens
            for _ in range(5): # Tenta bombear algumas vezes mais
                 pythoncom.PumpWaitingMessages()
                 time.sleep(0.02)

            if not stop_audio_flag.is_set(): print("Fala SAPI concluída.")
            else: print("Fala SAPI interrompida.")

        except Exception as e:
            print(f"Erro durante a fala SAPI com eventos: {e}")
            traceback.print_exc()
            janela_ref.after(0, lambda err=e: messagebox.showerror("Erro SAPI", f"Erro durante a fala:\n{err}", parent=janela_ref))
        finally:
            # Limpeza COM e GUI
            def restaurar_gui_thread_safe():
                print("Restaurando GUI após leitura/parada.") # Debug
                janela_ref.config(cursor="")
                botao_ler.config(state=tk.NORMAL)
                botao_parar.config(state=tk.DISABLED)
                # Tenta remover o destaque usando a instância do handler se existir
                if event_handler_instance:
                     event_handler_instance._remove_last_highlight()
                else: # Fallback tentando remover globalmente
                     try:
                          if global_text_widget: global_text_widget.tag_remove("highlight", "1.0", tk.END)
                     except Exception as e_clean: print(f"Erro fallback limpar destaque: {e_clean}")


            janela_ref.after(0, restaurar_gui_thread_safe)

            speak_obj_with_events = None # Libera referência
            speak_obj = None
            if com_initialized_thread:
                try: pythoncom.CoUninitialize(); print("COM da thread desinicializado.")
                except Exception as com_err: print(f"Aviso: Erro desinicializar COM thread: {com_err}")

    audio_thread = threading.Thread(target=tarefa_leitura_sapi); audio_thread.daemon = True; audio_thread.start()

# --- Função para Parar Leitura (SAPI) ---
def acao_parar_leitura():
    """Sinaliza para a thread de áudio SAPI parar a reprodução."""
    global stop_audio_flag, audio_thread
    if audio_thread and audio_thread.is_alive():
        print("Sinalizando parada SAPI...")
        stop_audio_flag.set()
        # Não tentamos mais parar diretamente aqui, a thread vai detectar a flag
    else:
        print("Nenhuma leitura SAPI ativa para parar.")

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
    global selected_voice_index, selected_speed, global_text_widget, global_root_window

    janela = tk.Tk(); janela.title("Leitor SAPI (Destaque + Clique)"); janela.geometry("700x600"); janela.attributes('-topmost', True)
    global_root_window = janela

    selected_voice_index = tk.IntVar(value=-1)
    selected_speed = tk.StringVar(value="Normal")

    if not inicializar_e_listar_sapi_voices(selected_voice_index):
         print("Falha inicializar SAPI.")

    frame_instrucoes = tk.Frame(janela, bd=1, relief=tk.SUNKEN); frame_instrucoes.pack(pady=5, padx=10, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=("Uso:\n- Cole texto ou use 'Buscar'/'Ler Imagem'.\n- Selecione Voz/Velocidade.\n- Clique 'Ler com SAPI'. O texto será destacado.\n- Clique em uma palavra para ler a partir dela (parará a leitura atual)."), justify=tk.LEFT); label_instrucoes.pack(pady=5, padx=5) # Instrução atualizada

    # --- Área de Texto Principal ---
    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=70, height=18, font=("Arial", 11));
    text_area.pack(pady=5, padx=10, expand=True, fill='both')
    text_area.tag_configure("highlight", background="yellow", foreground="black")
    global_text_widget = text_area

    # --- Frame para Seletores ---
    frame_selectors = tk.Frame(janela)
    frame_selectors.pack(pady=5, padx=10, fill=tk.X)

    # Seletor de Voz SAPI
    label_voice = tk.Label(frame_selectors, text="Voz SAPI:")
    label_voice.pack(side=tk.LEFT, padx=(0, 5))
    voice_options = list(available_sapi_voices.keys())
    voice_combo = ttk.Combobox(frame_selectors, values=voice_options, state='readonly', width=40)
    initial_voice_index = selected_voice_index.get()
    initial_voice_desc = "";
    for desc, idx in available_sapi_voices.items():
        if idx == initial_voice_index: initial_voice_desc = desc; break
    if initial_voice_desc: voice_combo.set(initial_voice_desc)
    elif voice_options: voice_combo.current(0)
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

    # --- Funções dos Botões e Eventos ---
    def acao_buscar_texto(): texto_cb = obter_texto_clipboard(); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_cb) if texto_cb else None
    def acao_ler_imagem():
        janela.config(cursor="watch"); botao_ler_imagem.config(state=tk.DISABLED); botao_ler_audio.config(state=tk.DISABLED); botao_parar_audio.config(state=tk.DISABLED)
        janela.update_idletasks()
        try:
            imagem_cb = obter_imagem_clipboard()
            if imagem_cb: texto_ocr = extrair_texto_de_imagem(imagem_cb); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_ocr) if texto_ocr else None
            else: messagebox.showinfo("OCR", "Nenhuma imagem.", parent=janela)
        finally: janela.config(cursor=""); botao_ler_imagem.config(state=tk.NORMAL); botao_ler_audio.config(state=tk.NORMAL)

    def acao_ler_texto_area_sapi(texto_override=None):
        global selected_voice_index, selected_speed, global_text_widget
        texto_para_ler = texto_override if texto_override is not None else global_text_widget.get("1.0", tk.END).strip()
        selected_voice_desc = voice_combo.get()
        voice_index_to_use = available_sapi_voices.get(selected_voice_desc, -1)
        speed_string_to_use = selected_speed.get()

        if voice_index_to_use == -1 and available_sapi_voices:
             print(f"Aviso: Descrição '{selected_voice_desc}' não encontrada. Usando padrão.")
             voice_index_to_use = selected_voice_index.get() if selected_voice_index.get() != -1 else 0

        ler_texto_pywin32(texto_para_ler, voice_index_to_use, speed_string_to_use, janela, botao_ler_audio, botao_parar_audio)

    # --- Evento de Clique no Texto (Lógica Simplificada) ---
    def on_text_click(event):
        """Para a leitura atual se estiver ativa, ou inicia a partir do clique."""
        global global_text_widget
        if not global_text_widget: return

        if audio_thread and audio_thread.is_alive():
            print("Clique durante leitura: Parando leitura atual.")
            acao_parar_leitura() # Apenas para a leitura atual
            # O usuário precisará clicar novamente para iniciar do novo ponto
        else:
            # Se não estava lendo, inicia a partir do clique
            print("Clique fora de leitura: Iniciando leitura a partir do ponto.")
            try:
                 click_index = global_text_widget.index(f"@{event.x},{event.y}")
                 word_start_index = global_text_widget.index(f"{click_index} wordstart")
                 print(f"Clique em: {click_index}, Início da palavra: {word_start_index}")
                 text_to_read = global_text_widget.get(word_start_index, tk.END)
                 if text_to_read.strip():
                     acao_ler_texto_area_sapi(texto_override=text_to_read)
                 else:
                     print("Nenhum texto para ler a partir do clique.")
            except tk.TclError as e: print(f"Erro obter índice clique: {e}")
            except Exception as e: print(f"Erro inesperado clique: {e}"); traceback.print_exc()

    # Associa o evento de clique (botão esquerdo) à função on_text_click
    text_area.bind("<Button-1>", on_text_click)

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

    criar_janela_preview()

