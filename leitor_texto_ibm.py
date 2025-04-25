import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, filedialog # Adicionado ttk, filedialog
import pyperclip
from ibm_watson import TextToSpeechV1
from ibm_cloud_sdk_core.authenticators import IAMAuthenticator
from ibm_cloud_sdk_core.api_exception import ApiException
import sounddevice as sd
import soundfile as sf
import numpy as np
import threading
import time
from PIL import Image, ImageGrab
import pytesseract
import io
import os
import tempfile # Ainda usado para salvar temporariamente antes de mover/renomear
import platform
import sys
import datetime # Adicionado para nomes de arquivo
import re # Adicionado para limpar nomes de arquivo
import html # Adicionado para escapar texto para SSML

# --- Constantes ---
AUDIO_SAVE_DIR = "audio_salvos" # Pasta para salvar os áudios

# --- Cliente IBM Watson TTS ---
ibm_tts_client = None
ibm_credentials_ok = False
available_voices = {} # Dicionário para guardar vozes {display_name: voice_id}

try:
    IBM_API_KEY = 'IBM_KEY'
    IBM_TTS_URL = 'IBM_SERVICE_URL'
    if not IBM_API_KEY or not IBM_TTS_URL:
        print("ERRO: Variáveis IBM_API_KEY/IBM_TTS_URL não definidas.")
    else:
        authenticator = IAMAuthenticator(IBM_API_KEY)
        ibm_tts_client = TextToSpeechV1(authenticator=authenticator)
        ibm_tts_client.set_service_url(IBM_TTS_URL)
        ibm_credentials_ok = True
        print("Cliente IBM Watson TTS inicializado.")

        # Tenta buscar as vozes disponíveis
        try:
            print("Buscando vozes IBM Watson...")
            voices_result = ibm_tts_client.list_voices().get_result()
            # Filtra e armazena vozes (ex: prioriza pt-BR)
            for voice in voices_result.get('voices', []):
                lang = voice.get('language', '')
                desc = voice.get('description', '')
                name = voice.get('name', '')
                gender = voice.get('gender', '')
                # Cria um nome de exibição mais amigável
                display_name = f"{lang} - {gender.capitalize()} ({desc.split(' ')[0]})" # Ex: pt-BR - Female (Isabela)
                if name: # Só adiciona se tiver um nome (ID)
                     available_voices[display_name] = name
            print(f"Vozes encontradas: {list(available_voices.keys())}")
            if not available_voices:
                 print("Aviso: Nenhuma voz encontrada via API. Usando fallback.")
                 # Fallback se a busca falhar
                 available_voices = {
                      "pt-BR - Female (IsabelaV3)": "pt-BR_IsabelaV3Voice",
                      "pt-BR - Male (DanielV3)": "pt-BR_DanielV3Voice",
                      # Adicione outras vozes conhecidas se necessário
                 }
        except ApiException as api_ex:
             print(f"Erro ao buscar vozes IBM: {api_ex}")
             ibm_credentials_ok = False # Considera falha se não puder listar vozes
        except Exception as e_voice:
             print(f"Erro inesperado ao buscar vozes: {e_voice}")
             # Poderia usar fallback aqui também

except Exception as e:
    print(f"ERRO CRÍTICO ao inicializar IBM Watson TTS: {e}")
    ibm_credentials_ok = False

# --- Configuração Tesseract ---
# try:
#     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # Ajuste se necessário
# except Exception as e:
#     print(f"Aviso: Não definiu caminho Tesseract: {e}")

# --- Variáveis Globais ---
audio_thread = None
stop_audio_flag = threading.Event()
# audio_temp_file = None # Não mais necessário globalmente da mesma forma
monitorar_clipboard = False
thread_monitoramento = None
ultimo_texto_clipboard = ""
saved_audio_list = [] # Lista para manter nomes de arquivos salvos

# --- Funções Auxiliares ---
def create_audio_save_dir():
    """Cria o diretório para salvar áudios se não existir."""
    if not os.path.exists(AUDIO_SAVE_DIR):
        try:
            os.makedirs(AUDIO_SAVE_DIR)
            print(f"Diretório '{AUDIO_SAVE_DIR}' criado.")
        except OSError as e:
            messagebox.showerror("Erro de Diretório", f"Não foi possível criar o diretório '{AUDIO_SAVE_DIR}':\n{e}")
            return False
    return True

def sanitize_filename(text, max_len=50):
    """Cria um nome de arquivo seguro a partir do texto."""
    # Remove caracteres inválidos para nomes de arquivo
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    # Substitui espaços por underscores
    text = text.replace(" ", "_")
    # Trunca para o tamanho máximo
    return text[:max_len]

def escape_ssml(text):
    """Escapa caracteres especiais para uso seguro dentro de SSML."""
    # Usa html.escape para escapar &, <, >
    # Embora não usemos mais SSML complexo, escapar ainda é uma boa prática.
    return html.escape(text, quote=True)

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

# --- Função TTS IBM Watson (Modificada) ---
def ler_texto_ibm_watson(texto, voice_id, speed_rate, listbox_widget, janela_ref, botao_ler, botao_parar):
    """Gera fala com IBM, salva e toca, usando voz selecionada (velocidade ignorada)."""
    global ibm_credentials_ok, ibm_tts_client, audio_thread, stop_audio_flag, saved_audio_list

    if not ibm_credentials_ok: messagebox.showerror("Erro IBM", "Cliente IBM não inicializado.", parent=janela_ref); return
    if audio_thread and audio_thread.is_alive(): messagebox.showwarning("Leitura", "Leitura em andamento.", parent=janela_ref); return
    stop_audio_flag.clear()
    if not texto: messagebox.showinfo("Leitura", "Nenhum texto.", parent=janela_ref); return
    if not voice_id: messagebox.showwarning("Seleção", "Nenhuma voz selecionada.", parent=janela_ref); return

    # Garante que o diretório de salvamento exista
    if not create_audio_save_dir(): return

    janela_ref.config(cursor="watch"); botao_ler.config(state=tk.DISABLED); botao_parar.config(state=tk.NORMAL)
    janela_ref.update_idletasks()

    def tarefa_leitura_ibm_watson():
        nonlocal texto, voice_id, speed_rate, listbox_widget, janela_ref, botao_ler, botao_parar
        audio_data = None; samplerate = None; saved_filepath = None

        try:
            print(f"Gerando áudio IBM: Voz={voice_id}, Texto={texto[:30]}...") # Removido log de velocidade

            # *** MODIFICAÇÃO AQUI: Remover uso de SSML para velocidade ***
            # Prepara o texto (apenas escapa caracteres básicos)
            # escaped_text = escape_ssml(texto) # Escapar ainda é bom
            # ssml_text = f'<speak><prosody rate="{ssml_rate}">{escaped_text}</prosody></speak>' # Linha removida
            # print(f"Texto SSML: {ssml_text[:100]}...") # Log removido

            # 1. Chamar a API IBM Watson TTS com texto simples
            # Nota: A velocidade selecionada (speed_rate) é ignorada agora.
            response = ibm_tts_client.synthesize(
                text=texto, # Envia o texto original (ou escapado se preferir)
                voice=voice_id,
                accept='audio/mp3'
            ).get_result()
            # *** FIM DA MODIFICAÇÃO ***

            # 2. Salvar o áudio permanentemente
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_filename = sanitize_filename(texto)
            filename = f"{timestamp}_{base_filename}.mp3"
            saved_filepath = os.path.join(AUDIO_SAVE_DIR, filename)

            with open(saved_filepath, 'wb') as f: f.write(response.content)
            print(f"Áudio IBM salvo permanentemente em: {saved_filepath}")

            # Adiciona à lista da GUI e à lista interna (na thread principal)
            def update_listbox():
                 listbox_widget.insert(tk.END, filename)
                 saved_audio_list.append(saved_filepath)
            janela_ref.after(0, update_listbox)


            if stop_audio_flag.is_set(): print("Parada antes de tocar."); return

            # 3. Ler Arquivo de Áudio Salvo
            print("Lendo arquivo salvo..."); audio_data, samplerate = sf.read(saved_filepath, dtype='float32')
            print(f"Áudio lido: {len(audio_data)} amostras, Taxa: {samplerate} Hz")

            # 4. Tocar Áudio com SoundDevice
            print("Reproduzindo..."); sd.play(audio_data, samplerate)
            while sd.get_stream().active:
                if stop_audio_flag.is_set(): sd.stop(); print("Reprodução interrompida."); break
                time.sleep(0.1)
            if not stop_audio_flag.is_set(): print("Reprodução IBM concluída.")

        except ApiException as ex: # Erro específico da API IBM
            print(f"Erro API IBM: {ex.code}, {ex.message}")
            # Captura o valor de 'ex' usando um argumento padrão na lambda
            janela_ref.after(0, lambda api_err=ex: messagebox.showerror("Erro API IBM", f"Erro API Watson:\n{api_err.message} (Código: {api_err.code})", parent=janela_ref))
        except sf.LibsndfileError as e: print(f"Erro SoundFile: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro Áudio", f"Erro ler áudio (libsndfile?):\n{err}", parent=janela_ref))
        except sd.PortAudioError as e: print(f"Erro SoundDevice: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro Áudio", f"Erro dispositivo áudio:\n{err}", parent=janela_ref))
        except Exception as e: print(f"Erro thread IBM: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro TTS", f"Erro inesperado:\n{err}", parent=janela_ref))
        finally:
            def restaurar_gui_thread_safe(): janela_ref.config(cursor=""); botao_ler.config(state=tk.NORMAL if ibm_credentials_ok else tk.DISABLED); botao_parar.config(state=tk.DISABLED)
            janela_ref.after(0, restaurar_gui_thread_safe)
            # Não removemos mais o arquivo, pois foi salvo permanentemente

    audio_thread = threading.Thread(target=tarefa_leitura_ibm_watson); audio_thread.daemon = True; audio_thread.start()

# --- Função para Tocar Áudio Salvo da Lista ---
def play_saved_audio(listbox_widget, janela_ref, botao_play_saved, botao_parar):
    """Toca o arquivo de áudio selecionado na lista."""
    global audio_thread, stop_audio_flag

    selected_indices = listbox_widget.curselection()
    if not selected_indices:
        messagebox.showinfo("Seleção", "Nenhum áudio selecionado na lista.", parent=janela_ref)
        return

    if audio_thread and audio_thread.is_alive():
        messagebox.showwarning("Leitura", "Outra leitura já está em andamento.", parent=janela_ref)
        return

    stop_audio_flag.clear()
    selected_filename = listbox_widget.get(selected_indices[0])
    filepath_to_play = os.path.join(AUDIO_SAVE_DIR, selected_filename)

    if not os.path.exists(filepath_to_play):
         messagebox.showerror("Erro Arquivo", f"Arquivo não encontrado:\n{filepath_to_play}", parent=janela_ref)
         return

    janela_ref.config(cursor="watch")
    botao_play_saved.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL)
    janela_ref.update_idletasks()

    def task_play_saved():
        """Thread para tocar o áudio salvo."""
        nonlocal filepath_to_play, janela_ref, botao_play_saved, botao_parar
        try:
            print(f"Lendo arquivo salvo: {filepath_to_play}...")
            audio_data, samplerate = sf.read(filepath_to_play, dtype='float32')
            print(f"Áudio lido: {len(audio_data)} amostras, Taxa: {samplerate} Hz")

            print("Reproduzindo áudio salvo..."); sd.play(audio_data, samplerate)
            while sd.get_stream().active:
                if stop_audio_flag.is_set(): sd.stop(); print("Reprodução interrompida."); break
                time.sleep(0.1)
            if not stop_audio_flag.is_set(): print("Reprodução do arquivo salvo concluída.")

        except sf.LibsndfileError as e: print(f"Erro SoundFile: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro Áudio", f"Erro ler áudio:\n{err}", parent=janela_ref))
        except sd.PortAudioError as e: print(f"Erro SoundDevice: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro Áudio", f"Erro dispositivo áudio:\n{err}", parent=janela_ref))
        except Exception as e: print(f"Erro thread play_saved: {e}"); janela_ref.after(0, lambda err=e: messagebox.showerror("Erro", f"Erro inesperado:\n{err}", parent=janela_ref))
        finally:
            def restore_gui(): janela_ref.config(cursor=""); botao_play_saved.config(state=tk.NORMAL); botao_parar.config(state=tk.DISABLED)
            janela_ref.after(0, restore_gui)

    audio_thread = threading.Thread(target=task_play_saved); audio_thread.daemon = True; audio_thread.start()

# --- Função para Parar Leitura ---
def acao_parar_leitura():
    global stop_audio_flag, audio_thread
    if audio_thread and audio_thread.is_alive(): print("Sinalizando parada..."); stop_audio_flag.set()
    else: print("Nenhuma leitura ativa.")

# --- Funções de Monitoramento ---
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

# --- Interface Gráfica Principal (Modificada) ---
def criar_janela_preview():
    janela = tk.Tk(); janela.title("Leitor de Tela Avançado (IBM TTS)"); janela.geometry("850x600") # Aumenta tamanho
    janela.attributes('-topmost', True)

    # --- Frames Principais (Esquerda para Texto/Controles, Direita para Lista) ---
    left_frame = tk.Frame(janela)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

    right_frame = tk.Frame(janela, bd=1, relief=tk.SUNKEN)
    right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)

    # --- Widgets do Painel Esquerdo ---
    frame_instrucoes = tk.Frame(left_frame, bd=1, relief=tk.SUNKEN); frame_instrucoes.pack(pady=5, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=("Uso:\n- Copie texto/imagem e use 'Buscar'/'Ler Imagem'.\n- Selecione Voz/Velocidade (Velocidade pode não funcionar em todas as vozes).\n- Clique 'Ler com IBM' para ouvir e salvar.\n- Use a lista à direita para tocar áudios salvos."), justify=tk.LEFT); label_instrucoes.pack(pady=5, padx=5) # Nota sobre velocidade adicionada

    text_area = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD, width=70, height=18, font=("Arial", 11)); text_area.pack(pady=5, expand=True, fill='both')

    # Frame para seletores (Voz e Velocidade)
    frame_selectors = tk.Frame(left_frame)
    frame_selectors.pack(pady=5, fill=tk.X)

    # Seletor de Voz
    label_voice = tk.Label(frame_selectors, text="Voz:")
    label_voice.pack(side=tk.LEFT, padx=(0, 5))
    voice_var = tk.StringVar()
    voice_options = list(available_voices.keys())
    voice_combo = ttk.Combobox(frame_selectors, textvariable=voice_var, values=voice_options, state='readonly', width=35)
    if voice_options: voice_var.set(voice_options[0]) # Define padrão se houver vozes
    voice_combo.pack(side=tk.LEFT, padx=5)
    if not ibm_credentials_ok or not available_voices: voice_combo.config(state=tk.DISABLED) # Desabilita se IBM falhou ou sem vozes

    # Seletor de Velocidade
    label_speed = tk.Label(frame_selectors, text="Velocidade:")
    label_speed.pack(side=tk.LEFT, padx=(10, 5))
    speed_var = tk.StringVar(value="Normal") # Padrão
    speed_options = ["Lenta", "Normal", "Rápida"]
    speed_combo = ttk.Combobox(frame_selectors, textvariable=speed_var, values=speed_options, state='readonly', width=10)
    speed_combo.pack(side=tk.LEFT, padx=5)
    if not ibm_credentials_ok: speed_combo.config(state=tk.DISABLED)
    # Nota: Este seletor agora pode não ter efeito dependendo da voz selecionada.

    # Frames para botões de ação e controle
    frame_botoes_acao = tk.Frame(left_frame); frame_botoes_acao.pack(pady=(5, 0))
    frame_botoes_controle = tk.Frame(left_frame); frame_botoes_controle.pack(pady=(0, 10))

    # --- Widgets do Painel Direito (Lista de Áudios) ---
    label_saved_list = tk.Label(right_frame, text="Áudios Salvos:")
    label_saved_list.pack(pady=(5,0))

    listbox_frame = tk.Frame(right_frame)
    listbox_frame.pack(fill=tk.BOTH, expand=True)

    scrollbar_y = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
    saved_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar_y.set, width=35, height=20)
    scrollbar_y.config(command=saved_listbox.yview)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    saved_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Botão para tocar áudio selecionado da lista
    botao_play_saved = tk.Button(right_frame, text="Tocar Selecionado", width=20)
    botao_play_saved.pack(pady=5)
    # O comando será definido depois, junto com o botão de parar principal

    # --- Botões (Declarados antes para referência) ---
    botao_ler_imagem = tk.Button(frame_botoes_acao, text="Ler Imagem", width=15, height=2)
    botao_ler_audio = tk.Button(frame_botoes_acao, text="Ler com IBM", width=15, height=2, fg="green")
    botao_parar_audio = tk.Button(frame_botoes_acao, text="Parar Leitura", command=acao_parar_leitura, width=12, height=2, state=tk.DISABLED)

    # Configura comando do botão Tocar Selecionado
    botao_play_saved.config(command=lambda: play_saved_audio(saved_listbox, janela, botao_play_saved, botao_parar_audio))

    # --- Funções dos Botões ---
    def acao_buscar_texto(): texto_cb = obter_texto_clipboard(); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_cb) if texto_cb else None
    def acao_ler_imagem():
        janela.config(cursor="watch"); botao_ler_imagem.config(state=tk.DISABLED); botao_ler_audio.config(state=tk.DISABLED); botao_parar_audio.config(state=tk.DISABLED); botao_play_saved.config(state=tk.DISABLED)
        janela.update_idletasks()
        try:
            imagem_cb = obter_imagem_clipboard()
            if imagem_cb: texto_ocr = extrair_texto_de_imagem(imagem_cb); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_ocr) if texto_ocr else None
            else: messagebox.showinfo("OCR", "Nenhuma imagem.", parent=janela)
        finally: janela.config(cursor=""); botao_ler_imagem.config(state=tk.NORMAL); botao_ler_audio.config(state=tk.NORMAL if ibm_credentials_ok else tk.DISABLED); botao_play_saved.config(state=tk.NORMAL)

    def acao_ler_texto_area_ibm():
        texto_para_ler = text_area.get("1.0", tk.END).strip()
        selected_voice_display = voice_var.get()
        selected_voice_id = available_voices.get(selected_voice_display) # Pega o ID real da voz
        selected_speed = speed_var.get() # Pega a velocidade selecionada (mas pode ser ignorada pela função TTS)
        if not selected_voice_id: messagebox.showwarning("Seleção", "Voz inválida selecionada.", parent=janela); return
        # Passa a listbox para a função de leitura adicionar o item
        ler_texto_ibm_watson(texto_para_ler, selected_voice_id, selected_speed, saved_listbox, janela, botao_ler_audio, botao_parar_audio)

    # --- Configuração final e Empacotamento dos Botões ---
    botao_buscar_texto = tk.Button(frame_botoes_acao, text="Buscar Texto", command=acao_buscar_texto, width=15, height=2); botao_buscar_texto.pack(side=tk.LEFT, padx=5, pady=5)
    botao_ler_imagem.config(command=acao_ler_imagem); botao_ler_imagem.pack(side=tk.LEFT, padx=5, pady=5)
    botao_ler_audio.config(command=acao_ler_texto_area_ibm); botao_ler_audio.pack(side=tk.LEFT, padx=5, pady=5)
    if not ibm_credentials_ok: botao_ler_audio.config(state=tk.DISABLED, text="Ler (Erro Cred.)")
    botao_parar_audio.pack(side=tk.LEFT, padx=5, pady=5)

    botao_monitor = tk.Button(frame_botoes_controle, text="Iniciar Monitoramento", width=20, height=2); botao_monitor.config(command=lambda: iniciar_parar_monitoramento(text_area, janela, botao_monitor)); botao_monitor.pack(side=tk.LEFT, padx=5, pady=5)
    botao_fechar = tk.Button(frame_botoes_controle, text="Fechar", command=lambda: ao_fechar(janela), width=10, height=2); botao_fechar.pack(side=tk.LEFT, padx=5, pady=5)

    # --- Carregar lista de áudios salvos existentes ---
    def load_saved_audio_list():
        global saved_audio_list
        saved_audio_list = []
        if create_audio_save_dir() and os.path.exists(AUDIO_SAVE_DIR):
            try:
                files = sorted(
                    [f for f in os.listdir(AUDIO_SAVE_DIR) if f.lower().endswith(".mp3")],
                    key=lambda f: os.path.getmtime(os.path.join(AUDIO_SAVE_DIR, f)),
                    reverse=True
                )
                saved_listbox.delete(0, tk.END)
                for filename in files:
                    saved_listbox.insert(tk.END, filename)
                    saved_audio_list.append(os.path.join(AUDIO_SAVE_DIR, filename))
                print(f"{len(saved_audio_list)} áudios salvos carregados.")
            except Exception as e:
                 print(f"Erro ao carregar lista: {e}")
                 messagebox.showerror("Erro Lista Áudio", f"Não carregou áudios salvos:\n{e}")

    load_saved_audio_list() # Carrega a lista ao iniciar

    # --- Tratamento ao Fechar Janela ---
    def ao_fechar(janela_a_fechar):
        global monitorar_clipboard, stop_audio_flag, audio_thread
        print("Fechando..."); monitorar_clipboard = False; stop_audio_flag.set()
        try: sd.stop()
        except Exception as e: print(f"Info: Erro sd.stop() ao fechar: {e}")
        if audio_thread and audio_thread.is_alive(): print("Aguardando thread..."); audio_thread.join(timeout=0.5)
        janela_a_fechar.destroy()
    janela.protocol("WM_DELETE_WINDOW", lambda: ao_fechar(janela))

    if not ibm_credentials_ok: messagebox.showwarning("Aviso IBM", "Não conectou à API IBM Watson TTS...", parent=janela)
    janela.mainloop()

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    try: print(f"Tesseract: {pytesseract.get_tesseract_version()}")
    except pytesseract.TesseractNotFoundError: print("AVISO: Tesseract não encontrado.")
    try: print(f"Dispositivo áudio: {sd.query_devices(kind='output')}")
    except Exception as e: print(f"Aviso: Não consultou dispositivos áudio: {e}")
    if not ibm_credentials_ok: print("AVISO: Credenciais IBM não carregadas.")
    create_audio_save_dir()
    criar_janela_preview()
