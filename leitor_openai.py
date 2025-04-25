import tkinter as tk
from tkinter import scrolledtext, messagebox
import pyperclip
from openai import OpenAI
# from playsound import playsound # Removido
import sounddevice as sd # Adicionado
import soundfile as sf # Adicionado
import numpy as np # Adicionado (dependência comum)
import threading
import time
from PIL import Image, ImageGrab
import pytesseract
import io
import os
import tempfile

# --- Configuração Inicial ---
# [Windows Apenas] Tesseract Path (se necessário)
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\orlando.santos\AppData\Local\Programs\Tesseract-OCR'
except Exception as e:
    print(f"Aviso: Não foi possível definir o caminho do Tesseract: {e}")

# --- Cliente OpenAI ---
client = None
openai_client_ok = False
try:
    client = OpenAI(api_key='OPEN_API_KEY') # Lê OPENAI_API_KEY do ambiente
    openai_client_ok = True
    print("Cliente OpenAI inicializado com sucesso.")
except Exception as e:
    print(f"ERRO CRÍTICO: Não foi possível inicializar o cliente OpenAI: {e}.")
    print("Verifique se a variável de ambiente OPENAI_API_KEY está configurada corretamente.")
    # É importante notificar o usuário na GUI também

# --- Variáveis Globais de Controle de Áudio ---
audio_thread = None
stop_audio_flag = threading.Event() # Evento para sinalizar parada

# --- Funções Clipboard e OCR ---
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
            return None
    except Exception as e:
        print(f"Erro ao obter imagem do clipboard: {e}")
        return None

def extrair_texto_de_imagem(imagem):
    if not isinstance(imagem, Image.Image):
        messagebox.showinfo("OCR", "Nenhuma imagem válida encontrada na área de transferência.")
        return ""
    try:
        lang = 'por' # Fixo por enquanto
        print(f"Executando OCR com idioma: {lang}")
        texto_extraido = pytesseract.image_to_string(imagem, lang=lang)
        if not texto_extraido.strip():
             messagebox.showinfo("OCR", "Não foi possível extrair texto da imagem.")
             return ""
        print("Texto extraído da imagem com sucesso.")
        return texto_extraido.strip()
    except PermissionError as pe:  # Captura especificamente o erro de permissão
        print(f"ERRO DE PERMISSÃO ao chamar Tesseract: {pe}")
        messagebox.showerror("Erro de Permissão OCR",
                             f"Acesso negado ao tentar executar o Tesseract.\n"
                             f"Causas comuns:\n"
                             f"- Tesseract não tem permissão de execução.\n"
                             f"- Antivírus bloqueando.\n"
                             f"- Permissão negada na pasta temporária.\n"
                             f"Verifique o caminho configurado (se houver) e as permissões.\n\n"
                             f"Erro detalhado: {pe}")
        return ""  # Retorna vazio em caso de erro
    except pytesseract.TesseractNotFoundError:
        messagebox.showerror("Erro OCR", "Tesseract não encontrado. Verifique a instalação e a configuração do PATH.")
        return ""
    except Exception as e:
        messagebox.showerror("Erro OCR", f"Ocorreu um erro durante o OCR:\n{e}")
        return ""

# --- Função TTS OpenAI com SoundDevice ---
audio_temp_file = None

def ler_texto_em_voz_alta_openai(texto, janela_ref, botao_ler, botao_parar):
    """Usa a API OpenAI TTS, SoundFile e SoundDevice para gerar e tocar a fala."""
    global audio_temp_file, openai_client_ok, client, audio_thread, stop_audio_flag

    # Verifica se já existe uma leitura em andamento
    if audio_thread and audio_thread.is_alive():
        messagebox.showwarning("Leitura", "Uma leitura já está em andamento.", parent=janela_ref)
        return

    # Limpa a flag de parada
    stop_audio_flag.clear()

    if not openai_client_ok:
         try:
             print("Tentando reinicializar cliente OpenAI...")
             client = OpenAI()
             openai_client_ok = True
             print("Cliente OpenAI reinicializado com sucesso.")
         except Exception as e:
              messagebox.showerror("Erro OpenAI", f"Cliente OpenAI não inicializado. Verifique a API Key.\n{e}", parent=janela_ref)
              return

    if not texto:
        messagebox.showinfo("Leitura", "Nenhum texto fornecido para leitura.", parent=janela_ref)
        return

    # Limpa arquivo temporário antigo
    if audio_temp_file and os.path.exists(audio_temp_file):
        try: os.remove(audio_temp_file)
        except Exception as e: print(f"Aviso: Não removeu temp anterior: {e}")
        audio_temp_file = None

    # Feedback visual e controle de botões
    janela_ref.config(cursor="watch")
    botao_ler.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL) # Habilita o botão de parar
    # Desabilitar outros botões se necessário...
    janela_ref.update_idletasks()

    def tarefa_leitura_openai_sounddevice():
        global audio_temp_file
        nonlocal texto, janela_ref, botao_ler, botao_parar # Captura variáveis
        audio_data = None
        samplerate = None
        audio_gerado_com_sucesso = False

        try:
            print(f"Gerando áudio com OpenAI para: {texto[:50]}...")
            response = client.audio.speech.create(
                model="gpt-4o-mini-tts", voice="echo", input=texto, response_format="mp3" # Garante MP3
            )

            # Salva o conteúdo MP3 em um arquivo temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as f:
                 f.write(response.content)
                 audio_temp_file = f.name
                 audio_gerado_com_sucesso = True
                 print(f"Áudio temporário salvo em: {audio_temp_file}")

            # Lê o arquivo de áudio usando soundfile
            print("Lendo arquivo de áudio...")
            # dtype='float32' é comum para sounddevice
            audio_data, samplerate = sf.read(audio_temp_file, dtype='float32')
            print(f"Áudio lido: {len(audio_data)} amostras, Taxa: {samplerate} Hz")

            # Toca o áudio usando sounddevice
            print("Reproduzindo áudio com SoundDevice...")
            # sd.play() não bloqueia por padrão, mas usamos wait()
            sd.play(audio_data, samplerate)

            # Espera a reprodução terminar OU a flag de parada ser setada
            # sd.wait() bloquearia, então usamos um loop com sleep e verificamos a flag
            while sd.get_stream().active:
                if stop_audio_flag.is_set():
                    sd.stop()
                    print("Reprodução interrompida pelo usuário.")
                    break
                time.sleep(0.1) # Pequena pausa para não consumir CPU excessivamente

            if not stop_audio_flag.is_set():
                 print("Reprodução concluída.")


        except sf.LibsndfileError as e:
             print(f"Erro ao ler arquivo de áudio com SoundFile: {e}")
             janela_ref.after(0, lambda: messagebox.showerror("Erro de Áudio", f"Não foi possível ler o formato do áudio:\n{e}\nVerifique se 'libsndfile' está instalado.", parent=janela_ref))
        except sd.PortAudioError as e:
             print(f"Erro do SoundDevice/PortAudio: {e}")
             janela_ref.after(0, lambda: messagebox.showerror("Erro de Áudio", f"Erro no dispositivo de áudio:\n{e}", parent=janela_ref))
        except Exception as e:
            print(f"Erro durante a geração ou reprodução do áudio: {e}")
            msg_erro = f"Ocorreu um erro:\n{e}"
            if "OPENAI_API_KEY" in str(e):
                 msg_erro = "Erro de Autenticação OpenAI.\nVerifique sua API Key."
            # Garante que a mensagem de erro seja exibida na thread principal
            janela_ref.after(0, lambda: messagebox.showerror("Erro OpenAI TTS", msg_erro, parent=janela_ref))
        finally:
            # Limpeza e restauração da GUI (sempre executado)
            # Garante que a restauração ocorra na thread principal da GUI
            def restaurar_gui():
                janela_ref.config(cursor="")
                botao_ler.config(state=tk.NORMAL)
                botao_parar.config(state=tk.DISABLED) # Desabilita parar após terminar/erro
                # Reabilitar outros botões...

            janela_ref.after(0, restaurar_gui)

            # Limpa o arquivo temporário APÓS tocar ou falhar
            if audio_gerado_com_sucesso and audio_temp_file and os.path.exists(audio_temp_file):
                try:
                    # Pequena pausa antes de remover, especialmente no Windows
                    time.sleep(0.2)
                    os.remove(audio_temp_file)
                    print(f"Arquivo temporário removido: {audio_temp_file}")
                    audio_temp_file = None
                except PermissionError:
                     print(f"Aviso: Permissão negada ao remover {audio_temp_file}. O arquivo pode ainda estar em uso.")
                except Exception as e:
                    print(f"Aviso: Não foi possível remover o arquivo temporário: {e}")

    # Inicia a thread de leitura
    stop_audio_flag.clear() # Garante que a flag esteja limpa antes de iniciar
    audio_thread = threading.Thread(target=tarefa_leitura_openai_sounddevice)
    audio_thread.daemon = True
    audio_thread.start()

# --- Função para Parar Leitura (SoundDevice) ---
def acao_parar_leitura():
    """Sinaliza para a thread de áudio parar a reprodução."""
    global stop_audio_flag, audio_thread
    if audio_thread and audio_thread.is_alive():
        print("Sinalizando para parar a leitura...")
        stop_audio_flag.set() # Define o evento, a thread de leitura verificará isso
        # sd.stop() # Chamada direta aqui pode causar problemas se feita de outra thread
                  # É melhor deixar a própria thread de leitura chamar sd.stop()
    else:
        print("Nenhuma leitura em andamento para parar.")


# --- Funções de Monitoramento ---
# (Código das funções tarefa_monitoramento_clipboard, atualizar_texto_area_monitor, iniciar_parar_monitoramento inalterado)
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
             pass
        time.sleep(1)
    print("Monitoramento parado.")

def atualizar_texto_area_monitor(widget_texto, novo_texto):
    widget_texto.delete("1.0", tk.END)
    widget_texto.insert(tk.INSERT, novo_texto)

def iniciar_parar_monitoramento(text_area_widget, janela_principal, botao_monitor):
    global monitorar_clipboard, thread_monitoramento, ultimo_texto_clipboard
    if monitorar_clipboard:
        monitorar_clipboard = False
        botao_monitor.config(text="Iniciar Monitoramento", bg="SystemButtonFace", fg="SystemButtonText") # Feedback visual
    else:
        monitorar_clipboard = True
        botao_monitor.config(text="Parar Monitoramento", bg="red", fg="white") # Feedback visual
        ultimo_texto_clipboard = ""
        thread_monitoramento = threading.Thread(target=tarefa_monitoramento_clipboard, args=(text_area_widget, janela_principal), daemon=True)
        thread_monitoramento.start()


# --- Interface Gráfica Principal ---
def criar_janela_preview():
    janela = tk.Tk()
    janela.title("Leitor de Tela (OCR + OpenAI TTS + SoundDevice)") # Título atualizado
    janela.geometry("650x550") # Aumentar um pouco para o novo botão
    janela.attributes('-topmost', True)

    # --- Widgets ---
    frame_instrucoes = tk.Frame(janela, bd=1, relief=tk.SUNKEN)
    frame_instrucoes.pack(pady=5, padx=10, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=(
        "Uso:\n"
        "- Copie texto (Ctrl+C) e clique 'Buscar Texto' OU 'Iniciar Monitoramento'.\n"
        "- Copie uma imagem (PrintScreen/Captura) e clique 'Ler de Imagem'.\n"
        "- Edite o texto na caixa e clique 'Ler com OpenAI'."
    ), justify=tk.LEFT)
    label_instrucoes.pack(pady=5, padx=5)

    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=70, height=18, font=("Arial", 11))
    text_area.pack(pady=5, padx=10, expand=True, fill='both')

    frame_botoes_acao = tk.Frame(janela)
    frame_botoes_acao.pack(pady=(5, 0))
    frame_botoes_controle = tk.Frame(janela)
    frame_botoes_controle.pack(pady=(0, 10))

    # --- Botões (Declarados antes para referência na função de leitura) ---
    botao_ler_audio = tk.Button(frame_botoes_acao, text="Ler com OpenAI", width=20, height=2, fg="purple")
    botao_parar_audio = tk.Button(frame_botoes_acao, text="Parar Leitura", command=acao_parar_leitura, width=15, height=2, state=tk.DISABLED) # Começa desabilitado

    # --- Funções dos Botões ---
    def acao_buscar_texto():
        texto_cb = obter_texto_clipboard()
        if texto_cb:
            text_area.delete("1.0", tk.END)
            text_area.insert(tk.INSERT, texto_cb)

    def acao_ler_imagem():
        janela.config(cursor="watch")
        # Desabilitar botões relevantes...
        botao_ler_imagem.config(state=tk.DISABLED)
        botao_ler_audio.config(state=tk.DISABLED)
        botao_parar_audio.config(state=tk.DISABLED)

        janela.update_idletasks()
        try:
            imagem_cb = obter_imagem_clipboard()
            if imagem_cb:
                texto_ocr = extrair_texto_de_imagem(imagem_cb)
                if texto_ocr:
                    text_area.delete("1.0", tk.END)
                    text_area.insert(tk.INSERT, texto_ocr)
            else:
                 messagebox.showinfo("OCR", "Nenhuma imagem encontrada na área de transferência.", parent=janela)
        finally:
             janela.config(cursor="")
             # Reabilitar botões
             botao_ler_imagem.config(state=tk.NORMAL)
             if openai_client_ok: # Só reabilita leitura se OpenAI estiver ok
                 botao_ler_audio.config(state=tk.NORMAL)
             # Botão parar continua desabilitado até iniciar leitura

    def acao_ler_texto_area_openai():
        texto_para_ler = text_area.get("1.0", tk.END).strip()
        # Passa referências dos botões para controle de estado
        ler_texto_em_voz_alta_openai(texto_para_ler, janela, botao_ler_audio, botao_parar_audio)

    # --- Configuração final e Empacotamento dos Botões ---
    botao_buscar_texto = tk.Button(frame_botoes_acao, text="Buscar Texto (Clipboard)", command=acao_buscar_texto, width=20, height=2)
    botao_buscar_texto.pack(side=tk.LEFT, padx=5, pady=5)

    botao_ler_imagem = tk.Button(frame_botoes_acao, text="Ler de Imagem (Clipboard)", command=acao_ler_imagem, width=20, height=2)
    botao_ler_imagem.pack(side=tk.LEFT, padx=5, pady=5)

    # Configura comando do botão de ler agora que a função está definida
    botao_ler_audio.config(command=acao_ler_texto_area_openai)
    botao_ler_audio.pack(side=tk.LEFT, padx=5, pady=5)
    if not openai_client_ok:
        botao_ler_audio.config(state=tk.DISABLED, text="Ler com OpenAI (Erro API Key)")

    # Empacota o botão de parar
    botao_parar_audio.pack(side=tk.LEFT, padx=5, pady=5)


    # Botões de controle (Monitoramento e Fechar)
    botao_monitor = tk.Button(frame_botoes_controle, text="Iniciar Monitoramento", width=20, height=2)
    botao_monitor.config(command=lambda: iniciar_parar_monitoramento(text_area, janela, botao_monitor))
    botao_monitor.pack(side=tk.LEFT, padx=5, pady=5)

    botao_fechar = tk.Button(frame_botoes_controle, text="Fechar", command=lambda: ao_fechar(janela), width=10, height=2)
    botao_fechar.pack(side=tk.LEFT, padx=5, pady=5)

    # --- Tratamento ao Fechar Janela ---
    def ao_fechar(janela_a_fechar):
        global monitorar_clipboard, audio_temp_file, stop_audio_flag
        print("Fechando a aplicação...")
        monitorar_clipboard = False
        stop_audio_flag.set() # Sinaliza para parar qualquer áudio em andamento
        sd.stop() # Tenta parar qualquer reprodução ativa imediatamente

        # Espera um pouco para a thread de áudio (se existir) terminar
        if audio_thread and audio_thread.is_alive():
             audio_thread.join(timeout=0.5) # Espera meio segundo

        if audio_temp_file and os.path.exists(audio_temp_file):
             try:
                 os.remove(audio_temp_file)
                 print(f"Arquivo temporário removido ao fechar: {audio_temp_file}")
             except Exception as e:
                 print(f"Aviso: Não foi possível remover o arquivo temp ao fechar: {e}")

        janela_a_fechar.destroy()

    janela.protocol("WM_DELETE_WINDOW", lambda: ao_fechar(janela))

    # --- Mensagem de erro inicial se OpenAI falhou ---
    if not openai_client_ok:
         messagebox.showwarning("Aviso OpenAI", "Não foi possível conectar à API da OpenAI.\nVerifique sua chave de API e conexão.\nA função de leitura por voz está desabilitada.", parent=janela)

    # --- Iniciar a GUI ---
    janela.mainloop()

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    # Verificar Tesseract no início
    try:
        print(f"Tesseract encontrado: {pytesseract.get_tesseract_version()}")
    except pytesseract.TesseractNotFoundError:
        print("AVISO: Tesseract não encontrado no PATH.")

    # Verificar dispositivo de áudio padrão (opcional, mas útil para debug)
    try:
        print(f"Dispositivo de áudio padrão: {sd.query_devices(kind='output')}")
    except Exception as e:
        print(f"Aviso: Não foi possível consultar dispositivos de áudio: {e}")

    criar_janela_preview()
