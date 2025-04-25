import tkinter as tk
from tkinter import scrolledtext, messagebox
import pyperclip
# from openai import OpenAI # Removido
# import sounddevice as sd # Removido
# import soundfile as sf # Removido
# import numpy as np # Removido
import threading
import time
from PIL import Image, ImageGrab
import pytesseract
import io # Mantido por segurança, pode ser removido se não usado em outro lugar
import os # Mantido por segurança, pode ser removido se não usado em outro lugar
import tempfile # Mantido por segurança, pode ser removido se não usado em outro lugar
import platform # Adicionado para verificar o OS
import sys # Adicionado para sair se não for Windows

# --- Verificação do Sistema Operacional ---
is_windows = platform.system() == "Windows"
if not is_windows:
    print("ERRO: Esta versão do script usa pywin32 e só funciona no Windows.")
    # Mostrar uma messagebox se o Tkinter já puder ser inicializado
    try:
        root = tk.Tk()
        root.withdraw() # Esconde a janela principal
        messagebox.showerror("Erro de Plataforma", "Esta aplicação requer Windows para a funcionalidade de Text-to-Speech (pywin32).")
        root.destroy()
    except Exception:
        pass # Evita erro se o Tkinter não puder ser inicializado
    sys.exit(1) # Sai do script

# Tenta importar pywin32 apenas se for Windows
try:
    import win32com.client as wincl
    # Tenta inicializar COM para a thread principal (necessário para pywin32)
    try:
        import pythoncom
        pythoncom.CoInitialize()
        print("COM inicializado para thread principal.")
    except ImportError:
        print("Aviso: pythoncom não encontrado. A inicialização COM pode falhar.")
    except Exception as com_err:
        print(f"Aviso: Erro ao inicializar COM: {com_err}")

    pywin32_ok = True
    print("pywin32 importado com sucesso.")
except ImportError:
    print("ERRO: A biblioteca pywin32 não está instalada.")
    print("Execute: pip install pywin32")
    # Mostrar messagebox
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Erro de Dependência", "A biblioteca 'pywin32' é necessária.\nExecute 'pip install pywin32' e tente novamente.")
        root.destroy()
    except Exception: pass
    sys.exit(1)
except Exception as e:
     print(f"ERRO inesperado ao importar pywin32: {e}")
     pywin32_ok = False # Define como False se houver outro erro na importação


# --- Configuração Inicial ---
# [Windows Apenas] Tesseract Path (se necessário)
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\orlando.santos\tesseract-installation'
    print(f"Tentando usar Tesseract em: {pytesseract.pytesseract.tesseract_cmd}")
    tesseract_version_check = pytesseract.get_tesseract_version()
    print(f"Versão do Tesseract encontrada via caminho definido: {tesseract_version_check}")
except Exception as e:
    print(f"Aviso: Não foi possível definir ou verificar o caminho do Tesseract: {e}")

# --- Cliente OpenAI Removido ---
# client = None
# openai_client_ok = False

# --- Variáveis Globais de Controle de Áudio ---
audio_thread = None
stop_audio_flag = threading.Event() # Evento para sinalizar parada
sapi_voice_object = None # Para manter a instância SAPI

# --- Funções Clipboard e OCR ---
# (Funções obter_texto_clipboard, obter_imagem_clipboard, extrair_texto_de_imagem inalteradas)
def obter_texto_clipboard():
    try: return pyperclip.paste()
    except Exception as e: messagebox.showerror("Erro Clipboard", f"Não acesso texto.\n{e}"); return ""
def obter_imagem_clipboard():
    try:
        imagem = ImageGrab.grabclipboard(); return imagem if isinstance(imagem, Image.Image) else None
    except Exception as e: print(f"Erro ao obter imagem: {e}"); return None
def extrair_texto_de_imagem(imagem):
    if not isinstance(imagem, Image.Image): messagebox.showinfo("OCR", "Nenhuma imagem válida."); return ""
    try:
        lang = 'por'; print(f"Executando OCR com idioma: {lang}")
        texto_extraido = pytesseract.image_to_string(imagem, lang=lang)
        if not texto_extraido.strip(): messagebox.showinfo("OCR", "Não extraiu texto."); return ""
        print("Texto extraído."); return texto_extraido.strip()
    except PermissionError as pe: print(f"ERRO PERMISSÃO Tesseract: {pe}"); messagebox.showerror("Erro Permissão OCR", f"Acesso negado Tesseract:\n{pe}"); return ""
    except pytesseract.TesseractNotFoundError: messagebox.showerror("Erro OCR", "Tesseract não encontrado."); return ""
    except Exception as e: messagebox.showerror("Erro OCR", f"Erro inesperado OCR:\n{e}"); return ""

# --- Função TTS com SAPI (pywin32) ---
def inicializar_sapi():
    """Inicializa e retorna o objeto SAPI SpVoice."""
    global sapi_voice_object
    if sapi_voice_object is None:
        try:
            print("Inicializando objeto SAPI SpVoice...")
            # Garante que COM foi inicializado para a thread (importante!)
            # pythoncom não é importado diretamente, mas usado por win32com
            try:
                 import pythoncom
                 # Tenta inicializar COM. Se já estiver inicializado pela thread principal,
                 # pode levantar um erro específico que podemos ignorar ou tratar.
                 # Usar CoInitializeEx com COINIT_MULTITHREADED pode ser uma alternativa,
                 # mas CoInitialize() é geralmente suficiente se chamado em cada thread que usa COM.
                 pythoncom.CoInitialize()
                 print("COM inicializado para esta thread.")
            except ImportError:
                 print("Aviso: pythoncom não encontrado. A inicialização COM pode falhar em algumas threads.")
            except Exception as com_err:
                 # Verifica se o erro é porque já foi inicializado (com código de erro específico)
                 # O código de erro pode variar, mas S_FALSE (-2147417850) ou RPC_E_CHANGED_MODE (0x80010106) são comuns
                 if hasattr(com_err, 'hresult') and com_err.hresult in [-2147417850, -2147221008]: # S_FALSE ou CO_E_ALREADYINITIALIZED
                      print("COM já inicializado para esta thread.")
                 else:
                      print(f"Aviso: Erro ao inicializar COM nesta thread: {com_err}")


            sapi_voice_object = wincl.Dispatch("SAPI.SpVoice")
            print(sapi_voice_object.GetVoices())
            # Opcional: Listar e selecionar vozes
            voices = sapi_voice_object.GetVoices()
            for i, v in enumerate(voices):
                print(f"{i}: {v.GetDescription()}")
            # Exemplo: Tentar encontrar uma voz em português
            target_voice_index = -1
            for i, v in enumerate(voices):
                 # Ajuste a string de busca conforme necessário (ex: 'Maria', 'Portuguese')
                 if 'maria' in v.GetDescription().lower() :
                      target_voice_index = i
                      break
            if target_voice_index != -1:
                 print(f"Selecionando voz: {voices.Item(target_voice_index).GetDescription()}")
                 sapi_voice_object.Voice = voices.Item(target_voice_index)

            print("Objeto SAPI inicializado.")
            return sapi_voice_object
        except Exception as e:
            print(f"ERRO CRÍTICO ao inicializar SAPI: {e}")
            messagebox.showerror("Erro SAPI", f"Não foi possível inicializar a API de Fala do Windows (SAPI).\n{e}")
            sapi_voice_object = None # Garante que está None se falhar
            return None
    return sapi_voice_object

def ler_texto_pywin32(texto, janela_ref, botao_ler, botao_parar):
    """Usa SAPI (pywin32) para ler o texto de forma assíncrona."""
    global audio_thread, stop_audio_flag, sapi_voice_object

    if not pywin32_ok: # Verifica se a importação inicial funcionou
         messagebox.showerror("Erro pywin32", "A biblioteca pywin32 não foi carregada corretamente.", parent=janela_ref)
         return

    if audio_thread and audio_thread.is_alive():
        messagebox.showwarning("Leitura", "Uma leitura já está em andamento.", parent=janela_ref)
        return

    stop_audio_flag.clear()

    if not texto:
        messagebox.showinfo("Leitura", "Nenhum texto fornecido para leitura.", parent=janela_ref)
        return

    # Feedback visual e botões (feito antes de iniciar a thread)
    janela_ref.config(cursor="watch")
    botao_ler.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL)
    janela_ref.update_idletasks()

    def tarefa_leitura_sapi():
        """Função executada na thread para controlar a fala SAPI."""
        nonlocal texto, janela_ref, botao_ler, botao_parar # Captura variáveis
        speak = None # Define speak como None inicialmente
        try:
            # *** Tenta inicializar SAPI dentro da thread ***
            speak = inicializar_sapi()
            if speak is None:
                 raise Exception("Falha ao inicializar SAPI na thread de leitura.") # Levanta exceção se falhar

            # Flags para fala assíncrona e para cancelar fala anterior
            SVSFlagsAsync = 1
            SVSFPurgeBeforeSpeak = 2

            voices = sapi_voice_object.GetVoices()
            print("Vozes disponíveis:")
            for i, v in enumerate(voices):
                print(f"{i}: {v.GetDescription()}")

            print(f"Falando com SAPI: {texto[:50]}...")
            # Inicia a fala assíncrona, cancelando qualquer fala anterior
            speak.Speak(texto, SVSFlagsAsync | SVSFPurgeBeforeSpeak)

            time.sleep(0.5)
            # Loop para esperar o fim da fala ou a interrupção
            # speak.Status.RunningState == 2 indica que está falando
            while speak.Status.RunningState == 2:
                if stop_audio_flag.is_set():
                    # Tenta parar a fala atual enviando um comando de purge
                    print("Parada solicitada, tentando interromper SAPI...")
                    # Usa speak (objeto SAPI inicializado na thread)
                    speak.Speak("", SVSFlagsAsync | SVSFPurgeBeforeSpeak) # Envia comando vazio para parar
                    break # Sai do loop de espera
                time.sleep(0.1) # Pausa para não sobrecarregar

            if not stop_audio_flag.is_set():
                print("Fala SAPI concluída.")
            else:
                print("Fala SAPI interrompida.")

        except Exception as e:
            print(f"Erro durante a fala SAPI: {e}")
            # *** CORREÇÃO APLICADA AQUI ***
            # Captura o valor de 'e' usando um argumento padrão na lambda
            janela_ref.after(0, lambda err=e: messagebox.showerror("Erro SAPI", f"Ocorreu um erro durante a fala:\n{err}", parent=janela_ref))
        finally:
            # Restauração da GUI (sempre executa, via 'after')
            def restaurar_gui_thread_safe():
                janela_ref.config(cursor="")
                botao_ler.config(state=tk.NORMAL) # Reabilita botão de ler
                botao_parar.config(state=tk.DISABLED) # Desabilita parar
            janela_ref.after(0, restaurar_gui_thread_safe)

            # Opcional: Desinicializar COM para a thread
            try:
                 import pythoncom
                 pythoncom.CoUninitialize()
                 print("COM desinicializado para esta thread.")
            except ImportError:
                 pass # Ignora se pythoncom não estiver disponível
            except Exception as com_err:
                 print(f"Aviso: Erro ao desinicializar COM: {com_err}")


    # Inicia a thread
    audio_thread = threading.Thread(target=tarefa_leitura_sapi)
    audio_thread.daemon = True
    audio_thread.start()

# --- Função para Parar Leitura (SAPI) ---
def acao_parar_leitura():
    """Sinaliza para a thread de áudio SAPI parar a reprodução."""
    global stop_audio_flag, audio_thread
    if audio_thread and audio_thread.is_alive():
        print("Sinalizando para parar a leitura SAPI...")
        stop_audio_flag.set() # Aciona o evento que a thread verifica
    else:
        print("Nenhuma leitura SAPI em andamento para parar.")

# --- Funções de Monitoramento ---
# (Funções tarefa_monitoramento_clipboard, atualizar_texto_area_monitor, iniciar_parar_monitoramento inalteradas)
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
    janela = tk.Tk(); janela.title("Leitor de Tela (OCR + SAPI TTS)"); janela.geometry("650x550"); janela.attributes('-topmost', True)
    frame_instrucoes = tk.Frame(janela, bd=1, relief=tk.SUNKEN); frame_instrucoes.pack(pady=5, padx=10, fill=tk.X)
    label_instrucoes = tk.Label(frame_instrucoes, text=("Uso (Windows SAPI):\n- Copie texto/imagem e use os botões 'Buscar'/'Ler Imagem'.\n- Clique 'Ler com SAPI' para ouvir (voz do Windows).\n- Use 'Parar' para interromper."), justify=tk.LEFT); label_instrucoes.pack(pady=5, padx=5)
    text_area = scrolledtext.ScrolledText(janela, wrap=tk.WORD, width=70, height=18, font=("Arial", 11)); text_area.pack(pady=5, padx=10, expand=True, fill='both')
    frame_botoes_acao = tk.Frame(janela); frame_botoes_acao.pack(pady=(5, 0))
    frame_botoes_controle = tk.Frame(janela); frame_botoes_controle.pack(pady=(0, 10))

    botao_ler_imagem = tk.Button(frame_botoes_acao, text="Ler Imagem", width=20, height=2)
    # Botão de leitura agora usa SAPI
    botao_ler_audio = tk.Button(frame_botoes_acao, text="Ler com SAPI", width=20, height=2, fg="blue") # Texto/Cor atualizado
    botao_parar_audio = tk.Button(frame_botoes_acao, text="Parar Leitura", command=acao_parar_leitura, width=15, height=2, state=tk.DISABLED)

    def acao_buscar_texto(): texto_cb = obter_texto_clipboard(); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_cb) if texto_cb else None
    def acao_ler_imagem():
        janela.config(cursor="watch"); botao_ler_imagem.config(state=tk.DISABLED); botao_ler_audio.config(state=tk.DISABLED); botao_parar_audio.config(state=tk.DISABLED)
        janela.update_idletasks()
        try:
            imagem_cb = obter_imagem_clipboard()
            if imagem_cb: texto_ocr = extrair_texto_de_imagem(imagem_cb); text_area.delete("1.0", tk.END); text_area.insert(tk.INSERT, texto_ocr) if texto_ocr else None
            else: messagebox.showinfo("OCR", "Nenhuma imagem.", parent=janela)
        finally: janela.config(cursor=""); botao_ler_imagem.config(state=tk.NORMAL); botao_ler_audio.config(state=tk.NORMAL) # SAPI está sempre 'disponível' no Windows

    # Função que chama a leitura SAPI
    def acao_ler_texto_area_sapi():
        texto_para_ler = text_area.get("1.0", tk.END).strip()
        ler_texto_pywin32(texto_para_ler, janela, botao_ler_audio, botao_parar_audio) # Chama a nova função

    botao_buscar_texto = tk.Button(frame_botoes_acao, text="Buscar Texto", command=acao_buscar_texto, width=20, height=2); botao_buscar_texto.pack(side=tk.LEFT, padx=5, pady=5)
    botao_ler_imagem.config(command=acao_ler_imagem); botao_ler_imagem.pack(side=tk.LEFT, padx=5, pady=5)
    # Configura comando do botão de ler para a função SAPI
    botao_ler_audio.config(command=acao_ler_texto_area_sapi)
    botao_ler_audio.pack(side=tk.LEFT, padx=5, pady=5)
    botao_parar_audio.pack(side=tk.LEFT, padx=5, pady=5)
    botao_monitor = tk.Button(frame_botoes_controle, text="Iniciar Monitoramento", width=20, height=2); botao_monitor.config(command=lambda: iniciar_parar_monitoramento(text_area, janela, botao_monitor)); botao_monitor.pack(side=tk.LEFT, padx=5, pady=5)
    botao_fechar = tk.Button(frame_botoes_controle, text="Fechar", command=lambda: ao_fechar(janela), width=10, height=2); botao_fechar.pack(side=tk.LEFT, padx=5, pady=5)

    def ao_fechar(janela_a_fechar):
        global monitorar_clipboard, stop_audio_flag, audio_thread, sapi_voice_object
        print("Fechando a aplicação..."); monitorar_clipboard = False; stop_audio_flag.set()
        # Tenta parar SAPI diretamente
        if sapi_voice_object:
             try:
                  # Envia comando vazio para interromper fala assíncrona
                  # É importante usar o objeto SAPI que foi inicializado
                  # Pode ser necessário garantir que 'speak' esteja acessível ou usar 'sapi_voice_object'
                  sapi_voice_object.Speak("", 3) # 3 = SVSFlagsAsync | SVSFPurgeBeforeSpeak
             except Exception as e:
                  print(f"Info: Erro ao tentar parar SAPI no fechamento: {e}")

        if audio_thread and audio_thread.is_alive(): print("Aguardando thread de áudio..."); audio_thread.join(timeout=0.5)
        # Não há arquivo temporário de áudio para remover com SAPI

        # Opcional: Desinicializar COM para a thread principal
        try:
             import pythoncom
             pythoncom.CoUninitialize()
             print("COM desinicializado para thread principal.")
        except ImportError: pass
        except Exception as com_err: print(f"Aviso: Erro ao desinicializar COM principal: {com_err}")

        janela_a_fechar.destroy()
    janela.protocol("WM_DELETE_WINDOW", lambda: ao_fechar(janela))

    # Não há verificação de cliente OpenAI necessária aqui
    janela.mainloop()

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    if not is_windows or not pywin32_ok:
         print("Saindo devido a erro de plataforma ou dependência.")
         sys.exit(1) # Sai se não for Windows ou pywin32 falhou

    try: print(f"Tesseract encontrado: {pytesseract.get_tesseract_version()}")
    except pytesseract.TesseractNotFoundError: print("AVISO: Tesseract OCR não encontrado no PATH.")
    # Não precisa verificar dispositivo de áudio com SAPI

    # Tenta inicializar SAPI uma vez no início (opcional, mas bom para feedback rápido)
    # if inicializar_sapi() is None:
    #      print("Falha ao inicializar SAPI no início. A leitura de voz pode não funcionar.")
         # Poderia mostrar um messagebox aqui

    criar_janela_preview()

