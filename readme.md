# Leitor OpenAI e SAPI

Este projeto é uma aplicação Python que utiliza interfaces gráficas (Tkinter) para realizar leitura de texto e imagens, com suporte a síntese de voz utilizando a API OpenAI ou o motor SAPI do Windows. Ele também inclui funcionalidades de OCR (reconhecimento óptico de caracteres) para extrair texto de imagens e monitoramento da área de transferência.

## Funcionalidades

- **Leitura de Texto**: Permite colar texto ou buscar diretamente da área de transferência para leitura em voz alta.
- **Leitura de Imagens**: Extrai texto de imagens copiadas para a área de transferência utilizando o Tesseract OCR.
- **Síntese de Voz**:
  - **OpenAI**: Utiliza a API OpenAI para leitura de texto em voz alta.
  - **SAPI (Windows)**: Usa o motor SAPI para leitura com suporte a destaque de palavras.
  - **PyWin32**: Integra-se ao Windows para leitura de texto com vozes do sistema.
  - **IBM Watson**: Suporte ao serviço de síntese de voz da IBM para leitura em voz alta.
- **Monitoramento da Área de Transferência**: Detecta automaticamente novos textos copiados e os exibe na interface.
- **Destaque de Texto**: Durante a leitura com SAPI, as palavras são destacadas em tempo real.
- **Controle de Voz**: Permite selecionar vozes disponíveis no sistema e ajustar a velocidade da leitura.

## Requisitos

### Dependências Python
- `tkinter`
- `pytesseract`
- `pyperclip`
- `Pillow`
- `pywin32` (somente para Windows)

### Outros Requisitos
- **Tesseract OCR**: Certifique-se de que o Tesseract está instalado e configurado no PATH do sistema.
- **Windows**: O suporte ao SAPI requer o sistema operacional Windows.

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/orlando-vandres/leitor-openai-sapi.git
   cd leitor-openai-sapi
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Certifique-se de que o Tesseract OCR está instalado e configurado no PATH.

4. Para usar a funcionalidade OpenAI, configure sua chave de API como variável de ambiente:
   ```bash
   export OPENAI_API_KEY="sua-chave-aqui"
   ```

## Uso

1. Execute o programa:
   ```bash
   python leitor_openai.py
   ```

2. **Interface Gráfica**:
   - Cole texto ou copie uma imagem para a área de transferência.
   - Use os botões para buscar texto, realizar OCR ou iniciar a leitura.
   - Selecione a voz e a velocidade desejadas (no modo SAPI).

3. **Monitoramento**:
   - Clique em "Iniciar Monitoramento" para detectar automaticamente novos textos copiados.

4. **Fechamento**:
   - O programa limpa arquivos temporários e encerra threads ao ser fechado.

## Estrutura do Projeto

- `leitor_openai.py`: Implementação principal com suporte à API OpenAI.
- `leitor_sapi_tts.py`: Implementação do motor SAPI com destaque de texto.
- `requirements.txt`: Lista de dependências do projeto.

## Observações

- O suporte ao OpenAI requer uma chave de API válida.
- O Tesseract OCR deve estar corretamente configurado para o reconhecimento de texto em imagens.
- O motor SAPI está disponível apenas no Windows.

## Licença

Este projeto é distribuído sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.