# Transcrição de Áudio com Whisper

Este programa realiza a transcrição automática de arquivos de áudio usando o modelo Whisper. A transcrição pode ser feita de forma individual ou em lote, permitindo processar múltiplos arquivos de áudio localizados em uma pasta e, opcionalmente, em suas subpastas. O programa possui uma interface gráfica simples construída com `Tkinter` e oferece suporte a formatos de áudio populares, como MP3, WAV, FLAC, M4A, OGG e MPEG.

## Funcionalidades

- **Transcrição Individual**: Selecione um único arquivo de áudio para transcrição.
- **Transcrição em Lote**: Selecione uma pasta e transcreva todos os arquivos de áudio compatíveis. Há uma opção para incluir arquivos em subpastas.
- **Cancelamento de Transcrição**: Permite cancelar o processo de transcrição em andamento.
- **Interface Gráfica**: Interface simples e intuitiva usando `Tkinter`.

## Tecnologias Utilizadas

- **Python 3.x**
- **Whisper**: Modelo de transcrição automática (biblioteca `whisper`).
- **Tkinter**: Biblioteca padrão para interfaces gráficas.
- **docx**: Utilizada para salvar as transcrições em arquivos `.docx`.

## Formatos de Áudio Suportados

- `.mp3`
- `.wav`
- `.flac`
- `.m4a`
- `.ogg`
- `.mpeg`

## Instalação

1. Clone o repositório ou baixe o código fonte.

2. Instale as dependências necessárias usando o `pip`:

```bash
pip install -r requirements.txt

````
O arquivo requirements.txt deve conter:
```
whisper
python-docx
tk
```

3. Para executar.
```bash
python transcriber.py
```

# Uso
Ao iniciar o programa, selecione o modelo Whisper desejado.

Escolha entre transcrição individual ou em lote.

Selecione os arquivos de áudio ou a pasta contendo os áudios.

Aguarde a transcrição ser concluída. As transcrições serão salvas automaticamente em arquivos .docx.
