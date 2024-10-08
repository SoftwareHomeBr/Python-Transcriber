import os
import whisper
import threading
from docx import Document
from tkinter import Tk, Label, Button, filedialog, StringVar, ttk, BooleanVar, Checkbutton

# Variável global para o cancelamento da transcrição
cancelar_transcricao = False

# Extensões de áudio suportadas
EXTENSOES_SUPORTADAS = ('.mp3', '.wav', '.flac', '.m4a', '.ogg', '.mpeg')

# Descrições dos modelos Whisper
MODELOS_DESCRICAO = {
    "tiny": "Tiny: O modelo mais leve e rápido, ideal para tarefas rápidas com precisão básica; requer poucos recursos.",
    "base": "Base: Equilíbrio entre velocidade e precisão moderada, funciona bem em dispositivos com recursos limitados.",
    "small": "Small: Boa precisão e velocidade razoável, recomendado para uso geral; requer recursos moderados.",
    "medium": "Medium: Muito preciso, excelente para diferentes sotaques e contextos, ideal para máquinas com boa capacidade.",
    "large": "Large: Máxima precisão, melhor qualidade de transcrição; exige muitos recursos, ideal para máquinas potentes.",
    "turbo": "Turbo: Otimizado para máxima velocidade com alta precisão, indicado para servidores ou estações de alta potência."
}


class TranscricaoWhisper:
    def __init__(self, modelo_escolhido):
        self.modelo = self.carregar_modelo(modelo_escolhido)

    def carregar_modelo(self, modelo_escolhido):
        try:
            print(f"Carregando o modelo '{modelo_escolhido}'...")
            return whisper.load_model(modelo_escolhido)
        except Exception as e:
            print(f"Erro ao carregar o modelo: {e}")
            return None

    def transcrever(self, caminho_audio, progresso_var, status_label, index, total):
        global cancelar_transcricao
        if cancelar_transcricao:
            status_label.config(text="Transcrição cancelada.")
            return False
        status_label.config(text=f"Transcrevendo arquivo {index + 1} de {total}: {os.path.basename(caminho_audio)}")
        try:
            resposta = self.modelo.transcribe(caminho_audio)
            if cancelar_transcricao:
                status_label.config(text="Transcrição cancelada.")
                return False
            texto_transcrito = resposta['text']
            self.salvar_transcricao(texto_transcrito, caminho_audio)
            progresso_var.set(100)
            return True
        except RuntimeError as e:
            status_label.config(text="Erro ao transcrever: arquivo corrompido ou incompatível.")
        except Exception as e:
            status_label.config(text=f"Erro inesperado: {e}")
        return False

    @staticmethod
    def salvar_transcricao(texto_transcrito, caminho_audio):
        nome_arquivo, _ = os.path.splitext(os.path.basename(caminho_audio))
        diretorio_arquivo = os.path.dirname(caminho_audio)
        caminho_arquivo_docx = os.path.join(diretorio_arquivo, f"{nome_arquivo}_transcrito.docx")
        doc = Document()
        doc.add_heading(f"Transcrição de {nome_arquivo}", level=1)
        doc.add_paragraph(texto_transcrito)
        doc.save(caminho_arquivo_docx)
        print(f"Transcrição salva no arquivo: {caminho_arquivo_docx}")


# Função para selecionar arquivos de áudio em lote com suporte para subpastas
def selecionar_arquivos_audio(pasta, incluir_subpastas):
    arquivos_audio = []
    if incluir_subpastas:
        for root_dir, _, files in os.walk(pasta):
            arquivos_audio.extend([os.path.join(root_dir, f) for f in files if f.endswith(EXTENSOES_SUPORTADAS)])
    else:
        arquivos_audio = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith(EXTENSOES_SUPORTADAS)]
    return arquivos_audio


# Função para cancelar a transcrição
def cancelar_processo():
    global cancelar_transcricao
    cancelar_transcricao = True
    print("Processo de transcrição cancelado.")


# Função principal da interface gráfica
def criar_interface():
    global incluir_subpastas_checkbox

    root = Tk()
    root.title("Transcrição de Áudio com Whisper")
    root.geometry("450x420")
    root.resizable(False, False)

    progresso_var = StringVar(value=0)
    Label(root, text="Escolha o modelo de transcrição Whisper:", padx=20, pady=10).grid(row=0, column=0, columnspan=2)
    modelo_escolhido = StringVar(value="tiny")

    combobox = ttk.Combobox(root, textvariable=modelo_escolhido, values=list(MODELOS_DESCRICAO.keys()), width=15)
    combobox.grid(row=1, column=0, columnspan=2, padx=20, pady=5)

    status_label = Label(root, text=MODELOS_DESCRICAO[modelo_escolhido.get()], padx=20, pady=5, wraplength=400,
                         justify="left")
    status_label.grid(row=3, column=0, columnspan=2, padx=20, pady=5)

    def atualizar_descricao(event=None):
        descricao = MODELOS_DESCRICAO.get(modelo_escolhido.get(), "Descrição não disponível.")
        status_label.config(text=descricao)

    combobox.bind("<<ComboboxSelected>>", atualizar_descricao)

    # Botões de transcrição individual e em lote
    Button(root, text="Transcrição Individual",
           command=lambda: iniciar_transcricao(modelo_escolhido.get(), progresso_var, root, status_label,
                                               botao_cancelar, incluir_subpastas), width=25).grid(row=2, column=0,
                                                                                                  padx=20, pady=10)
    Button(root, text="Transcrição em Lote",
           command=lambda: iniciar_transcricao_em_lote(modelo_escolhido.get(), progresso_var, root, status_label,
                                                       botao_cancelar, incluir_subpastas), width=25).grid(row=2,
                                                                                                          column=1,
                                                                                                          padx=20,
                                                                                                          pady=10)

    incluir_subpastas = BooleanVar()
    incluir_subpastas_checkbox = Checkbutton(root, text="Incluir Subpastas", variable=incluir_subpastas)
    incluir_subpastas_checkbox.grid(row=4, column=0, columnspan=2, pady=5)

    progresso_barra = ttk.Progressbar(root, length=300, mode='determinate', variable=progresso_var)
    progresso_barra.grid(row=5, column=0, columnspan=2, padx=20, pady=20)

    botao_cancelar = Button(root, text="Cancelar Transcrição", command=cancelar_processo, width=25, state="disabled")
    botao_cancelar.grid(row=6, column=0, padx=20, pady=10)

    botao_fechar = Button(root, text="Fechar", command=root.quit, width=25)
    botao_fechar.grid(row=6, column=1, padx=20, pady=10)

    atualizar_descricao()

    root.mainloop()


def iniciar_transcricao(modelo_escolhido, progresso_var, root, status_label, botao_cancelar, incluir_subpastas):
    transcricao = TranscricaoWhisper(modelo_escolhido)
    caminho_audio = filedialog.askopenfilename(title="Selecione o arquivo de áudio", filetypes=[
        ("Arquivos de áudio", "*.mp3 *.wav *.flac *.m4a *.ogg *.mpeg")])
    if caminho_audio:
        botao_cancelar.config(state="normal")
        threading.Thread(target=transcricao.transcrever,
                         args=(caminho_audio, progresso_var, status_label, 0, 1)).start()


def iniciar_transcricao_em_lote(modelo_escolhido, progresso_var, root, status_label, botao_cancelar, incluir_subpastas):
    transcricao = TranscricaoWhisper(modelo_escolhido)
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos de áudio")
    if pasta:
        botao_cancelar.config(state="normal")
        arquivos_audio = selecionar_arquivos_audio(pasta, incluir_subpastas.get())
        threading.Thread(target=processar_em_lote,
                         args=(transcricao, arquivos_audio, progresso_var, status_label)).start()


def processar_em_lote(transcricao, arquivos_audio, progresso_var, status_label):
    global cancelar_transcricao
    progresso_total = 0
    progresso_por_arquivo = 100 / len(arquivos_audio) if arquivos_audio else 100
    for index, arquivo in enumerate(arquivos_audio):
        if cancelar_transcricao:
            status_label.config(text="Transcrição cancelada.")
            break
        transcricao.transcrever(arquivo, progresso_var, status_label, index, len(arquivos_audio))
        progresso_total += progresso_por_arquivo
        progresso_var.set(progresso_total)
    if not cancelar_transcricao:
        status_label.config(text=f"As {len(arquivos_audio)} transcrições foram realizadas com sucesso.")


if __name__ == "__main__":
    criar_interface()
