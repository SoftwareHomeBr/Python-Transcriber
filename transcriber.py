import os
import whisper
import threading
import logging
import time
from docx import Document
from tkinter import Tk, Label, Button, filedialog, StringVar, ttk, BooleanVar, Checkbutton, Text, Scrollbar, Frame, \
    NORMAL, DISABLED
from pydub import AudioSegment

# Configuração básica do logger
logging.basicConfig(filename='transcricao.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class TranscricaoAudio:
    EXTENSOES_SUPORTADAS = ('.mp3', '.wav', '.flac', '.m4a', '.ogg', '.mpeg')
    MODELOS_DESCRICAO = {
        "tiny": "Tiny: O modelo mais leve e rápido, ideal para tarefas rápidas com precisão básica; requer poucos recursos.",
        "base": "Base: Equilíbrio entre velocidade e precisão moderada, funciona bem em dispositivos com recursos limitados.",
        "small": "Small: Boa precisão e velocidade razoável, recomendado para uso geral; requer recursos moderados.",
        "medium": "Medium: Muito preciso, excelente para diferentes sotaques e contextos, ideal para máquinas com boa capacidade.",
        "large": "Large: Máxima precisão, melhor qualidade de transcrição; exige muitos recursos, ideal para máquinas potentes.",
        "turbo": "Turbo: Otimizado para máxima velocidade com alta precisão, indicado para servidores ou estações de alta potência."
    }

    def __init__(self):
        self.cancelar_transcricao = False
        self.pausar_transcricao = False
        self.total_bytes = 0
        self.processed_bytes = 0
        self.root = Tk()
        self.modelo_escolhido = StringVar(value="tiny")
        self.progresso_var = StringVar(value=0)
        self.incluir_subpastas = BooleanVar()
        self.formato_saida = StringVar(value="docx")
        self._configurar_interface()

    def _configurar_interface(self):
        self.root.title("Transcrição de Áudio com Whisper")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        # Barra de Progresso
        self.progresso_barra = ttk.Progressbar(self.root, length=600, mode='determinate', variable=self.progresso_var)
        self.progresso_barra.grid(row=6, column=0, columnspan=2, padx=20, pady=20)

        # Label para o progresso, imediatamente após a barra de progresso
        self.progresso_text_label = Label(self.root, text="", wraplength=600, justify="left")
        self.progresso_text_label.grid(row=7, column=0, columnspan=2, padx=20, pady=5)

        # Movendo outros elementos para a linha seguinte
        self.botao_cancelar = Button(self.root, text="Cancelar Transcrição", command=self.cancelar_processo, width=25,
                                     state="disabled")
        self.botao_cancelar.grid(row=8, column=0, padx=20, pady=10)
        self.botao_pausar = Button(self.root, text="Pausar Transcrição", command=self.pausar_processo, width=25,
                                   state="disabled")
        self.botao_pausar.grid(row=8, column=1, padx=20, pady=10)
        # Caixa de Seleção do Modelo
        Label(self.root, text="Escolha o modelo de transcrição Whisper:", padx=20, pady=10).grid(row=0, column=0,
                                                                                                 columnspan=2)
        self.combobox_modelo = ttk.Combobox(self.root, textvariable=self.modelo_escolhido,
                                            values=list(self.MODELOS_DESCRICAO.keys()), width=15)
        self.combobox_modelo.grid(row=1, column=0, columnspan=2, padx=20, pady=5)
        self.combobox_modelo.bind("<<ComboboxSelected>>", self._atualizar_descricao_modelo)

        # Explicação sobre o Modelo
        self.modelo_explicacao = Label(self.root, text=self.MODELOS_DESCRICAO[self.modelo_escolhido.get()],
                                       wraplength=600, justify="left")
        self.modelo_explicacao.grid(row=2, column=0, columnspan=2, padx=20, pady=5)

        # Botões de Transcrição (na mesma linha)
        Button(self.root, text="Transcrição Individual", command=self.iniciar_transcricao, width=25).grid(row=3,
                                                                                                          column=0,
                                                                                                          padx=20,
                                                                                                          pady=10)
        Button(self.root, text="Transcrição em Lote", command=self.iniciar_transcricao_em_lote, width=25).grid(row=3,
                                                                                                               column=1,
                                                                                                               padx=20,
                                                                                                               pady=10)

        # Checkbox "Incluir Subpastas"
        Checkbutton(self.root, text="Incluir Subpastas", variable=self.incluir_subpastas).grid(row=4, column=0,
                                                                                               columnspan=2, pady=5)

        # Formato de Saída
        Label(self.root, text="Formato de Saída:", padx=20).grid(row=5, column=0)
        ttk.Combobox(self.root, textvariable=self.formato_saida, values=["docx", "txt", "markdown"], width=10).grid(
            row=5, column=1)

        # Barra de Progresso
        self.progresso_barra = ttk.Progressbar(self.root, length=600, mode='determinate', variable=self.progresso_var)
        self.progresso_barra.grid(row=6, column=0, columnspan=2, padx=20, pady=20)

        # Adiciona uma Label para exibir progresso
        self.progresso_text_label = Label(self.root, text="", wraplength=600, justify="left")
        self.progresso_text_label.grid(row=7, column=0, columnspan=2, padx=20, pady=5)

        # Botões "Cancelar Transcrição" e "Pausar Transcrição" (na mesma linha)
        self.botao_cancelar = Button(self.root, text="Cancelar Transcrição", command=self.cancelar_processo, width=25,
                                     state="disabled")
        self.botao_cancelar.grid(row=8, column=0, padx=20, pady=10)
        self.botao_pausar = Button(self.root, text="Pausar Transcrição", command=self.pausar_processo, width=25,
                                   state="disabled")
        self.botao_pausar.grid(row=8, column=1, padx=20, pady=10)
        self.botao_pausar.pack_forget()
        self.botao_cancelar.pack_forget()

        # Botões "Ver Mais Detalhes" e "Fechar"
        self.botao_detalhes = Button(self.root, text="Ver Mais Detalhes", command=self._alternar_painel_detalhes,
                                     width=25)
        self.botao_detalhes.grid(row=9, column=0, padx=20, pady=10)
        self.botao_fechar = Button(self.root, text="Fechar", command=self.root.quit, width=25)
        self.botao_fechar.grid(row=9, column=1, padx=20, pady=10)

        # Painel de Detalhes
        self.painel_detalhes = Frame(self.root)
        self.detalhes_text = Text(self.painel_detalhes, height=8, width=70, state=DISABLED, wrap="word")
        self.detalhes_text.pack(side="left", fill="both", expand=True)
        scrollbar = Scrollbar(self.painel_detalhes, command=self.detalhes_text.yview)
        self.detalhes_text['yscrollcommand'] = scrollbar.set
        scrollbar.pack(side="right", fill="y")

    def _atualizar_descricao_modelo(self, event=None):
        descricao = self.MODELOS_DESCRICAO.get(self.modelo_escolhido.get(), "Descrição não disponível.")
        self.modelo_explicacao.config(text=descricao)

    def carregar_modelo(self):
        logging.info(f"Carregando o modelo '{self.modelo_escolhido.get()}'...")
        return whisper.load_model(self.modelo_escolhido.get())

    def iniciar_transcricao(self):
        self._limpar_detalhes()
        caminho_audio = filedialog.askopenfilename(title="Selecione o arquivo de áudio", filetypes=[
            ("Arquivos de áudio", "*.mp3 *.wav *.flac *.m4a *.ogg *.mpeg")])
        if caminho_audio:
            self.start_time = time.time()
            self._habilitar_botoes_transcricao(True)
            self.total_bytes = os.path.getsize(caminho_audio)
            self.processed_bytes = 0
            threading.Thread(target=self.transcrever_audio, args=(self.carregar_modelo(), caminho_audio, 1, 1)).start()
        else:
            self.progresso_text_label.config(text="Nenhum arquivo selecionado.")

    def iniciar_transcricao_em_lote(self):
        self._limpar_detalhes()
        pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos de áudio")
        if pasta:
            self.start_time = time.time()
            self._habilitar_botoes_transcricao(True)
            arquivos_audio = self._selecionar_arquivos_audio(pasta)
            self.total_bytes = sum(os.path.getsize(arquivo) for arquivo in arquivos_audio)
            self.processed_bytes = 0
            threading.Thread(target=self.processar_em_lote, args=(self.carregar_modelo(), arquivos_audio)).start()
        else:
            self.progresso_text_label.config(text="Nenhuma pasta selecionada.")

    def _selecionar_arquivos_audio(self, pasta):
        arquivos_audio = []
        if self.incluir_subpastas.get():
            for root_dir, _, files in os.walk(pasta):
                arquivos_audio.extend(
                    [os.path.join(root_dir, f) for f in files if f.endswith(self.EXTENSOES_SUPORTADAS)])
        else:
            arquivos_audio = [os.path.join(pasta, f) for f in os.listdir(pasta) if
                              f.endswith(self.EXTENSOES_SUPORTADAS)]
        return arquivos_audio

    def transcrever_audio(self, modelo, caminho_audio, indice, total, ultimo_arquivo=False):
        pos_inicial = self._inserir_detalhes(f"Transcrevendo {caminho_audio}")
        segment_duration = 30
        audio = AudioSegment.from_file(caminho_audio)
        duration = len(audio) / 1000  # Duração total em segundos
        segments = range(0, int(duration), segment_duration)
        self.cancelar_transcricao = False
        self.pausar_transcricao = False
        try:
            if self.cancelar_transcricao:
                self.progresso_text_label.config(text="Transcrição cancelada.")
                return
            self.botao_pausar.grid()
            self.botao_cancelar.grid()
            self.progresso_var.set("0")
            self.progresso_text_label.config(text="iniciando . . . ")
            transcricao_completa = ""
            for i, start_time in enumerate(segments):
                end_time = min(start_time + segment_duration, duration)
                segment = audio[start_time * 1000:end_time * 1000]
                segment_path = f"segmento_{i}.wav"
                segment.export(segment_path, format="wav")

                # Transcreve o segmento
                result = modelo.transcribe(segment_path)
                transcricao_completa += result["text"] + " "

                # Monitorar progresso
                progresso = ((i + 1) / len(segments)) * 100
                ## print(f"Progresso: {progresso:.2f}% concluído")
                self.progresso_var.set(f"{progresso:.2f}")
                # Limpar o arquivo temporário
                os.remove(segment_path)

            ## resposta = modelo.transcribe(caminho_audio, verbose=True)
            if self.cancelar_transcricao:
                self.progresso_text_label.config(text="Transcrição cancelada.")
                return
            texto_transcrito = transcricao_completa  ## resposta['text']
            self.salvar_transcricao(texto_transcrito, caminho_audio, self.formato_saida.get())
            self.processed_bytes += os.path.getsize(caminho_audio)
            self._atualizar_progresso(indice, total)
            self._substituir_detalhes(pos_inicial, f"Transcrito com sucesso: {caminho_audio}")
        except RuntimeError:
            self.progresso_text_label.config(text="Erro: arquivo corrompido ou incompatível")
            self._substituir_detalhes(pos_inicial, f"Erro na transcrição: {caminho_audio}")
        except Exception as e:
            logging.error(f"Erro inesperado: {e}")
            self.progresso_text_label.config(text=f"Erro inesperado: {e}")
            self._substituir_detalhes(pos_inicial, f"Erro inesperado: {caminho_audio}")
        finally:
            # if ultimo_arquivo and not self.cancelar_transcricao:
            # self.progresso_text_label.config(text="Todas as transcrições foram realizadas com sucesso.")
            # self._habilitar_botoes_transcricao(False)
            self.botao_cancelar.grid_remove()
            self.botao_pausar.grid_remove()
            self.botao_cancelar.pack_forget()
            self.botao_pausar.pack_forget()
            if ultimo_arquivo and not self.cancelar_transcricao:
                self.progresso_text_label.config(text="Todas as transcrições foram realizadas com sucesso.")

    def _atualizar_progresso(self, indice, total):
        elapsed_time = time.time() - self.start_time
        horas, rem = divmod(elapsed_time, 3600)
        minutos, segundos = divmod(rem, 60)
        progresso_percentual = (self.processed_bytes / self.total_bytes) * 100 if self.total_bytes > 0 else 0
        self.progresso_var.set(progresso_percentual)

        # Atualiza a nova Label com o texto de progresso
        progresso_text = f"{progresso_percentual:.0f}% [{indice}/{total}] Tempo: {int(horas):02}:{int(minutos):02}:{int(segundos):02}"
        self.progresso_text_label.config(text=progresso_text)

    def salvar_transcricao(self, texto_transcrito, caminho_audio, formato='docx'):
        nome_arquivo, _ = os.path.splitext(os.path.basename(caminho_audio))
        caminho_saida = os.path.join(os.path.dirname(caminho_audio), f"{nome_arquivo}_transcrito.{formato}")

        if formato == 'txt':
            with open(caminho_saida, 'w', encoding='utf-8') as f:
                f.write(texto_transcrito)
        elif formato == 'markdown':
            with open(caminho_saida, 'w', encoding='utf-8') as f:
                f.write(f"# Transcrição de {nome_arquivo}\n\n{texto_transcrito}")
        else:  # Padrão para docx
            doc = Document()
            doc.add_heading(f"Transcrição de {nome_arquivo}", level=1)
            doc.add_paragraph(texto_transcrito)
            doc.save(caminho_saida)

        logging.info(f"Transcrição salva no arquivo: {caminho_saida}")

    def processar_em_lote(self, modelo, arquivos_audio):
        total = len(arquivos_audio)
        self.cancelar_transcricao = False
        self.pausar_transcricao = False
        for index, caminho_audio in enumerate(arquivos_audio):
            if self.cancelar_transcricao:
                self.progresso_text_label.config(text="Transcrição cancelada.")
                break
            while self.pausar_transcricao:
                pass
            ultimo_arquivo = index == total - 1
            self.transcrever_audio(modelo, caminho_audio, index + 1, total, ultimo_arquivo=ultimo_arquivo)
        self.botao_cancelar.grid_remove()
        self.botao_pausar.grid_remove()

    def cancelar_processo(self):
        self.cancelar_transcricao = True
        logging.info("Processo de transcrição cancelado.")

    def pausar_processo(self):
        self.pausar_transcricao = not self.pausar_transcricao
        novo_texto = "Continuar Transcrição" if self.pausar_transcricao else "Pausar Transcrição"
        self.botao_pausar.config(text=novo_texto, bg="orange" if self.pausar_transcricao else "SystemButtonFace")

    def _habilitar_botoes_transcricao(self, habilitar):
        estado = "normal" if habilitar else "disabled"
        self.botao_cancelar.config(state=estado)
        self.botao_pausar.config(state=estado)

    def _inserir_detalhes(self, mensagem):
        self.detalhes_text.config(state=NORMAL)
        pos_inicial = self.detalhes_text.index("end-1c")
        self.detalhes_text.insert("end", f"{mensagem}\n")
        self.detalhes_text.config(state=DISABLED)
        self.detalhes_text.see("end")
        return pos_inicial

    def _substituir_detalhes(self, pos_inicial, mensagem):
        self.detalhes_text.config(state=NORMAL)
        self.detalhes_text.delete(pos_inicial, f"{pos_inicial} + 1 line")
        self.detalhes_text.insert(pos_inicial, f"{mensagem}\n")
        self.detalhes_text.config(state=DISABLED)
        self.detalhes_text.see("end")

    def _alternar_painel_detalhes(self):
        if self.painel_detalhes.winfo_ismapped():
            self.painel_detalhes.grid_remove()
            self.botao_detalhes.config(text="Ver Mais Detalhes")
            self.root.geometry("700x500")
        else:
            self.painel_detalhes.grid(row=19, column=0, columnspan=2, padx=10, pady=5, sticky="we")
            self.botao_detalhes.config(text="Ocultar Detalhes")
            self.root.geometry("700x550")

    def _limpar_detalhes(self):
        self.detalhes_text.config(state=NORMAL)
        self.detalhes_text.delete(1.0, "end")
        self.detalhes_text.config(state=DISABLED)

    def iniciar_interface(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = TranscricaoAudio()
    app.iniciar_interface()