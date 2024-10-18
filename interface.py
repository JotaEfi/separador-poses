import os
import pandas as pd
from tkinter import Tk, Button, Label, filedialog, messagebox
from main import (
    carregar_planilha,
    listar_pastas,
    contar_fotos_por_pose,
    contar_fotos_alunos,
    gerar_planilha_resultado,
    adicionar_somas,
)


class Interface:
    def __init__(self, master):
        self.master = master
        self.master.title("Contador de Fotos")

        self.label = Label(master, text="Contador de Fotos de Alunos")
        self.label.pack(pady=10)

        self.btn_selecionar_pasta_fotos = Button(
            master,
            text="Selecionar Pasta de Fotos das Poses",
            command=self.selecionar_pasta_fotos,
        )
        self.btn_selecionar_pasta_fotos.pack(pady=5)

        self.btn_selecionar_planilha = Button(
            master,
            text="Selecionar Planilha de Pacotes",
            command=self.selecionar_planilha,
        )
        self.btn_selecionar_planilha.pack(pady=5)

        self.btn_selecionar_pasta_alunos = Button(
            master,
            text="Selecionar Pasta de Fotos dos Alunos",
            command=self.selecionar_pasta_alunos,
        )
        self.btn_selecionar_pasta_alunos.pack(pady=5)

        self.btn_processar = Button(
            master, text="Processar Fotos", command=self.processar_fotos
        )
        self.btn_processar.pack(pady=20)

        self.pasta_poses = None
        self.caminho_planilha = None
        self.pasta_alunos = None

    def selecionar_pasta_fotos(self):
        self.pasta_poses = filedialog.askdirectory(
            title="Selecione a pasta com as fotos das poses"
        )
        if not self.pasta_poses:
            messagebox.showwarning("Aviso", "Nenhuma pasta selecionada.")

    def selecionar_planilha(self):
        self.caminho_planilha = filedialog.askopenfilename(
            title="Selecione a planilha com os pacotes",
            filetypes=[("Excel files", "*.xlsx;*.xls")],
        )
        if not self.caminho_planilha:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada.")

    def selecionar_pasta_alunos(self):
        self.pasta_alunos = filedialog.askdirectory(
            title="Selecione a pasta com as fotos dos alunos"
        )
        if not self.pasta_alunos:
            messagebox.showwarning("Aviso", "Nenhuma pasta de alunos selecionada.")

    def processar_fotos(self):
        if not self.pasta_poses or not self.caminho_planilha or not self.pasta_alunos:
            messagebox.showwarning(
                "Aviso", "Por favor, selecione todas as pastas e a planilha."
            )
            return

        # Carregar a planilha com os pacotes
        alunos = carregar_planilha(self.caminho_planilha)

        # Listar pastas (poses) na pasta de fotos
        nomes_poses = listar_pastas(self.pasta_poses)

        # Contar fotos por pose
        contagem_fotos = contar_fotos_por_pose(self.pasta_poses, nomes_poses)

        # Contar fotos dos alunos, usando a pasta de poses corretamente
        # Contar fotos dos alunos, usando a pasta de poses corretamente
        contagem_alunos = contar_fotos_alunos(self.pasta_alunos, alunos, nomes_poses, self.pasta_poses)


        # Gerar a planilha de resultado
        resultado = gerar_planilha_resultado(alunos, contagem_fotos, contagem_alunos)

        # Salvar a nova planilha
        caminho_arquivo = "resultado_fotos.xlsx"
        resultado.to_excel(caminho_arquivo, index=False)

        # Adicionar somas e formatação condicional
        adicionar_somas(resultado, caminho_arquivo)

        messagebox.showinfo(
            "Sucesso",
            "Processamento concluído! O resultado foi salvo em 'resultado_fotos.xlsx'.",
        )


if __name__ == "__main__":
    root = Tk()
    app = Interface(root)
    root.mainloop()
