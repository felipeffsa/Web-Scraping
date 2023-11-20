import os
import openpyxl


class FileOpenpy:
    

    def openDiretorio(self, nomearquivo):
   
        # Obtém o diretório da área de trabalho
        self.diretorio_area_trabalho = os.path.join(os.path.expanduser("~"), "Desktop")
        # Define o nome do arquivo
        self.nome_arquivo = f"{nomearquivo}.xlsx"
        # Cria o caminho completo para o arquivo
        self.caminho_arquivo = os.path.normpath(os.path.join(
            self.diretorio_area_trabalho, self.nome_arquivo))
        # Salva o workbook no caminho especificado
        self.arquivo.save(self.caminho_arquivo)   

    def criar(self):
        self.arquivo = openpyxl.Workbook()
        self.planilha = self.arquivo.active
        self.planilha["A1"] = "Nome"
        self.planilha["B1"] = "Estrelas"
        self.planilha["C1"] = "Preço"
    
    def openpy_laço(self, numero, nome, estrela, preco):
        
        self.planilha[f"A{numero}"] = nome.text
        self.planilha[f"B{numero}"] = estrela.text
        self.planilha[f"C{numero}"] = preco.text


    def salvar(self, arquivo):
        self.arquivo.save(arquivo)


