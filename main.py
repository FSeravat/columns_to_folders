import os
import pandas
import re
import PySimpleGUI as sg

class Window:
    def __init__(self):
        layout = [  
                    [sg.Text("Arquivo Excel: ", size=(15), justification="right"),sg.InputText(key="excelFile"),sg.FileBrowse()],
                    [sg.Text("Modelador: ", size=(15), justification="right"),sg.InputText(key="modelador")],
                    [sg.Text("Situação: ", size=(15), justification="right"),sg.InputText(key="situação")],
                    [sg.Text("Caminho do output: ", size=(15), justification="right"),sg.InputText(key="outputPath"),sg.FolderBrowse()],
                    [sg.Push(), sg.Column([[sg.Button("Confirmar", key="BtnOk")]],element_justification='c'), sg.Push()]
                ]

        self.window = sg.Window('Criador de pastas GASLUB', layout)

    def open(self):
        while True:
            event, values = self.window.read()
            if event == sg.WIN_CLOSED:
                break
            if event == "BtnOk":
                try:
                    self.createFolders(values["excelFile"], values["outputPath"], values["modelador"], values["situação"])
                    sg.popup("PASTAS CRIADAS COM SUCESSO")
                except Exception as e:
                    sg.popup(str(e))

    def close(self):
        self.window.close()

    def createFolders(self, file, output, modelador, situacao):
        df = pandas.read_excel(file, sheet_name="Controle",header=1)
        data = []
        for index, row in df.iterrows():
            if row["MODELADOR"] == modelador and row["SITUAÇÃO"]==situacao:
                data.append(row)

        for x in data:
            tag = re.sub(r'["./]',"",x["TAG'S"])
            if not os.path.exists(os.path.join(output,tag)):
                os.mkdir(os.path.join(output,tag))

Window().open()