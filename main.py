from tkinter import filedialog
from tkinter import messagebox
import time
import xlwings as xw
import cellMap
import messages

print(messages.inicio)

messagebox.showwarning(
    "Atenção!",
    messages.aviso,
)

# Caminhos do arquivo da PCI e RAE
PCI_path = filedialog.askopenfilename(
    title="SELECIONE O ARQUIVO: >>>>>     PCI      <<<<",
    filetypes=[("---> PCI <---", "*.xls;*.xlsx;*.xlsm")],
)
RAE_path = filedialog.askopenfilename(
    title="SELECIONE O ARQUIVO: >>>>>     RAE      <<<<",
    filetypes=[("RAE - VERSÃO 11-AGO-2023 !", "*.xlsm")],
)

if not PCI_path or not RAE_path:
    print("Planilhas não selecionadas!")
    time.sleep(3)
    exit()


print("Iniciando a leitura da PCI...")
PCI = xw.Book(PCI_path).sheets["Proposta_Constr_Individual"]

# Determina a versão da PCI
print("Determinando a versão da PCI...")
footer = PCI.api.PageSetup.LeftFooter
version = footer.split(" ")[1:][0]

PCI_Cells = {}

if version == "11/08/2023":
    PCI_Cells = cellMap.Cells_11_08_2023
else:
    PCI_Cells = cellMap.Cells_28_06_2022

print("Versão da PCI: " + version)

# Procura e determina o endereço correto das celula de 'incidência' e 'cronograma acumulado'

Inc_Pointer = ""
for i in range(10):
    if PCI.range("X" + str(90 + i)).value == "Incidência":
        Inc_Pointer = "X" + str(90 + i + 1)
PCI_Cells["Inidencia_Pointer"] = Inc_Pointer
Cro_Pointer = ""
Prazo_Pointer = ""
for i in range(10):
    if PCI.range("AK" + str(138 + i)).value == "Etapa":
        Cro_Pointer = "AO" + str(138 + i + 3)
        Prazo_Pointer = "AS" + str(138 + i - 1)
PCI_Cells["Crono_Pointer"] = Cro_Pointer
Prazo = int(PCI.range(Prazo_Pointer).value)

print("Abrindo arquivo do RAE... Clique em 'Fim' na mensagem de erro!")
RAE = xw.Book(RAE_path).sheets["RAE"]

print("Iniciando cópia da PCI para RAE...")

# Copiando dados iniciais
for campo, cellRae in cellMap.Rae_11AGO2023.items():
    RAE.range(cellRae).value = PCI.range(PCI_Cells[campo]).value
    print(RAE.range(cellRae).value)

# INCIDÊNCIAS
IN_numInicio = int(PCI_Cells["Inidencia_Pointer"][1:])
IN_letraInicio = PCI_Cells["Inidencia_Pointer"][:1]
print("<--- Incidências Inicio --->")
for i in range(20):
    RAE.range("S" + str(68 + i)).value = PCI.range(
        IN_letraInicio + str(IN_numInicio + i)
    ).value
    print(PCI.range(IN_letraInicio + str(IN_numInicio + i)).value)
print("<--- Incidências Fim --->")

# CRONOGRAMA PREVISTO ACUMULADO
CRO_numInicio = int(PCI_Cells["Crono_Pointer"][2:])
CRO_letraInicio = PCI_Cells["Crono_Pointer"][:2]
print("Copiando cronograma acumulado. Prazo = " + str(Prazo) + " meses")
for i in range(Prazo):
    RAE.range("AG" + str(72 + i)).value = PCI.range(
        CRO_letraInicio + str(CRO_numInicio + i)
    ).value
    print(RAE.range("AG" + str(72 + i)).value)

input("Cópia Finalizada! Pode fechar essa janela...")
