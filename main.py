from data_processing.DataReader import DataReader
from data_processing.DataWriter import DataWriter
from view.Display import Display
from pathlib import Path
import subprocess

dataReader = DataReader()
dataWriter = DataWriter()

"""
EXT240 = fb0db840-6e0e-4d3e-872f-f1ce5177d064
COB240 = 1861311e-17aa-4673-b8a2-18d0b52bb392
COB400 = 27f18568-442d-4635-82e2-745c1c9edc55
PAG240 = 23568aab-40f5-417d-876a-dd65a30a1a90"""

def selectService():
    """Exibe o menu e retorna o tipo de serviço selecionado."""
    service_options = {
        1: "PAG240.txt",
        2: "EXT240.txt",
        3: "COB240.txt",
        4: "COB400.txt",
        5: "VAR240.txt",
        10: "PROPERTIES.txt",
        11: "EXTRATOELET.txt"
    }

    while True:
        try:
            invalidOptionError = "Opção inválida. Por favor, insira um valor numérico inteiro."
            Display.displayMenu()
            serviceType = int(input("\n\nDigite o número do serviço desejado: "))
            
            if serviceType in service_options:
                return service_options[serviceType]
            else:
                print(invalidOptionError)
        except ValueError:
            print(invalidOptionError)
            
def selectedTemplateKey(serviceType):
    service_map = {
        "PAG240.txt": {"template_key": "23568aab-40f5-417d-876a-dd65a30a1a90"},
        "EXT240.txt": {"template_key": "fb0db840-6e0e-4d3e-872f-f1ce5177d064"},
        "COB240.txt": {"template_key": "1861311e-17aa-4673-b8a2-18d0b52bb392"},
        "COB400.txt": {"template_key": "27f18568-442d-4635-82e2-745c1c9edc55"},
        "VAR240.txt": {"template_key": "a8e06106-1c98-4767-a24b-7c107c1affa0"},
        "PROPERTIES.txt":{"template_key": "null"},
        "EXTRATOELET.txt":{"template_key": "null"}
    }

    service_data = service_map.get(serviceType)
    if service_data:
        return {"template_key": service_data["template_key"]}
    else:
        print("ERRO: Template Key não encontrado!")


def loadTemplateMapping(scriptDir, serviceType):
    """Carrega e combina o mapeamento do template padrão e o selecionado."""
    if serviceType == "PROPERTIES.txt":
        return dataReader.readTxt(scriptDir / "template_mapping" / serviceType)
    
    if serviceType == "EXTRATOELET.txt":
        return dataReader.readTxt(scriptDir / "template_mapping" / serviceType)
    
    defaultMapping = dataReader.readTxt(scriptDir / "template_mapping" / "default.txt")
    serviceMapping = dataReader.readTxt(scriptDir / "template_mapping" / serviceType)
    
    return defaultMapping | serviceMapping

def processSpreadsheet():
    """Abre a planilha para processamento."""
    subprocess.run(["start", "input_email.xlsx"], shell=True)
    input("\nApós fechar a planilha, pressione 'ENTER' para continuar...")
    sheetData = dataReader.readSheet("input_email.xlsx")
    return sheetData

def processTxt():
    """Abre o arquivo de texto e processa o conteúdo."""
    subprocess.run(["notepad.exe", "input_manual.txt"])
    txtData = dataReader.readTxt("input_manual.txt")
    return txtData

def openSpreadsheet(filePath):
    """Abre a planilha associada ao arquivo de serviço."""
    input("\nPressione 'ENTER' para visualizar os dados escritos...")
    filePath = filePath.replace(".txt", ".csv")
    filePath = f"csv_templates\\{filePath}"
    subprocess.run(["start", filePath], shell=True)

def printDictionaries(dictionaries):
    """Recebe uma lista de dicionários e imprime cada um com um contador."""
    maxKeyLength = max(len(key) for dic in dictionaries for key in dic.keys())

    for i, dic in enumerate(dictionaries, start=1):
        print(f"Dicionário {i}:")
        for key, value in dic.items():
            points = '.' * (maxKeyLength - len(key) + 5)
            print(f"  {key.upper()}{points}: {value}")
        print()

def main():
    continuar = True
    while continuar:
        # Início
        Display.displayBanner()
        Display.displayWarning()
        
        scriptDir = Path(__file__).parent

        # Serviço Selecionado
        serviceType = selectService()
        
        # Mostra o mapeamento do serviço selecionado
        Display.printMessageBox(f"MAPEAMENTO DE TABELA SELECIONADO: {serviceType.replace('.txt', '')}")
        mapping = loadTemplateMapping(scriptDir, serviceType)
        Display.printFormattedDict(mapping)
        
        # Transforma os inputs em dicionários
        emailData = processSpreadsheet()
        txtData =  selectedTemplateKey(serviceType) | processTxt() 
        
        # Mostra todos os dicionários que foram gravados
        Display.printMessageBox("DADOS AGRUPADOS")
        finalDictionaries = dataWriter.createFinalDict(emailData, txtData, mapping)
        printDictionaries(finalDictionaries)
        
        # Escreve na planilha
        dataWriter.writeCsv(finalDictionaries, mapping, serviceType)
        openSpreadsheet(serviceType)
        
        # Limpa a planilha que foi escrita
        input("Pressione 'ENTER' para limpar a planilha...")
        csvPath = Path(scriptDir) / 'csv_templates'
        dataWriter.clearCsvFiles(csvPath)
        option = input(f"Planilha limpa com sucesso!\n\nDeseja realizar outra operação? (S/N): ")
        
        continuar = True if "S" == option.upper() else False
    
if __name__ == "__main__":
    main()
