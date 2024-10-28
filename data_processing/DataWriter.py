import csv
from pathlib import Path

class DataWriter:

    def createFinalDict(self, emailData, txtData, mapping):
        finalDictionaries = []

        for dataDict in emailData:
            finalDict = dict(mapping)

            for key, value in dataDict.items():
                if key in finalDict:
                    finalDict[key] = value

            for key, value in txtData.items():
                if key in finalDict:
                    finalDict[key] = value

            finalDictionaries.append(finalDict)

        return finalDictionaries
    
    def writeCsv(self, data, mapping, serviceType):
        templatesDir = Path(__file__).parent.parent / 'csv_templates'
        templateFile = templatesDir / f"{serviceType.replace('.txt', '')}.csv"
        
        if not templateFile.exists():
            raise FileNotFoundError(f"Template CSV '{templateFile}' não encontrado.")
        
        with open(templateFile, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            header = next(reader)
        
        numColumns = len(header)
        
        with open(templateFile, 'a', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)

            for dataDict in data:
                row = [''] * numColumns

                for key, value in dataDict.items():
                    if key in mapping:
                        columnIndex = ord(mapping[key].upper()) - ord('A')
                        
                        if 0 <= columnIndex < numColumns:
                            row[columnIndex] = value
                        else:
                            print(f"Índice da coluna fora dos limites: {columnIndex}")
                
                writer.writerow(row)

        print(f"Dados gravados com sucesso no arquivo CSV:\n Caminho:'{templateFile}'.")
        
    def clearCsvFiles(self, csvPath):
        csv_dir = Path(csvPath)
        
        for csv_file in csv_dir.glob('*.csv'):
            with open(csv_file, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                header = next(reader, None)  

            with open(csv_file, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                if header:
                    writer.writerow(header)