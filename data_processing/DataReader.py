import pandas as pd
import re
from pathlib import Path

class DataReader:
    
    # Método para Ler a Planilha
    def readSheet(self, sheetPath):
        df = pd.read_excel(sheetPath)

        cnpjPattern = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'
        fullAgencyAccountDvPattern = r'(\d{4,})/(\d{5,})-(\d)'
        agencyPattern = r'\b\d{4}\b'
        accountDvPattern = r'\b(\d{5})-(\d)\b'
        companyNamePattern = r'(?i)raz[aã]o social[:\-]?\s*(.*)'
        nicknamePattern = r'\.([^\.]+)\.341'
        companySuffixesPattern = r'\b(EIRELI|ME|LTDA|SA|EPP|CNPJ)\b'

        returnColIndex = self.findReturnColumnIndex(df.columns)

        allData = []

        for _, row in df.iterrows():
            firstColumn = row.iloc[0]
            cnpj = self.extractCnpj(firstColumn, cnpjPattern)
            agency, account, dv = self.extractAgencyAccountDv(firstColumn, fullAgencyAccountDvPattern, agencyPattern, accountDvPattern)
            companyName = self.extractCompanyName(firstColumn, cnpj, agencyPattern, accountDvPattern, companyNamePattern, companySuffixesPattern)
            returnField = row.iloc[returnColIndex]
            nickname = self.extractNickname(returnField, nicknamePattern)
            remittance, return_ = row.iloc[1], row.iloc[2]

            dataDict = {
                'cnpj': cnpj.strip().zfill(14) if cnpj else "N/A",
                'empresa': companyName.strip(),
                'agencia': agency if agency else "N/A",
                'conta': account if account else "N/A",
                'dv': dv if dv else "N/A",
                'dsn_remessa': remittance,
                'dsn_retorno': return_.strip(),
                'apelido': nickname.strip(),
            }

            allData.append(dataDict)

        return allData

    # Busca a coluna do Retorno
    def findReturnColumnIndex(self, columns):
        for index, columnName in enumerate(columns):
            if columnName.strip().lower() == "retorno":
                return index
        raise ValueError("Coluna 'retorno' não encontrada.")

    # Extrai o CNPJ
    def extractCnpj(self, firstColumn, cnpjPattern):
        cnpjMatch = re.search(cnpjPattern, firstColumn)
        if cnpjMatch:
            return cnpjMatch.group().replace('.', '').replace('/', '').replace('-', '')
        return None

    # Extrai Agência, Conta e DV, tentando o formato completo 0000/00000-0 e depois separadamente
    def extractAgencyAccountDv(self, firstColumn, fullAgencyAccountDvPattern, agencyPattern, accountDvPattern):
        # Primeiro tenta buscar o padrão completo
        fullMatch = re.search(fullAgencyAccountDvPattern, firstColumn)
        if fullMatch:
            return fullMatch.groups()  # Retorna agência, conta e DV

        # Se não encontrar o padrão completo, busca a agência e a conta separadamente
        agencyMatch = re.search(agencyPattern, firstColumn)
        accountDvMatch = re.search(accountDvPattern, firstColumn)

        # Retorna os valores encontrados ou None se não encontrados
        agency = agencyMatch.group() if agencyMatch else None
        account, dv = accountDvMatch.groups() if accountDvMatch else (None, None)

        return agency, account, dv

    # Extrai o Nome da Empresa
    def extractCompanyName(self, firstColumn, cnpj, agencyPattern, accountDvPattern, companyNamePattern, companySuffixesPattern):
        companyNameMatch = re.search(companyNamePattern, firstColumn)
        if companyNameMatch:
            companyName = companyNameMatch.group(1)
        else:
            companyName = firstColumn
            if cnpj:
                companyName = re.sub(cnpj, '', companyName)
            companyName = re.sub(agencyPattern, '', companyName)
            companyName = re.sub(accountDvPattern, '', companyName)
        
        # Remover caracteres especiais como -, /, .
        companyName = re.sub(r'[-/.]', '', companyName).strip()

        # Remover números antes do primeiro caractere alfabético
        companyName = re.sub(r'^[^a-zA-Z]*', '', companyName)

        # Remover números após o último caractere alfabético
        companyName = re.sub(r'[^a-zA-Z]*$', '', companyName)
        
        # Remover sufixos como EIRELI, ME, LTDA, etc.
        companyName = re.sub(companySuffixesPattern, '', companyName)

        # Limitar a 30 caracteres e marcar com '*' se exceder
        if len(companyName) > 30:
            companyName = companyName[:30] + '*' + companyName[30:]

        return companyName.strip()

    # Extrai o Apelido
    def extractNickname(self, returnField, nicknamePattern):
        nicknameMatch = re.search(nicknamePattern, returnField, re.IGNORECASE)
        return nicknameMatch.group(1).strip() if nicknameMatch else "N/A"

    def readTxt(self, txtPath):
        dictionary = {}
        with open(txtPath, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if ':' in line:
                    key, value = line.split(':', 1)
                    dictionary[key.strip()] = value.strip()
        return dictionary