import ctypes
import os 

class Display:
    def displayBanner():
        print("""
╔═════════════════════════════════════════════╗
║                                             ║
║          Tradutor E-Mails para CSV          ║
║             (Auto-Implantação)              ║
║          Por: Luiz Fernando Kuhn            ║
║                                      v1.10  ║
╚═════════════════════════════════════════════╝""")
    
    def displayMenu():
        print("""
╔═════════════════════════════════════════════╗
║             Seleção Layout CSV              ║
╠═════════════════════════════════════════════╣
║  1 - Pagamento (240)                        ║
║  2 - Extrato   (240)                        ║
║  3 - Cobrança  (240)                        ║
║  4 - Cobrança  (400)                        ║
║  5 - Varredura (240)                        ║
╠═════════════════════════════════════════════╣
║              Outros Serviços                ║
╠═════════════════════════════════════════════╣
║  10 - Template Properties                   ║
║  11 - Extrato Eletrônico                    ║
╚═════════════════════════════════════════════╝""")
        
    def displayWarning():
        print("""
╔═════════════════════════════════════════════╗
║                    AVISO                    ║
╠═════════════════════════════════════════════╣
║ # Para finalizar o programa aperte "Ctrl+C" ║
║                                             ║
║ # Antes de fechar os arquivos, salve-os     ║
║   utilizando "Ctrl+S"                       ║
║                                             ║
║ # Um "*" no nome da empresa indica que ela  ║
║   possui mais de 30 caracteres.             ║
╚═════════════════════════════════════════════╝""")
     
    def printMessageBox(string):
        boxWidth = 45
        p1 = (boxWidth - len(string)) // 2
        p2 = p1 + 1 if (boxWidth - len(string)) % 2 != 0 else p1
        
        print( "╔" + "═" * boxWidth + "╗")
        print( "║" + " " * p1 + f"{string}" + " " * p2 + "║")
        print( "╚" + "═" * boxWidth + "╝")

    def printFormattedDict(dicionario):
        if not dicionario:
            return

        print("╔═════════════════════════════════════════════╗")
        for idx, (chave, valor) in enumerate(dicionario.items()):
            pontos = '.' * (41 - len(chave))
            print(f"║ {chave.upper()}{pontos} {valor} ║")
        print("╚═════════════════════════════════════════════╝")