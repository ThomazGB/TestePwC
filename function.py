import xlsxwriter

# Função para casos simples
def separar_casos_simples(endereco):
    partes_endereco = endereco.split()
    return partes_endereco

# Função para casos complicados
def separar_casos_complicados(endereco):
    partes_endereco = endereco.split()
    numero = ""
    partes_rua = []

    
    for i, parte in enumerate(partes_endereco):
        if parte.isdigit():
            # Assim que encontrar um dígito, adicione esse dígito e todas as partes seguintes ao número
            numero += ' '.join(partes_endereco[i:])  # Junta todas as partes após o número
            break  # Encerra o loop, já que todos os itens restantes foram adicionados ao número
        else:
            partes_rua.append(parte)  # Adiciona à rua

    rua = ' '.join(partes_rua).strip()
    return [rua, numero.strip()]

# Função para casos internacionais
def separar_casos_complexos(endereco):
    partes_endereco = endereco.replace(',', '').split()
    numero = ""
    partes_rua = []

    i = 0
    while i < len(partes_endereco):
        parte = partes_endereco[i]

        # Verifica se a parte é "No" ou "NO"
        if parte.upper() == "NO":
            # Adiciona todas as partes após o "No" ao número
            numero += parte + " "
            numero += ' '.join(partes_endereco[i + 1:])  # Adiciona o resto da string ao número
            break  # Encerra o loop após processar o "No"
        # Se for um número isolado, considera como número do endereço
        if parte.isdigit() and "No" not in partes_endereco:
            numero += parte + " "
        else:
            partes_rua.append(parte)

        i += 1

    rua = ' '.join(partes_rua).strip()
    return [rua, numero.strip()]

#Exportar para XLSX
enderecosSimples = ["Miritiba 339", "Babaçu 500", "Cambuí 804B"]
enderecosComplicados = ["Rio Branco 23", "Quirino dos Santos 23 b"]
enderecosComplexos = ["4, Rue de la République", "100 Broadway Av", "Calle Sagasta, 26", "Calle 44 No 1991"]

workbook = xlsxwriter.Workbook("../TestePwC.xlsx")
worksheet = workbook.add_worksheet("testSheet")

worksheet.write(0, 0, "Casos Simples")
worksheet.write(0, 1, "Rua")
worksheet.write(0, 2, "Número")
worksheet.write(0, 3, "Casos Complicados")
worksheet.write(0, 4, "Rua")
worksheet.write(0, 5, "Número")
worksheet.write(0, 6, "Casos Complexos")
worksheet.write(0, 7, "Rua")
worksheet.write(0, 8, "Número")

for index, entry in enumerate(enderecosSimples):    
    rua, numero = separar_casos_simples(entry)
    worksheet.write(index+1, 0, index+1)
    worksheet.write(index+1, 1, rua)
    worksheet.write(index+1, 2, numero)
    
for index, entry in enumerate(enderecosComplicados):    
    rua, numero = separar_casos_complicados(entry)
    worksheet.write(index+1, 3, index+1)
    worksheet.write(index+1, 4, rua)
    worksheet.write(index+1, 5, numero)

for index, entry in enumerate(enderecosComplexos):    
    rua, numero = separar_casos_complexos(entry)
    worksheet.write(index+1, 6, index+1)
    worksheet.write(index+1, 7, rua)
    worksheet.write(index+1, 8, numero)
    
workbook.close()    