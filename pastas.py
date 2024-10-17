import os

def listar_pastas(pasta):
    # Listar todos os itens na pasta
    itens = os.listdir(pasta)
    # Filtrar apenas as pastas
    pastas = [item for item in itens if os.path.isdir(os.path.join(pasta, item))]
    return pastas

# Exemplo de uso
caminho_pasta = './pastas'
lista_pastas = listar_pastas(caminho_pasta)
print(lista_pastas)
