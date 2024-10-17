# Sistema de Nomeação e Classificação de Fotos

Este projeto é um sistema automatizado para nomeação, contagem, e classificação de fotos. Ele processa fotos selecionadas por clientes e organiza-as em uma planilha Excel, contabilizando fotos por pose e determinando a quantidade de fotos extras selecionadas.

## Funcionalidades

- Carrega uma planilha inicial com nomes dos clientes e a quantidade de fotos em seu pacote.
- Lê pastas de fotos organizadas por poses.
- Conta quantas fotos de cada pose foram selecionadas por cada cliente.
- Calcula o total de fotos selecionadas e fotos extras (fora do pacote contratado).
- Gera uma planilha Excel com a contagem detalhada por pose e soma total, incluindo formatação condicional de cores para destacar as poses mais escolhidas.

## Como Usar

1. **Pré-requisitos**

   - Python 3.7 ou superior
   - Instale as dependências utilizando o comando:
     ```bash
     pip install -r requirements.txt
     ```
2. **Executar o Projeto**

   - Execute o arquivo `main.py` para iniciar o processo:
     ```bash
     python main.py
     ```
3. **Fluxo do Programa**

   - Selecione a planilha com os pacotes dos clientes.
   - Selecione a pasta com as fotos organizadas por poses.
   - Selecione a pasta com as fotos selecionadas pelos alunos.
   - O programa irá gerar uma nova planilha com as informações processadas, incluindo o total de fotos e fotos extras por cliente.
4. **Formatação Condicional**

   - A linha que contém a soma total de fotos por pose recebe uma formatação de cores verdes, onde o menor valor recebe um tom mais claro, e o maior valor, um tom mais forte.

## Estrutura do Projeto

```plaintext
📁 contador-auto
 ├── main.py                # Código principal do sistema
 ├── requirements.txt        # Dependências do projeto
 └── resultado_fotos.xlsx    # Planilha gerada com os resultados
```
