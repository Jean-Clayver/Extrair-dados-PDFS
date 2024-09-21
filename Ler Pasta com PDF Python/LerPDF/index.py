import os
import tabula
import re
import pandas as pd  # Para criar o arquivo Excel

def extrair_candidatos_e_telefones_em_pdfs(diretorio_pdf):
    # Expressão regular ajustada para identificar telefones (com ou sem parênteses, espaços e hífens opcionais)
    padrao_telefone = re.compile(r'\d{9}')  # Telefone no formato brasileiro
    # Expressão regular para identificar candidatos (com conectores como 'de', 'da', 'dos')
    padrao_candidato = re.compile(r'[A-ZÁÉÍÓÚÂÊÔÃÕÇ]+(?: [a-z]{1,2})?(?: [A-ZÁÉÍÓÚÂÊÔÃÕÇ]+)+')

    # Lista para armazenar os resultados
    dados = []

    # Itera sobre cada arquivo PDF no diretório
    for arquivo in os.listdir(diretorio_pdf):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio_pdf, arquivo)
            print(f"Lendo arquivo: {arquivo}")
            
            # Extrai as tabelas de todas as páginas do PDF
            try:
                lista_tabelas = tabula.read_pdf(caminho_pdf, pages="all", multiple_tables=True)
                
                # Para armazenar as informações concatenadas de cada PDF
                candidatos_pdf = []
                telefones_pdf = []

                # Itera sobre as tabelas extraídas
                for i, tabela in enumerate(lista_tabelas):
                    tabela_str = tabela.to_string()  # Converte a tabela para string
                    telefones_encontrados = padrao_telefone.findall(tabela_str)
                    candidatos_encontrados = padrao_candidato.findall(tabela_str)

                    # Debug para verificar se encontra candidatos e telefones
                    print(f"Telefones encontrados na página {i + 1}: {telefones_encontrados}")
                    print(f"Candidatos encontrados na página {i + 1}: {candidatos_encontrados}")
                    
                    # Armazena os candidatos e telefones encontrados separadamente
                    if candidatos_encontrados:
                        candidatos_pdf.extend(candidatos_encontrados)
                    if telefones_encontrados:
                        telefones_pdf.extend(telefones_encontrados)
                
                # Adiciona os dados separados de todos os PDFs
                if candidatos_pdf or telefones_pdf:
                    dados.append({
                        'Arquivo': arquivo,
                        'Candidatos': ', '.join(candidatos_pdf) if candidatos_pdf else 'Nenhum candidato encontrado',
                        'Telefones': ', '.join(telefones_pdf) if telefones_pdf else 'Nenhum telefone encontrado'
                    })
            
            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")
    
    # Criar um DataFrame com os dados extraídos
    df = pd.DataFrame(dados, columns=['Arquivo', 'Candidatos', 'Telefones'])
    
    # Exibir o conteúdo da planilha no terminal
    print("\nConteúdo da planilha:")
    print(df)
    
    # Salvar o DataFrame em um arquivo Excel
    df.to_excel("Dados_Drap.xlsx", index=False)
    print("Planilha Excel criada com sucesso!")

# Caminho da pasta com PDFs
diretorio_pdf = r"C:\Users\jeanc\Desktop\Ler PDF Python\Drap"

# Chama a função para ler todos os PDFs na pasta e salvar os resultados em uma planilha
extrair_candidatos_e_telefones_em_pdfs(diretorio_pdf)
