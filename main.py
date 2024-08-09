import xmltodict

# abrir e ler o arquivo

def ler_xml_danfe(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    # print(documento)
    dic_notafiscal = documento['nfeProc']['NFe']['infNFe']

    valor_total = dic_notafiscal['total']['ICMSTot']['vNF']
    cnpj_vendeu = dic_notafiscal['emit']['CNPJ']
    nome_vendeu = dic_notafiscal['emit']['xNome']
    cpf_comprou = dic_notafiscal['dest']['CPF']
    nome_comprou = dic_notafiscal['dest']['xNome']
    produtos = dic_notafiscal['det']
    lista_produtos = []
    for produto in produtos:
        valor_produto = produto['prod']['vProd']
        nome_produto = produto['prod']['xProd']
        lista_produtos.append((nome_produto, valor_produto))
    resposta = {
        'nome_vendeu': [nome_vendeu],
        'cnpj_vendeu': [cnpj_vendeu],
        'nome_comprou': [nome_comprou],
        'cpf_comprou': [cpf_comprou],
        'lista_produtos': [lista_produtos],
        'valor_total': [valor_total],
    }
    return resposta


def ler_xml_servico(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    # print(documento)
    dic_notafiscal = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']

    valor_total = dic_notafiscal['Servico']['Valores']['ValorServicos']
    cnpj_vendeu = dic_notafiscal['PrestadorServico']['IdentificacaoPrestador']['Cnpj']
    nome_vendeu = dic_notafiscal['PrestadorServico']['RazaoSocial']
    cpf_comprou = dic_notafiscal['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    nome_comprou = dic_notafiscal['TomadorServico']['RazaoSocial']
    produtos = dic_notafiscal['Servico']['Discriminacao']
    resposta = {
        'nome_vendeu': [nome_vendeu],
        'cnpj_vendeu': [cnpj_vendeu],
        'nome_comprou': [nome_comprou],
        'cpf_comprou': [cpf_comprou],
        'lista_produtos': [produtos],
        'valor_total': [valor_total],
    }
    return resposta

import os

caminho = os.getcwd()
pasta = caminho + r'\NFs Finais'
lista_arquivos = os.listdir(pasta) # lista os nomes dos arquivos de uma pasta

# caso queria adicionar todas as NFs em um mesmo arquivo Excel:
import pandas as pd

df_final = pd.DataFrame()
dfs =[]
for arquivo in lista_arquivos:
    if arquivo.endswith('.xml'): # verifica entensão do arquivo se é XML
        if 'DANFE' in arquivo.upper():
            df = pd.DataFrame.from_dict(ler_xml_danfe(f'{pasta}/{arquivo}'))
        else:
            df = pd.DataFrame.from_dict(ler_xml_servico(f'{pasta}/{arquivo}'))
        dfs.append(df)
        
df_final = pd.concat(dfs, ignore_index= True)

# salvando o arquivo em formato Excel
df_final.to_excel(f'{pasta}/Notas Fiscais.xlsx', index=False)
