import xmltodict
import os
import json
import pandas as pd


def pegar_infos(nome_arquivo, valores):
    # print(f"Pegou as informações{nome_arquivo}")
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        try:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
            numero_nota = infos_nf["ide"]["nNF"]
            serie_nota = infos_nf["ide"]["serie"]
            empresa_emissora = infos_nf["emit"]["xNome"]
            cnpj = infos_nf["emit"]["CNPJ"]
            nome_cliente = infos_nf["dest"]["xNome"]
            endereco = infos_nf["dest"]["enderDest"]
            peso = infos_nf["transp"]["vol"]["pesoB"]
            imposto = infos_nf["det"]["imposto"]["ICMS"]["ICMSSN102"]["CSOSN"]
            data_emissao = infos_nf["ide"]["dhEmi"]
            valores.append([
            numero_nota,
            serie_nota,
            empresa_emissora,
            cnpj,
            nome_cliente,
            endereco,
            peso,
            imposto,
            data_emissao])
        except Exception as e:
            print(e)
            print(json.dumps(dic_arquivo, indent=4))
lista_arquivos = os.listdir("nfs")

colunas = [" numero_nota", "serie_nota", "empresa_emissora", "cnpj", "nome_cliente", "endereco", "peso", "imposto", "data_emissao" ]
valores = []
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)
tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)