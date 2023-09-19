import xmltodict
import os
import pandas as pd


def extract_xml_data(file_path):
    """
    Extrai informações de um arquivo XML e as retorna como um dicionário.

    Args:
        file_path (str): O caminho para o arquivo XML.

    Returns:
        dict: Um dicionário contendo informações extraídas.
    """
    with open(file_path, "rb") as xml_file:
        xml_data = xml_file.read()
        return xmltodict.parse(xml_data)


def process_xml_files(directory):
    """
    Processa todos os arquivos XML no diretório fornecido e retorne uma lista de dicionários.

    Args:
        diretório (str): O diretório que contém arquivos XML.

    Returns:
        list: Uma lista de dicionários, cada um contendo informações extraídas de um arquivo XML.

    """
    xml_data_list = []

    for file_name in os.listdir(directory):
        file_path = os.path.join(directory, file_name)
        if file_name.endswith(".xml") and os.path.isfile(file_path):
            xml_data = extract_xml_data(file_path)
            xml_data_list.append(xml_data)

    return xml_data_list


def main():
    input_directory = "nfs"
    output_file = "NotasFiscais.xlsx"
    columns = ["numero_nota", "emissor_nota", "nome_cliente", "endereco", "peso"]

    xml_data_list = process_xml_files(input_directory)
    extracted_data = []

    for xml_data in xml_data_list:
        if "NFe" in xml_data:
            info_nf = xml_data["NFe"]["infNFe"]
        else:
            info_nf = xml_data["nfeProc"]["NFe"]["infNFe"]

        numero_nota = info_nf["@Id"]
        emissor_nota = info_nf["emit"]["xNome"]
        nome_cliente = info_nf["dest"]["xNome"]
        endereco = info_nf["dest"]["enderDest"]
        peso = (
            info_nf["transp"]["vol"]["pesoB"]
            if "vol" in info_nf["transp"]
            else "Não informado"
        )

        extracted_data.append([numero_nota, emissor_nota, nome_cliente, endereco, peso])

    df = pd.DataFrame(columns=columns, data=extracted_data)
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    main()
