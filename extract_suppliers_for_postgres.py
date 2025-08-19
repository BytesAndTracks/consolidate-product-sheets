import pandas as pd
import unidecode
import re
import os
import numpy as np


EXCEL_FILE_PATH = r"C:\Users\Rubens\Downloads\Catálogos\MasterFornecedor_teste.xlsm"

OUTPUT_CSV_FILE = "suppliers_for_postgres.csv"

HEADER_ROW = 1


def clean_text(data):
    """
    Limpa e padroniza uma string: remove acentos, converte para maiúsculas,
    remove espaços extras e caracteres que podem quebrar o CSV.
    """
    if pd.isna(data):
        return np.nan
    
    data_str = str(data)
    
    cleaned_text = unidecode.unidecode(data_str).upper().strip()
    cleaned_text = re.sub(r'["\';\n\r]', '', cleaned_text)
    return cleaned_text


def main():
    print("Iniciando o processo de extração da aba de fornecedores...")

    try:
        df = pd.read_excel(
            EXCEL_FILE_PATH,
            sheet_name=0,
            header=HEADER_ROW,
            engine='openpyxl'
        )
        print(f"Sucesso! Aba '{pd.ExcelFile(EXCEL_FILE_PATH).sheet_names[0]}' lida com sucesso.")

    except FileNotFoundError:
        print(f"ERRO: O arquivo '{EXCEL_FILE_PATH}' não foi encontrado.")
        return
    except Exception as e:
        print(f"ERRO: Ocorreu um problema ao ler a planilha. Detalhes: {e}")
        return

    column_mapping = {
        'CÓD': 'code',
        'CATEGORIA FORN.': 'supplier_category',
        'PREENCHEU O FORMS': 'filled_form',
        'GRUPO WHATSAPP': 'whatsapp_group',
        'PEDIDO MÍNINO R$': 'minimum_order_value',
        'NOME DO FORNECEDOR': 'supplier_name',
        'SITE': 'website',
        'CNPJ': 'cnpj',
        'TELEFONE': 'phone',
        'ENDEREÇO COMPLETO': 'full_address',
        'NOME DO REPRESENTANTE': 'representative_name',
        'EMAIL DO REPRESENTANTE': 'representative_email',
        'TELEFONE DO REPRESENTANTE': 'representative_phone',
        'NOME DO GERENTE COMERCIAL': 'sales_manager_name',
        'EMAIL GERENTE COMERCIAL': 'sales_manager_email',
        'PODEMOS UTILIZAR A IMAGEM DE VOCÊS EM NOSSO SITE PARA PROMOVER NOSSA PARCERIA?': 'image_use_permission',
        'ENQUADRAMENTO TRIBUTÁRIO': 'tax_regime',
        'POSSUI ALGUM REGIME ESPECIAL ESTADUAL?': 'has_special_state_regime',
        'INDICAR A(S) CERTIFICAÇÃO(ÕES) DA QUALIDADE EXISTENTE(S)': 'quality_certifications',
        'OUTRO TIPO DE CERTIFICAÇÃO DA QUALIDADE: (DESCREVER)': 'other_quality_certifications',
        'SELECIONAR OS DOCUMENTOS APLICÁVEIS QUE SERÃO ENVIADOS PARA O E-MAIL COMPRAS@TREMED.COM.BR:': 'applicable_documents_sent',
        'TIPO DE ORÇAMENTO': 'quote_type',
        'FORMA DE PAGAMENTO? CONDIÇÕES DE PAGAMENTO? TEMPO ESTIMADO DE ENTREGA?': 'payment_and_delivery_terms',
        'EMAIL ORÇAMENTO': 'quote_email',
        'EMAIL FORMS': 'form_email',
        'DOC': 'doc_reference'
    }

    df.rename(columns=column_mapping, inplace=True)

    final_columns = [col for col in column_mapping.values() if col in df.columns]
    
    df = df[final_columns]
    
    print(f"Colunas selecionadas e renomeadas para o padrão do banco de dados.")

    if 'supplier_name' in df.columns:
        df.dropna(subset=['supplier_name'], inplace=True, how='all')
    else:
        print("ERRO CRÍTICO: A coluna 'NOME DO FORNECEDOR' não foi encontrada. Verifique o cabeçalho no Excel.")
        return

    print("Iniciando limpeza dos dados...")
    for col in df.columns:
        df[col] = df[col].apply(clean_text)
    print("Limpeza concluída.")

    try:
        df.to_csv(OUTPUT_CSV_FILE, sep=';', encoding='utf-8-sig', index=False, na_rep='NULL')
        
        caminho_completo = os.path.abspath(OUTPUT_CSV_FILE)
        print("\n---------------------------------------------------------")
        print("SUCESSO! O processo foi concluído.")
        print(f"O arquivo final foi salvo em: {caminho_completo}")
        print("---------------------------------------------------------")

    except Exception as e:
        print(f"\nERRO: Ocorreu um problema ao salvar o arquivo CSV: {e}")

if __name__ == '__main__':
    main()