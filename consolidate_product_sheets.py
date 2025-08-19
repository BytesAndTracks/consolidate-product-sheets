import argparse
import logging
import os
import re

import pandas as pd
import unidecode


EXCEL_FILE_PATH = r"C:\Users\Rubens\Downloads\Catálogos\MasterFornecedor_teste.xlsm"
OUTPUT_CSV_FILE = "dados_finais_prontos.csv"

EXPECTED_COLS = {'PRODUTO', 'MARCA', 'COD. INTERNO', 'NCM', 'PREÇO FORNECEDOR'}
EXPECTED_COLS_NORMALIZED = {unidecode.unidecode(x).upper().strip() for x in EXPECTED_COLS}


def clean_text(data):
    if not isinstance(data, str):
        return data
    cleaned_text = unidecode.unidecode(data).upper().strip()
    cleaned_text = re.sub(r'["\';\n]', '', cleaned_text)
    return cleaned_text

def clean_numeric(data):
    """
    Limpa e converte uma string para um formato numérico,
    tratando corretamente o formato brasileiro (ex: 1.234,50).
    """
    if isinstance(data, (int, float)):
        return data if pd.notna(data) else pd.NA

    if data is None:
        return pd.NA

    if isinstance(data, str):
        s = data.strip()
        if s == '':
            return pd.NA

        if '.' in s and ',' in s:
            s = s.replace('.', '').replace(',', '.')
        else:
            if ',' in s and not '.' in s:
                s = s.replace(',', '.')

        return pd.to_numeric(s, errors='coerce')

    return pd.to_numeric(data, errors='coerce')


def find_header_row(df_preview):
    for index, row in df_preview.iterrows():
        row_values = {unidecode.unidecode(str(v)).upper().strip() for v in row.dropna()}
        if len(EXPECTED_COLS_NORMALIZED.intersection(row_values)) >= 3:
            return index
    return None


def normalize_col_name(s):
    if not isinstance(s, str):
        return str(s)
    return unidecode.unidecode(s).upper().strip()


def sanitize_column_to_field(s):
    """Return a safe snake_case, lower-case name for unknown columns."""
    s = unidecode.unidecode(str(s)).strip()
    s = re.sub(r"[^0-9A-Za-z]+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip('_').lower()


def process_excel(input_path, output_csv, markup=1.55, preview_rows=20):
    logging.info("Abrindo arquivo Excel: %s", input_path)
    try:
        xls = pd.ExcelFile(input_path, engine='openpyxl')
    except FileNotFoundError:
        logging.error("Arquivo não encontrado: %s", input_path)
        raise

    column_mapping = {
        'CATEGORIA': 'categoria', 'SUB/CATEGORIA': 'subcategoria',
        'AREA DE ATUACAO': 'area_de_atuacao', 'PRODUTO': 'produto', 'NCM': 'ncm',
        'IPI': 'ipi', 'ICMS': 'icms', 'IMPORTADO OU NACIONAL': 'origem',
        'COD. INTERNO': 'cod_interno', 'COD. FORNECEDOR': 'cod_fornecedor',
        'ANVISA': 'anvisa', 'PRECO FORNECEDOR': 'preco_fornecedor',
        'PRECO UNITARIO': 'preco_unitario',
        'PRECO UNITARIO VENDA': 'preco_unitario_venda',
        'PRECO VENDA': 'preco_venda',
        'UNIDADE DE MEDIDA': 'unidade_medida', 'QTD': 'qtd',
        'PRECO VENDA EMBALAGEM': 'preco_venda_embalagem', 'FORNECEDOR': 'fornecedor',
        'MARCA': 'marca', 'OBSERVACAO': 'observacao'
    }

    column_mapping_normalized = {normalize_col_name(k): v for k, v in column_mapping.items()}

    collected = []
    for sheet_name in xls.sheet_names:
        logging.info("Verificando aba: %s", sheet_name)
        try:
            df_preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=preview_rows)
        except Exception as e:
            logging.warning("Não foi possível ler preview da aba %s: %s", sheet_name, e)
            continue

        header_row_index = find_header_row(df_preview)
        if header_row_index is None:
            logging.info("Pulando aba %s - cabeçalho não encontrado", sheet_name)
            continue

        logging.info("Processando aba '%s' (cabeçalho linha %s)", sheet_name, header_row_index + 1)
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_index)
        except Exception as e:
            logging.warning("Erro lendo aba %s: %s", sheet_name, e)
            continue

        orig_cols = list(df.columns)
        new_cols = []
        for c in orig_cols:
            normalized = normalize_col_name(c)
            if normalized in column_mapping_normalized:
                new_cols.append(column_mapping_normalized[normalized])
            else:
                new_cols.append(sanitize_column_to_field(c))

        df.columns = new_cols
        if 'fornecedor' not in df.columns:
            df['fornecedor'] = sheet_name

        if 'produto' not in df.columns:
            logging.info("Aba %s ignorada: coluna 'produto' não encontrada após normalização", sheet_name)
            continue

        collected.append(df)

    if not collected:
        logging.info("Nenhuma aba válida encontrada no arquivo.")
        return None

    master_df = pd.concat(collected, ignore_index=True)
    logging.info("Consolidação finalizada. Linhas totais: %s", len(master_df))

    numeric_cols = ['ipi', 'icms', 'preco_fornecedor', 'preco_unitario',
                    'preco_unitario_venda', 'preco_venda', 'qtd', 'preco_venda_embalagem']

    if 'preco_unitario' in master_df.columns:
        logging.info("Limpando 'preco_unitario' (fonte da verdade)")
        master_df['preco_unitario'] = master_df['preco_unitario'].apply(clean_numeric)
    else:
        logging.info("Coluna 'preco_unitario' não encontrada; será criada se possível")

    logging.info("Recalculando preços de venda com multiplicador: %s", markup)
    if 'preco_unitario' in master_df.columns:
        master_df['preco_unitario_venda'] = master_df['preco_unitario'] * markup
        master_df['preco_venda'] = master_df['preco_unitario'] * markup

    logging.info("Limpando demais colunas numéricas")
    for col in numeric_cols:
        if col in master_df.columns and col not in ['preco_unitario', 'preco_unitario_venda', 'preco_venda']:
            master_df[col] = master_df[col].apply(clean_numeric)

    logging.info("Limpando colunas de texto")
    text_cols = master_df.select_dtypes(include=['object']).columns.tolist()
    for col in text_cols:
        if col not in numeric_cols:
            master_df[col] = master_df[col].apply(lambda v: clean_text(v) if pd.notna(v) else v)

    logging.info("Arredondando e tratanto zeros")
    for col in numeric_cols:
        if col in master_df.columns:
            try:
                master_df[col] = pd.to_numeric(master_df[col], errors='coerce')
                master_df[col] = master_df[col].round(5)
                master_df.loc[master_df[col] == 0, col] = pd.NA
            except Exception:
                logging.debug("Falha ao processar coluna numérica: %s", col)

    master_df = master_df.loc[:, ~master_df.columns.str.lower().str.startswith('unnamed')]

    master_df.to_csv(output_csv, sep=';', encoding='utf-8-sig', index=False, decimal=',', na_rep='NULL')
    return os.path.abspath(output_csv)


def main():
    parser = argparse.ArgumentParser(description='Consolida planilhas de produtos de fornecedores')
    parser.add_argument('--input', '-i', default=EXCEL_FILE_PATH, help='Caminho do arquivo Excel de entrada')
    parser.add_argument('--output', '-o', default=OUTPUT_CSV_FILE, help='Arquivo CSV de saída')
    parser.add_argument('--markup', '-m', type=float, default=1.55, help='Multiplicador para preço de venda')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
    logging.info('Iniciando o processo de consolidação')

    try:
        result_path = process_excel(args.input, args.output, markup=args.markup)
        if result_path:
            logging.info('Processo concluído. Arquivo salvo em: %s', result_path)
        else:
            logging.info('Nenhum arquivo gerado (nenhuma aba válida encontrada).')
    except Exception as e:
        logging.exception('Erro fatal durante o processamento: %s', e)


if __name__ == '__main__':
    main()