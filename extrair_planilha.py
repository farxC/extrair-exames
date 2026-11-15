"""Converte uma planilha de exames por colaborador para o CSV padrão."""
import argparse
import csv
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


FIELDS = ['funcionário', 'cpf', 'setor', 'função', 'data exame',
          'procedimento', 'prestador', 'tipo do exame', 'valor']
REQUIRED_HEADERS = (
    'FORNECEDOR',
    'CPF COLABORADOR',
    'NOME COLABORADOR',
    'CENTRO DE CUSTO',
    'TIPO DO EXAME',
)
HEADER_SEARCH_ROWS = 50


def normalized_header(value):
    """Normaliza caixa, acentos e espaços para comparar cabeçalhos."""
    text = '' if value is None else str(value)
    text = ''.join(char for char in unicodedata.normalize('NFD', text.upper())
                   if unicodedata.category(char) != 'Mn')
    return ' '.join(text.split())


def text_value(value):
    """Converte uma célula do Excel em texto de CSV."""
    return '' if value is None else ' '.join(str(value).split())


def header_positions(header):
    positions = {}
    for index, value in enumerate(header):
        key = normalized_header(value)
        if key not in REQUIRED_HEADERS:
            continue
        if key in positions:
            raise ValueError(f'Cabeçalho duplicado: {key}')
        positions[key] = index
    return positions


def find_header(rows):
    """Encontra o cabeçalho entre as primeiras linhas da primeira aba."""
    for row_number, header in enumerate(rows, 1):
        positions = header_positions(header)
        if all(name in positions for name in REQUIRED_HEADERS):
            return header, positions
        if row_number == HEADER_SEARCH_ROWS:
            break
    raise ValueError(
        'Cabeçalho obrigatório não encontrado nas primeiras '
        f'{HEADER_SEARCH_ROWS} linhas: {", ".join(REQUIRED_HEADERS)}'
    )


def extract_workbook(input_path, output_path):
    """Exporta os exames preenchidos da primeira aba de ``input_path``."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    if input_path.suffix.lower() != '.xlsx':
        raise ValueError('A planilha de entrada deve ser um arquivo .xlsx.')

    workbook = load_workbook(input_path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        header, positions = find_header(rows)
        type_position = positions['TIPO DO EXAME']
        exams = [
            (index, text_value(name))
            for index, name in enumerate(header[type_position + 1:], type_position + 1)
            if text_value(name)
        ]
        if not exams:
            raise ValueError('Nenhuma coluna de exame foi encontrada após TIPO DO EXAME.')

        exported = []
        for source_row in rows:
            person = {
                field: text_value(source_row[index]) if index < len(source_row) else ''
                for field, index in positions.items()
            }
            for index, procedure in exams:
                value = text_value(source_row[index]) if index < len(source_row) else ''
                if not value:
                    continue
                exported.append([
                    person['NOME COLABORADOR'],
                    person['CPF COLABORADOR'],
                    person['CENTRO DE CUSTO'],
                    '',
                    '',
                    procedure,
                    person['FORNECEDOR'],
                    person['TIPO DO EXAME'],
                    value,
                ])
    finally:
        workbook.close()

    if not exported:
        raise ValueError('Nenhum exame preenchido foi encontrado na planilha.')
    with output_path.open('x', encoding='utf-8-sig', newline='') as target:
        writer = csv.writer(target, delimiter=';')
        writer.writerow(FIELDS)
        writer.writerows(exported)
    return len(exported)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('planilha', type=Path, help='Arquivo .xlsx com uma linha por colaborador.')
    parser.add_argument('csv', type=Path, help='CSV de saída, que não pode existir.')
    args = parser.parse_args()
    try:
        count = extract_workbook(args.planilha, args.csv)
    except (InvalidFileException, OSError, ValueError) as error:
        parser.exit(1, f'Erro: {error}\n')
    print(f'{count} exames exportados para {args.csv}')


if __name__ == '__main__':
    main()
