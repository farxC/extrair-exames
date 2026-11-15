"""Insere tipos de exame de um CSV nas colunas correspondentes de uma planilha."""
import argparse
import csv
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


CSV_FIELDS = ('procedimento', 'funcionario', 'tipo do exame')
CPF_HEADERS = ('CPF', 'CPF COLABORADOR')
EMPLOYEE_HEADERS = ('FUNCIONARIO', 'NOME COLABORADOR')
HEADER_SEARCH_ROWS = 50
PROCEDURE_TO_COLUMN_RAW = {
    'AUDIOMETRICO': 'EXAME CLINICO AUDIOMETRICO',
    'AVALIACAO ACID DE TRABALHO': 'EXAME CLINICO AVALIACAO ACID DE TRABALHO',
    'AVALIACAO DE AFASTAMENTO': 'EXAME CLINICO AVALIACAO DE AFASTAMENTO',
    'ECG-ELETROCARDIOGRAMA': 'EXAME CLINICO ECG - ELETROCARDIOGRAMA',
    'EEG - ELETROENCEFALOGRAMA': 'EXAME CLINICO EEG - ELETROENCEFALOGRAMA',
    'EGG - ELETROENCEFALOGRAMA': 'EXAME CLINICO EEG - ELETROENCEFALOGRAMA',
    'ESPIROMETRIA': 'EXAME CLINICO ESPIROMETRIA',
    'OCUPACIONAL': 'EXAME CLINICO OCUPACIONAL',
    'OPTOMETRIA - TESTE VISUAL': 'EXAME CLINICO OPTOMETRIA - TESTE VISUAL',
    'ACIDO HIPURICO': 'EXAME LAB. ACIDO HIPURICO',
    'ACIDO METIL HIPURICO': 'EXAME LAB. ACIDO METIL HIPURICO',
    'ALAU': 'EXAME LAB ALAU',
    'BETA': 'EXAME LAB BETA',
    'CADMIO': 'EXAME LAB. CADMIO',
    'CARBOXIHEMOGLOBINA': 'EXAME LAB. CARBOXIHEMOGLOBINA',
    'CHUMBO': 'EXAME LAB. CHUMBO',
    'COLESTEROL TOTAL E FRACOES': 'EXAME LAB. COLESTEROL TOTAL E FRACOES',
    'COLINESTERASE': 'EXAME LAB. COLINESTERASE',
    'CROMO': 'EXAME LAB. CROMO',
    'CROPOCULTURA': 'EXAME LAB CROPOCULTURA',
    'GAMA GT': 'EXAME LAB. GAMA GT',
    'GLICEMIA DE JEJUM': 'EXAME LAB. GLICEMIA DE JEJUM',
    'HEMOGRAMA COMPLETO': 'EXAME LAB. HEMOGRAMA COMPLETO',
    'MANGANES': 'EXAME LAB. MANGANES',
    'PARASITOLOGICO DE FEZES': 'EXAME LAB. PARASITOLOGICO DE FEZES',
    'SOROLOGIA CHAGAS': 'EXAME LAB SOROLOGIA CHAGAS',
    'TGO': 'EXAME LAB. TGO',
    'TGP': 'EXAME LAB. TGP',
    'URINA TIPO I': 'EXAME LAB. URINA TIPO I',
    'RX COLUNA LOMBO SACRA': 'EXAME RADIOLOG. - RX COLUNA LOMBO SACRA',
    'RX TORAX PA PADRAO OIT': 'EXAME RADIOLOG. - RX TORAX PA PADRAO OIT',
    'FRETE MUNICIPAL': 'FRETE MUNICIPAL',
    'UREIA': 'EXAME LAB. UERIA',
    'CREATININA': 'EXAME LAB. CREATININA',
    'PSA TOTAL': 'EXAME LAB. PSA',
}


def canonical(value):
    """Compara textos sem acentos, pontuação, caixa ou espaços extras."""
    text = '' if value is None else str(value)
    text = ''.join(char for char in unicodedata.normalize('NFD', text.upper())
                   if unicodedata.category(char) != 'Mn')
    return ' '.join(''.join(char if char.isalnum() else ' ' for char in text).split())


PROCEDURE_TO_COLUMN = {
    canonical(procedure): target_column
    for procedure, target_column in PROCEDURE_TO_COLUMN_RAW.items()
}


def text_value(value):
    return '' if value is None else ' '.join(str(value).split())


def cpf_digits(value):
    digits = ''.join(char for char in text_value(value) if char.isdigit())
    return digits.zfill(11) if digits else ''


def indexed_headers(row):
    headers = {}
    for index, value in enumerate(row):
        key = canonical(value)
        if key:
            headers.setdefault(key, index)
    return headers


def read_csv_rows(source_path):
    with source_path.open(encoding='utf-8-sig', newline='') as source:
        reader = csv.DictReader(source, delimiter=';')
        if not reader.fieldnames:
            raise ValueError('O CSV está vazio ou não possui cabeçalho.')
        fields = {canonical(name): name for name in reader.fieldnames if name}
        missing = [field for field in CSV_FIELDS if canonical(field) not in fields]
        if missing:
            raise ValueError(f'Coluna obrigatória ausente no CSV: {", ".join(missing)}')
        cpf_field = fields.get(canonical('cpf'))

        mapped_rows = []
        unknown_procedures = set()
        for line_number, row in enumerate(reader, 2):
            procedure = text_value(row[fields[canonical('procedimento')]])
            employee = text_value(row[fields[canonical('funcionario')]])
            cpf = cpf_digits(row[cpf_field]) if cpf_field else ''
            exam_type = text_value(row[fields[canonical('tipo do exame')]])
            if not procedure and not employee and not cpf and not exam_type:
                continue
            if not procedure or not employee or not exam_type:
                raise ValueError(f'Linha {line_number} do CSV possui dado obrigatório vazio.')
            target_column = PROCEDURE_TO_COLUMN.get(canonical(procedure))
            if not target_column:
                unknown_procedures.add(procedure)
                continue
            mapped_rows.append((cpf, employee, target_column, exam_type))
    if unknown_procedures:
        raise ValueError('Procedimento sem nomenclatura cadastrada: '
                         + ', '.join(sorted(unknown_procedures)))
    if not mapped_rows:
        raise ValueError('Nenhuma linha de exame foi encontrada no CSV.')
    return mapped_rows


def find_target_header(sheet):
    for row_number, row in enumerate(
            sheet.iter_rows(max_row=HEADER_SEARCH_ROWS, values_only=True), 1):
        headers = indexed_headers(row)
        if (any(canonical(name) in headers for name in CPF_HEADERS)
                or any(canonical(name) in headers for name in EMPLOYEE_HEADERS)):
            return row_number, headers
    raise ValueError('Coluna CPF ou funcionário não encontrada nas primeiras '
                     f'{HEADER_SEARCH_ROWS} linhas da planilha alvo.')


def column_for(headers, names):
    for name in names:
        column = headers.get(canonical(name))
        if column is not None:
            return column
    return None


def append_value(existing, new_value):
    existing_text = text_value(existing)
    if not existing_text:
        return new_value
    values = [part.strip() for part in existing_text.split(';') if part.strip()]
    if canonical(new_value) in {canonical(value) for value in values}:
        return existing_text
    return f'{existing_text}; {new_value}'


def fill_workbook(csv_path, target_path, output_path):
    """Preenche uma cópia de ``target_path`` com tipos de exame do CSV."""
    csv_path = Path(csv_path)
    target_path = Path(target_path)
    output_path = Path(output_path)
    if target_path.suffix.lower() != '.xlsx':
        raise ValueError('A planilha alvo deve ser um arquivo .xlsx.')
    if output_path.suffix.lower() != '.xlsx':
        raise ValueError('A planilha de saída deve usar a extensão .xlsx.')
    if target_path.resolve() == output_path.resolve():
        raise ValueError('A saída deve ser diferente da planilha alvo para preservar o original.')
    if output_path.exists():
        raise FileExistsError(f'A planilha de saída já existe: {output_path}')

    source_rows = read_csv_rows(csv_path)
    workbook = load_workbook(target_path)
    try:
        sheet = workbook.active
        header_row, headers = find_target_header(sheet)
        cpf_column = column_for(headers, CPF_HEADERS)
        employee_column = column_for(headers, EMPLOYEE_HEADERS)
        if employee_column is None:
            raise ValueError('Coluna funcionário não encontrada na planilha alvo.')
        if any(cpf for cpf, _, _, _ in source_rows) and cpf_column is None:
            raise ValueError('Coluna CPF não encontrada na planilha alvo.')

        required_columns = {canonical(target_column) for _, _, target_column, _ in source_rows}
        missing_columns = sorted(
            target_column for target_column in required_columns if target_column not in headers
        )
        if missing_columns:
            raise ValueError('Coluna de exame não encontrada na planilha alvo: '
                             + ', '.join(missing_columns))

        target_rows_by_cpf = {}
        target_rows_by_employee = {}
        for row in sheet.iter_rows(min_row=header_row + 1):
            if cpf_column is not None:
                cpf = cpf_digits(row[cpf_column].value)
                if cpf:
                    if cpf in target_rows_by_cpf:
                        raise ValueError(f'CPF duplicado na planilha alvo: {cpf}')
                    target_rows_by_cpf[cpf] = row
            employee = canonical(row[employee_column].value)
            if employee:
                if employee in target_rows_by_employee:
                    raise ValueError(f'Funcionário duplicado na planilha alvo: {employee}')
                target_rows_by_employee[employee] = row

        missing_people = []
        selected_rows = []
        for cpf, employee, target_column, exam_type in source_rows:
            row_by_cpf = target_rows_by_cpf.get(cpf) if cpf else None
            row_by_employee = target_rows_by_employee.get(canonical(employee))
            if cpf and row_by_cpf is None:
                missing_people.append(f'CPF {cpf}')
                continue
            if row_by_employee is None:
                missing_people.append(f'funcionário {employee}')
                continue
            if row_by_cpf is not None and row_by_cpf != row_by_employee:
                raise ValueError('CPF e funcionário apontam para pessoas diferentes: '
                                 f'{cpf} / {employee}')
            selected_rows.append((row_by_cpf or row_by_employee, target_column, exam_type))
        if missing_people:
            raise ValueError('Pessoa não encontrada na planilha alvo: '
                             + ', '.join(sorted(set(missing_people))))

        changed = 0
        for row, target_column, exam_type in selected_rows:
            cell = row[headers[canonical(target_column)]]
            merged_value = append_value(cell.value, exam_type)
            if merged_value != text_value(cell.value):
                cell.value = merged_value
                changed += 1
        workbook.save(output_path)
    finally:
        workbook.close()
    return changed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('csv', type=Path,
                        help='CSV com procedimento, funcionário, CPF opcional e tipo do exame.')
    parser.add_argument('planilha_alvo', type=Path, help='Planilha .xlsx que receberá os dados.')
    parser.add_argument('saida', type=Path, help='Nova planilha .xlsx preenchida.')
    args = parser.parse_args()
    try:
        changed = fill_workbook(args.csv, args.planilha_alvo, args.saida)
    except (InvalidFileException, OSError, ValueError) as error:
        parser.exit(1, f'Erro: {error}\n')
    print(f'{changed} células atualizadas em {args.saida}')


if __name__ == '__main__':
    main()
