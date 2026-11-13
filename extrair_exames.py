"""Extração local de tabelas de exames com colunas alinhadas."""
import argparse
import csv
import json
import re
import unicodedata
from pathlib import Path

import pdfplumber


FIELDS = ['funcionário', 'cpf', 'setor', 'função', 'data exame',
          'procedimento', 'prestador', 'tipo do exame', 'valor']
LABEL = re.compile(r'\b(FUNCION[ÁA]RIO|CPF|SETOR|FUN[ÇC][ÃA]O)\s*:\s*')
DATE = re.compile(r'^\d{2}/\d{2}/\d{4}\b')
ALIASES = [['DATA EXAME', 'DATA DO EXAME', 'DATA'], ['PROCEDIMENTO'],
           ['PRESTADOR'], ['TIPO DO EXAME', 'TIPO DE EXAME', 'TIPO EXAME', 'TIPO'], ['VALOR']]


def norm(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text.upper())
                   if unicodedata.category(c) != 'Mn')


def lines(words, tolerance):
    grouped = []
    for word in sorted(words, key=lambda w: (w['top'], w['x0'])):
        if not grouped or abs(word['top'] - grouped[-1][0]['top']) > tolerance:
            grouped.append([word])
        else:
            grouped[-1].append(word)
    return [sorted(group, key=lambda w: w['x0']) for group in grouped]


def detect_columns(words):
    tokens = [norm(w['text']) for w in words]
    starts = []
    for aliases in ALIASES:
        match = None
        for alias in aliases:
            parts = alias.split()
            for i in range(len(tokens) - len(parts) + 1):
                if tokens[i:i + len(parts)] == parts:
                    match = words[i]['x0']
                    break
            if match is not None:
                break
        if match is None:
            return None
        starts.append(match)
    return starts if starts == sorted(set(starts)) else None


def cells(words, starts):
    values = [[] for _ in starts]
    for word in words:
        index = max([i for i, start in enumerate(starts) if word['x0'] >= start - 3] or [0])
        values[index].append(word['text'])
    return [' '.join(value) for value in values]


def extract(pdf_path, output_path, starts=None, tolerance=3):
    rows, warnings = [], []
    employee, pending = {}, None
    columns = starts
    in_table = False

    def flush():
        nonlocal pending
        if pending is not None:
            if not all(employee.get(k) for k in FIELDS[:4]):
                raise ValueError('Bloco com dados de funcionário incompletos; CSV não gerado.')
            if not pending[4]:
                raise ValueError('Exame sem valor identificado; ajuste os limites das colunas.')
            rows.append([employee[k] for k in FIELDS[:4]] + pending)
            pending = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=2, y_tolerance=tolerance)
            if not words:
                raise ValueError(f'Página {page_number} sem texto extraível. É necessário OCR local.')
            for line in lines(words, tolerance):
                text = ' '.join(w['text'] for w in line)
                labels = list(LABEL.finditer(text))
                if labels:
                    if any(norm(m.group(1)) == 'FUNCIONARIO' for m in labels):
                        flush()
                        employee = {}
                        columns = starts
                    elif in_table:
                        flush()
                    in_table = False
                    for i, match in enumerate(labels):
                        key = {'FUNCIONARIO': 'funcionário', 'CPF': 'cpf',
                               'SETOR': 'setor', 'FUNCAO': 'função'}[norm(match.group(1))]
                        end = labels[i + 1].start() if i + 1 < len(labels) else len(text)
                        employee[key] = text[match.end():end].strip()
                    continue
                header = detect_columns(line)
                if header:
                    columns = starts or header
                    in_table = True
                    continue
                if DATE.match(text):
                    if not columns:
                        raise ValueError(f'Cabeçalho não identificado na página {page_number}. Use --colunas.')
                    flush()
                    pending = cells(line, columns)
                    in_table = True
                elif in_table and pending:
                    # Continuations require an empty date AND value column. Totals and
                    # page footers end the table, never becoming another examination.
                    parts = cells(line, columns)
                    if re.search(r'\b(TOTAL|SUBTOTAL|PAGINA|PAG\.)\b', norm(text)):
                        flush()
                        in_table = False
                    elif not parts[0] and not parts[4]:
                        pending = [' '.join(filter(None, pair)) for pair in zip(pending, parts)]
                    else:
                        warnings.append(f'Página {page_number}: linha não interpretada em y={line[0]["top"]:.1f}.')
                        flush()
                        in_table = False
        flush()
    if not rows:
        raise ValueError('Nenhum exame identificado. Verifique cabeçalhos e datas (DD/MM/AAAA).')
    if warnings:
        raise ValueError('\n'.join(warnings) + '\nCSV não gerado para evitar extração parcial silenciosa.')
    with output_path.open('x', encoding='utf-8-sig', newline='') as target:
        writer = csv.writer(target, delimiter=';')
        writer.writerow(FIELDS)
        writer.writerows(rows)
    return len(rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('csv', type=Path)
    parser.add_argument('--colunas', type=Path, help='JSON com cinco posições iniciais em pontos PDF.')
    parser.add_argument('--tolerancia-linha', type=float, default=3)
    args = parser.parse_args()
    try:
        starts = json.loads(args.colunas.read_text()) if args.colunas else None
        if starts is not None and (not isinstance(starts, list) or len(starts) != 5
                                  or not all(type(x) in (int, float) for x in starts)
                                  or starts != sorted(set(starts))):
            raise ValueError('A configuração precisa conter cinco números em ordem crescente.')
        count = extract(args.pdf, args.csv, starts, args.tolerancia_linha)
    except (ValueError, OSError) as error:
        parser.exit(1, f'Erro: {error}\n')
    print(f'{count} exames exportados para {args.csv}')


if __name__ == '__main__':
    main()
