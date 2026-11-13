"""Extrai todos os PDFs diretamente dentro de uma pasta, um CSV por documento."""
import argparse
import json
from pathlib import Path

from extrair_exames import extract


def process_folder(source, destination, starts=None, tolerance=3, footer_lines=0):
    if not source.is_dir():
        raise ValueError(f'Pasta de entrada não encontrada: {source}')
    files = sorted((p for p in source.iterdir() if p.is_file() and p.suffix.lower() == '.pdf'),
                   key=lambda p: p.name)
    if not files:
        raise ValueError('Nenhum arquivo PDF encontrado diretamente nessa pasta.')
    destination.mkdir(parents=True, exist_ok=True)
    successes = failures = 0
    for pdf in files:
        output = destination / f'{pdf.stem}.csv'
        try:
            # Never overwrite a prior export, including collisions between PDF names.
            if output.exists():
                raise FileExistsError(f'CSV já existe: {output.name}')
            count = extract(pdf, output, starts, tolerance, footer_lines)
        except Exception as error:
            # One malformed, encrypted or unreadable document must not stop the batch.
            failures += 1
            print(f'ERRO: {pdf.name}: {error}')
        else:
            successes += 1
            print(f'OK: {pdf.name} → {output.name} ({count} exames)')
    print(f'Resumo: {successes} concluídos; {failures} com erro; {len(files)} PDFs encontrados.')
    return successes, failures


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pasta', type=Path, help='Pasta que contém os PDFs (sem subpastas).')
    parser.add_argument('saida', type=Path, help='Pasta onde salvar os CSVs.')
    parser.add_argument('--colunas', type=Path)
    parser.add_argument('--tolerancia-linha', type=float, default=3)
    parser.add_argument('--linhas-rodape', type=int, default=0)
    args = parser.parse_args()
    try:
        starts = json.loads(args.colunas.read_text()) if args.colunas else None
        if starts is not None and (not isinstance(starts, list) or len(starts) != 5
                                  or not all(type(x) in (int, float) for x in starts)
                                  or starts != sorted(set(starts))):
            raise ValueError('A configuração precisa conter cinco números em ordem crescente.')
        if args.linhas_rodape < 0:
            raise ValueError('--linhas-rodape deve ser não negativo.')
        if args.tolerancia_linha <= 0:
            raise ValueError('--tolerancia-linha deve ser positiva.')
        _, failures = process_folder(args.pasta, args.saida, starts,
                                     args.tolerancia_linha, args.linhas_rodape)
    except (ValueError, OSError) as error:
        parser.exit(1, f'Erro: {error}\n')
    if failures:
        parser.exit(1)


if __name__ == '__main__':
    main()
