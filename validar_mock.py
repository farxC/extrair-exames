"""Executa extração real e compara todas as células com o resultado esperado."""
import csv
import tempfile
from pathlib import Path
from extrair_exames import extract

directory = Path(__file__).resolve().parent / 'mock'

def read(path):
    with path.open(encoding='utf-8-sig', newline='') as source:
        return list(csv.reader(source, delimiter=';'))

with tempfile.TemporaryDirectory() as temp:
    output = Path(temp) / 'extraido.csv'
    count = extract(directory / 'exames_mock.pdf', output)
    assert read(output) == read(directory / 'esperado.csv'), 'CSV diferente do esperado'
    destination = directory / 'extraido.csv'
    with destination.open('xb') as target:
        target.write(output.read_bytes())
    print(f'OK: {count} exames; todas as células conferidas. CSV: {destination}')
