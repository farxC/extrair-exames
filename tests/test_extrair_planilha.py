import csv
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from extrair_planilha import extract_workbook


FIELDS = ['funcionário', 'cpf', 'setor', 'função', 'data exame',
          'procedimento', 'prestador', 'tipo do exame', 'valor']


def read_csv(path):
    with path.open(encoding='utf-8-sig', newline='') as source:
        return list(csv.reader(source, delimiter=';'))


def create_workbook(path, headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


class ExtractWorkbookTests(unittest.TestCase):
    def test_extract_workbook_finds_header_after_introductory_rows(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'exames.xlsx'
            output = workspace / 'exames.csv'
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(['RELATÓRIO DE EXAMES OCUPACIONAIS'])
            sheet.append([
                'FORNECEDOR', 'CPF COLABORADOR', 'NOME COLABORADOR',
                'CENTRO DE CUSTO', 'TIPO DO EXAME', 'UREIA',
            ])
            sheet.append(['Clínica Alfa', '000.000.000-00', 'Ana Exemplo',
                          'Administrativo', 'Periódico', '12,50'])
            workbook.save(source)

            count = extract_workbook(source, output)

            self.assertEqual(count, 1)
            self.assertEqual(read_csv(output)[1][5], 'UREIA')

    def test_extract_workbook_creates_one_csv_row_for_each_filled_exam(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'exames.xlsx'
            output = workspace / 'exames.csv'
            create_workbook(source, [
                ' fornecedor ', 'LOCAL PRESTAÇÃO\nSERVIÇO', 'CPF colaborador',
                'nome colaborador', 'centro de custo', 'tipo do exame',
                'UREIA', 'CREATININA', 'EM BRANCO',
            ], [
                ['Clínica Alfa', 'Unidade Centro', '000.000.000-00', 'Ana Exemplo',
                 'Administrativo', 'Periódico', 12.5, '18,20', '   '],
                ['Clínica Beta', 'Unidade Norte', '111.111.111-11', 'Bruno Exemplo',
                 'Operações', 'Admissional', None, None, None],
            ])

            count = extract_workbook(source, output)

            self.assertEqual(count, 2)
            self.assertEqual(read_csv(output), [
                FIELDS,
                ['Ana Exemplo', '000.000.000-00', 'Administrativo', '', '',
                 'UREIA', 'Clínica Alfa', 'Periódico', '12.5'],
                ['Ana Exemplo', '000.000.000-00', 'Administrativo', '', '',
                 'CREATININA', 'Clínica Alfa', 'Periódico', '18,20'],
            ])

    def test_extract_workbook_rejects_missing_required_header(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'sem_nome.xlsx'
            output = workspace / 'exames.csv'
            create_workbook(source, [
                'FORNECEDOR', 'CPF COLABORADOR', 'CENTRO DE CUSTO',
                'TIPO DO EXAME', 'UREIA',
            ], [['Clínica Alfa', '000.000.000-00', 'Administrativo', 'Periódico', 'X']])

            with self.assertRaisesRegex(ValueError, 'NOME COLABORADOR'):
                extract_workbook(source, output)

            self.assertFalse(output.exists())

    def test_extract_workbook_never_overwrites_existing_csv(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'exames.xlsx'
            output = workspace / 'exames.csv'
            create_workbook(source, [
                'FORNECEDOR', 'CPF COLABORADOR', 'NOME COLABORADOR',
                'CENTRO DE CUSTO', 'TIPO DO EXAME', 'UREIA',
            ], [['Clínica Alfa', '000.000.000-00', 'Ana Exemplo',
                'Administrativo', 'Periódico', 'X']])
            output.write_text('conteúdo anterior', encoding='utf-8')

            with self.assertRaises(FileExistsError):
                extract_workbook(source, output)

            self.assertEqual(output.read_text(encoding='utf-8'), 'conteúdo anterior')


if __name__ == '__main__':
    unittest.main()
