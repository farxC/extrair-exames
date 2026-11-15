import csv
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from preencher_planilha import fill_workbook


def create_csv(path, rows):
    with path.open('w', encoding='utf-8-sig', newline='') as target:
        writer = csv.DictWriter(
            target,
            fieldnames=['procedimento', 'cpf', 'tipo do exame', 'campo ignorado'],
            delimiter=';',
        )
        writer.writeheader()
        writer.writerows(rows)


def create_target_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['PLANILHA DE EXAMES'])
    sheet.append([
        'CPF COLABORADOR', 'EXAME LAB. UERIA',
        'EXAME CLINICO AUDIOMETRICO', 'OUTRA COLUNA',
    ])
    sheet.append(['000.000.000-00', 'ADMISSIONAL', '', 'não alterar'])
    sheet.append(['111.111.111-11', '', '', 'não alterar'])
    workbook.save(path)


class FillWorkbookTests(unittest.TestCase):
    def test_fill_workbook_maps_procedures_matches_unformatted_cpf_and_concatenates(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'base.csv'
            target = workspace / 'modelo.xlsx'
            output = workspace / 'resultado.xlsx'
            create_csv(source, [
                {'procedimento': 'UREIA', 'cpf': '00000000000',
                 'tipo do exame': 'PERIÓDICO', 'campo ignorado': 'ignorar'},
                {'procedimento': 'AUDIOMETRICO', 'cpf': '00000000000',
                 'tipo do exame': 'PERIÓDICO', 'campo ignorado': 'ignorar'},
                {'procedimento': 'UREIA', 'cpf': '00000000000',
                 'tipo do exame': 'PERIÓDICO', 'campo ignorado': 'ignorar'},
            ])
            create_target_workbook(target)

            changed = fill_workbook(source, target, output)

            self.assertEqual(changed, 2)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook.active
            self.assertEqual(sheet.cell(3, 2).value, 'ADMISSIONAL; PERIÓDICO')
            self.assertEqual(sheet.cell(3, 3).value, 'PERIÓDICO')
            self.assertEqual(sheet.cell(3, 4).value, 'não alterar')
            workbook.close()
            self.assertFalse(output.samefile(target))

    def test_fill_workbook_rejects_csv_without_required_columns(self):
        with TemporaryDirectory() as directory:
            workspace = Path(directory)
            source = workspace / 'base.csv'
            target = workspace / 'modelo.xlsx'
            output = workspace / 'resultado.xlsx'
            with source.open('w', encoding='utf-8-sig', newline='') as target_file:
                writer = csv.writer(target_file, delimiter=';')
                writer.writerow(['procedimento', 'cpf'])
                writer.writerow(['UREIA', '00000000000'])
            create_target_workbook(target)

            with self.assertRaisesRegex(ValueError, 'tipo do exame'):
                fill_workbook(source, target, output)

            self.assertFalse(output.exists())


if __name__ == '__main__':
    unittest.main()
