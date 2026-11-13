# Extrair exames de PDF para CSV

Processamento inteiramente local com Python. O documento deve conter texto
selecionável; OCR não está incluído.

## Instalação e execução

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python extrair_exames.py documento.pdf exames.csv
```

O CSV usa `;`, UTF-8 com BOM e uma linha por exame. Colunas:
funcionário, cpf, setor, função, data exame, procedimento, prestador,
tipo do exame e valor. Valores monetários e CPFs são preservados como texto.
Ao importar no Excel, escolha o tipo Texto para CPF, preservando zeros iniciais.
Um arquivo de saída existente nunca é sobrescrito.

## Layout esperado

- Rótulos `FUNCIONÁRIO:`, `CPF:`, `SETOR:` e `FUNÇÃO:` antes da tabela,
  juntos ou em linhas separadas. Cada valor deve estar na mesma linha do rótulo.
- Um novo `FUNCIONÁRIO:` inicia um novo bloco.
- Cabeçalho em uma linha: Data exame, Procedimento, Prestador, Tipo do exame, Valor.
- Cada exame começa com uma data `DD/MM/AAAA` na primeira coluna.
- Linhas de continuação podem complementar procedimento, prestador e tipo,
  desde que as colunas de data e valor estejam vazias.
- Tabelas podem continuar na página seguinte, com cabeçalho repetido.

## Ajustar as colunas

A detecção automática usa o começo dos títulos como limites de coluna.
Se os títulos estiverem centralizados, quebrados em linhas ou forem diferentes,
crie `colunas.json` com cinco limites horizontais em pontos PDF (72 pontos =
1 polegada), medidos da esquerda para a direita. Exemplo apenas ilustrativo:

```json
[30, 110, 270, 390, 510]
```

```sh
python extrair_exames.py documento.pdf exames.csv --colunas colunas.json
```

Cada limite deve ficar antes de todos os textos da respectiva coluna e depois
dos textos da coluna anterior. A configuração vale para todas as páginas.
`--tolerancia-linha 3` controla a tolerância vertical em pontos para agrupar palavras.

## Validação necessária

Se houver duas linhas de rodapé entre a tabela e a próxima página, use:

```sh
python extrair_exames.py documento.pdf exames.csv --linhas-rodape 2
```

Isso ignora as duas últimas linhas **de texto** de cada página, exceto a última,
antes de interpretar a tabela. O funcionário e as posições das colunas são
mantidos para a página seguinte. Linhas vazias e bordas não entram na contagem.
Use essa opção apenas quando esse rodapé existir em todas as páginas intermediárias.
O script interrompe se detectar data de exame ou rótulos pessoais nas linhas que
seriam descartadas. O padrão continua sendo não descartar linhas.

Este extrator é uma base para o layout descrito, não uma garantia de extração
universal. Confira localmente a quantidade de exames, os vínculos com cada
funcionário e os valores. Cabeçalhos/rodapés incomuns, campos quebrados em linhas,
datas em outro formato e mudanças de layout precisam de adaptação.
O script interrompe em alguns casos detectáveis (dados pessoais incompletos,
exame sem valor, página sem texto ou linha ambígua), mas isso não detecta toda
possível omissão ou associação incorreta. Nenhum conteúdo do PDF é enviado à rede.

A instalação das dependências requer acesso ao PyPI. Neste ambiente, esse acesso
não estava disponível; a leitura de PDFs reais não foi validada.

## Mock reproduzível

`mock/exames_mock.pdf` contém dois funcionários fictícios e quatro exames em
duas páginas, sem bordas. Inclui acentos, procedimento em duas linhas, prestador
vazio, totais e continuação de tabela na segunda página.
`mock/esperado.csv` é a referência escrita pelo gerador, **não uma extração**.

```sh
python -m pip install -r requirements.txt reportlab
python gerar_mock.py
python validar_mock.py
```

O validador executa o extrator real e compara todas as células com a referência.
Somente após essa comparação passar ele grava `mock/extraido.csv`. Esse arquivo
não pode existir antes da execução. A tentativa neste ambiente foi bloqueada
por `ModuleNotFoundError: No module named 'pdfplumber'`.
