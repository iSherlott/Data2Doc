
# Data2Doc

API HTTP (Gin) para gerar documentos a partir de dados tabulares.

- Rotas: `POST /generate/excel`, `POST /generate/pdf`, `POST /generate/word`
- Body: JSON ou XML (`Content-Type: application/json` ou `application/xml`)
- Resposta: download do arquivo (header `Content-Disposition: attachment; filename="..."`)

## Executar localmente

Pré-requisitos:

- Go instalado
- Um `APP_MODE` definido (o app não inicia sem isso)
- `AUTH_JWKS_URL` definido (as rotas são protegidas por Bearer token)

Exemplo de `.env`:

```bash
APP_MODE=dev
AUTH_JWKS_URL=https://example.com/.well-known/jwks.json
# opcional
AUTH_EXPECTED_CLIENT_ID=my-client-id
```

Rodar:

```bash
go run ./cmd
```

Swagger:

- `GET /swagger` (redireciona para `/swagger/index.html`)

## Autenticação e headers

Todas as rotas de geração exigem:

- `Authorization: Bearer <token>`

Headers úteis:

- `Content-Type`: `application/json` ou `application/xml`
- `Accept`: opcional (a rota sempre retorna o tipo do arquivo)
- `X-Request-Id`: opcional; define o nome base do arquivo retornado.
	- Ex.: `X-Request-Id: report-2026-03` → `report-2026-03.pdf` / `.xlsx` / `.docx`
	- Se omitido, o nome padrão é `document.*`

Sem query params (não existe `?id=...`).

## Formato do payload: `data`

O campo `data` é flexível e pode representar 1 dataset (default) ou vários datasets (sources).

### 1) Dataset único (default)

Aceita JSON como:

- Array de objetos (mais comum):

```json
"data": [
	{"name": "Ana", "salary": 1234.56},
	{"name": "Bruno", "salary": 2000}
]
```

- Objeto único (vira 1 linha):

```json
"data": {"name": "Ana", "salary": 1234.56}
```

### 2) Múltiplos datasets (sources)

Quando você precisa referenciar datasets por nome (ex.: `layout.sheets[].dataSource` no Excel ou `layout.blocks[].dataSource` em PDF/Word), use um objeto cujos valores sejam arrays/objetos:

```json
"data": {
	"employees": [{"name": "Ana"}, {"name": "Pedro"}],
	"departments": [{"dept": "IT"}]
}
```

Regras práticas:

- Objetos de escalares (ex.: `{ "name": "Ana", "age": 30 }`) são tratados como dataset default (1 linha).
- Datasets vazios são ignorados (ex.: `{ "a": {}, "b": {} }` não conta como “multi-dataset válido”).

## Como funciona (arquitetura resumida)

Fluxo (alto nível):

1. Router: `internal/routers/*` registra as rotas
2. Auth: middleware `AuthIdentityMiddleware()` valida JWT via JWKS
3. Handler: `internal/handlers/generate.go`
	 - Bind JSON/XML
	 - `Validate()` do request específico
	 - Converte para `models.DocumentRequest`
	 - Chama `service.DocumentService.Generate()`
4. Renderer: `internal/service/document.go` gera o arquivo e retorna bytes + content-type

Modelos principais:

- Requests por rota: `ExcelGenerateRequest`, `PDFGenerateRequest`, `WordGenerateRequest`
- Dados: `DataPayload` (default + sources) e `DynamicData`
- Builder mode (PDF/Word): `layout.blocks[]` com tipos `Text|SectionTitle|Table|Chart|Image|Spacer|PageBreak`

## Rotas

### 1) POST /generate/excel

Introdução:

- Gera `.xlsx` (excelize)
- Suporta dataset único ou múltiplas abas via `layout.sheets`

Aceita (`ExcelGenerateRequest`):

- `layout` (opcional)
	- `freezeHeader` (bool): congela a primeira linha
	- `freezeColumns` (int): congela as N primeiras colunas
	- `autoSizeColumns` (bool)
	- `hideEmptyRows` (bool) e `maxVisibleRows` (int): controla visibilidade de linhas/colunas excedentes
	- `groupBy` (string): agrupa e insere subtotais quando o valor do campo muda (o dataset é ordenado por esse campo). Em workbooks com múltiplas abas, o agrupamento é aplicado apenas nas abas cujo dataset contém esse campo.
	- `showTotalRow` (bool): força a linha final `TOTAL` mesmo sem agregações
	- `sheets[]` (opcional): cria múltiplas abas
		- `name` (string): nome da aba
		- `dataSource` (string): chave dentro de `data` (vazio usa `default`)
	- `columns[]` (opcional): define colunas e formatação
		- `field` (string): chave do objeto da linha
		- `title` (string)
		- `formula` (string): expressão por linha baseada em fields (ex.: `price * qty`) — o serviço resolve automaticamente para referências de célula
		- `sheetFormula` (string): fórmula Excel “crua” (pode referenciar outras abas)
		- `aggregate` (string): agregação automática (hoje: `sum`) — cria linha(s) de subtotal/total com `SUM(...)`
		- `percentageOf` (string): calcula percentual da coluna referenciada sobre o total (ex.: `salary / totalSalary`)
		- `format`: `currency|number|percentage|date|dateTime`
		- `cellType`: `Text|Number|Currency|Date|Select`
			- Se `Select`, `options[]` é obrigatório
		- `hidden` (bool), `locked` (bool)
- `data` (obrigatório)
	- Dataset único (default) quando `layout.sheets` não é usado
	- Multi-dataset quando `layout.sheets` referencia `dataSource`

Significado do payload:

- Sem `layout.sheets`: o Excel usa o dataset default como “tabela principal”.
- Com `layout.sheets`: cada aba lê seu dataset por `dataSource`.
- `layout.columns` é global: se as colunas configuradas não existirem no dataset de uma aba, as colunas dessa aba são inferidas automaticamente a partir das chaves do dataset.

Cross-sheet formulas:

- Dentro de `formula`/`sheetFormula`, você pode usar o token `sheet:<NomeDaAba>.<field>`.
	- Ex.: `SUM(sheet:Employees.salary)` → `SUM(Employees!B2:B10)` (o range é resolvido automaticamente com base nas linhas geradas).

### 2) POST /generate/pdf

Introdução:

- Gera `.pdf` (gofpdf)
- Tem dois modos:
	- Legacy (sem `layout.blocks`): renderiza uma tabela com o dataset default
	- Builder mode (`layout.blocks`): documento declarativo com conteúdo misto

Aceita (`PDFGenerateRequest`):

- `layout` (opcional)
	- `pageMargin`, `pageOrientation`, `defaultFont`
	- `pageBreak.enabled` e `pageBreak.rowsPerPage`
		- `rowsPerPage > 0`: quebra previsível por número de registros
		- `rowsPerPage == 0`: pagina por altura disponível
	- `spacing.paragraphSpacing` e `spacing.tableSpacing`
	- `footer.show`/`footer.alignment`/`footer.pageNumber.*`
	- `headerImage.*` (somente no PDF)
	- `columns[]` (somente legacy): define colunas da tabela
	- `blocks[]` (builder mode)
		- `type`: `Text|SectionTitle|Table|Chart|Image|Spacer|PageBreak`
		- `content`: usado por `Text|SectionTitle`
		- `height`: usado por `Spacer` e pode ser usado por `Chart`
		- `dataSource`: usado por `Table|Chart`
		- `columns`: usado por `Table` (lista de colunas)
		- `chartType`, `categoryField`, `valueField`, `title`: usados por `Chart`
		- `data`: usado por `Image` (base64 ou data URI)
- `data`
	- Legacy: obrigatório (dataset default)
	- Builder mode: obrigatório apenas se existir `Table`/`Chart` (ou seja, bloco que usa dataset)

Significado do payload:

- `PageBreak` (block): força quebra de página manual.
- `Chart`: precisa que `valueField` seja numérico; caso contrário falha.

### 3) POST /generate/word

Introdução:

- Gera `.docx` (ZIP + XML)
- Mesma ideia de legacy vs builder mode, mas com opções específicas de Word

Aceita (`WordGenerateRequest`):

- `layout` (opcional)
	- `layout.word.*`
		- `ignorePageMargins` (bool): usa largura total da página (ignora margens)
		- `centerContent` (*bool): centraliza por padrão (true)
		- `pageOrientation`: `Portrait|Landscape` (orientação do documento)
	- `pageMargin`, `pageBreak`, `footer`, `defaultFont`
	- `columns[]` (legacy): define colunas da tabela principal
	- `blocks[]` (builder mode): mesmos tipos do PDF (`Text|SectionTitle|Table|Chart|Image|Spacer|PageBreak`)
- `data`
	- Legacy: obrigatório (dataset default)
	- Builder mode: obrigatório apenas se existir `Table`/`Chart`

Nota:

- O renderer de Word suporta imagens em blocos (`Image`) e charts (são renderizados como PNG embutido).
- `headerImage` é exposto no request do PDF; não faz parte do request público do Word.

## Exemplos (3)

Substitua `$TOKEN` por um JWT válido (assinado pelo JWKS configurado).

### Exemplo 1 — Excel (múltiplas abas)

```bash
curl -X POST "http://localhost:8080/generate/excel" \
	-H "Authorization: Bearer $TOKEN" \
	-H "Content-Type: application/json" \
	-H "X-Request-Id: employees-and-depts" \
	-o employees-and-depts.xlsx \
	--data '{
		"layout": {
			"sheets": [
				{"name": "Employees", "dataSource": "employees"},
				{"name": "Departments", "dataSource": "departments"}
			],
			"columns": [
				{"field": "name", "title": "Name"},
				{"field": "salary", "title": "Salary", "format": "currency"}
			]
		},
		"data": {
			"employees": [{"name": "Ana", "salary": 1234.56}],
			"departments": [{"dept": "IT"}]
		}
	}'
```

### Exemplo 2 — PDF (builder mode: texto + tabela + chart + page break)

```bash
curl -X POST "http://localhost:8080/generate/pdf" \
	-H "Authorization: Bearer $TOKEN" \
	-H "Content-Type: application/json" \
	-H "X-Request-Id: report" \
	-o report.pdf \
	--data '{
		"layout": {
			"pageBreak": {"enabled": true},
			"blocks": [
				{"type": "SectionTitle", "content": "Relatório Financeiro"},
				{"type": "Text", "content": "Resumo"},
				{
					"type": "Table",
					"dataSource": "employees",
					"columns": [
						{"field": "name", "title": "Nome"},
						{"field": "salary", "title": "Salário"}
					]
				},
				{"type": "PageBreak"},
				{
					"type": "Chart",
					"chartType": "Column",
					"dataSource": "sales",
					"categoryField": "department",
					"valueField": "total",
					"title": "Vendas",
					"height": 80
				}
			]
		},
		"data": {
			"employees": [{"name": "Ana", "salary": 5000}, {"name": "Pedro", "salary": 6200}],
			"sales": [{"department": "Financeiro", "total": 120000}, {"department": "TI", "total": 180000}]
		}
	}'
```

### Exemplo 3 — Word (legacy: tabela com paginação por rowsPerPage)

```bash
curl -X POST "http://localhost:8080/generate/word" \
	-H "Authorization: Bearer $TOKEN" \
	-H "Content-Type: application/json" \
	-H "X-Request-Id: people" \
	-o people.docx \
	--data '{
		"layout": {
			"word": {"pageOrientation": "Portrait", "centerContent": true},
			"pageBreak": {"enabled": true, "rowsPerPage": 25},
			"columns": [
				{"field": "name", "title": "Name"},
				{"field": "city", "title": "City"}
			]
		},
		"data": [
			{"name": "Ana", "city": "São Paulo"},
			{"name": "Bruno", "city": "Porto"}
		]
	}'
```

