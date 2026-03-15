
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
			- `cellType`: `Text|Number|Currency|Date|Select|Formula|Lookup`
			- Se `Select`, `options[]` é obrigatório
				- Se `Formula`, o valor do campo na linha deve ser uma fórmula Excel (ex.: `"=SUM(A2:A10)"`)
				- Se `Lookup`, `lookup` é obrigatório (ver “Passo a passo” abaixo)
			- `validationRange` (string): validação por lista usando um range (ex.: `"Products!$A$2:$A$50"`)
			- `conditionalFormatting[]` (opcional): regras de formatação condicional por coluna
			- `backgroundColor`, `textColor`, `headerColor` (opcional): cores por coluna (hex como `#RRGGBB`)
		- `hidden` (bool), `locked` (bool)
		- `charts[]` (opcional): adiciona charts no Excel
- `data` (obrigatório)
	- Dataset único (default) quando `layout.sheets` não é usado
	- Multi-dataset quando `layout.sheets` referencia `dataSource`

Significado do payload:

- Sem `layout.sheets`: o Excel usa o dataset default como “tabela principal”.
- Com `layout.sheets`: cada aba lê seu dataset por `dataSource`.
- `layout.columns` é global: se as colunas configuradas não existirem no dataset de uma aba, as colunas dessa aba são inferidas automaticamente a partir das chaves do dataset.
	- `layout.columns` é global, mas o renderer aplica uma heurística de “colunas aplicáveis por aba”:
		- Se uma aba tiver poucas/nenhuma coluna aplicável (ex.: aba auxiliar como `Products`, `Fleet`, `Logistics`, `Summary`), as colunas são inferidas automaticamente a partir das chaves do dataset daquela aba.
		- Se uma aba tiver colunas suficientes aplicáveis (incluindo colunas calculadas como `formula`, `percentageOf`, `sheetFormula` e `cellType: Lookup`), a aba usa as colunas configuradas.
		- Isso evita abas cheias de colunas vazias quando você usa um `layout.columns` focado em uma aba “principal” (ex.: `Purchases`).

Cross-sheet formulas:

- Dentro de `formula`/`sheetFormula`, você pode usar o token `sheet:<NomeDaAba>.<field>`.
	- Ex.: `SUM(sheet:Employees.salary)` → `SUM(Employees!B2:B10)` (o range é resolvido automaticamente com base nas linhas geradas).

#### Passo a passo (Excel multi-abas + Lookup + charts)

1) Defina `layout.sheets[]` e coloque cada dataset em `data` com a chave igual ao `dataSource`.

2) Se você quiser uma aba “principal” (ex.: `Purchases`) com colunas calculadas, use `layout.columns` global.
	- As abas auxiliares (`Products`, `Fleet`, `Logistics`, `FinanceSummary`) podem ficar sem `layout.columns` específico — o serviço vai inferir as colunas nelas automaticamente quando o layout global não for aplicável.

3) Para preencher um campo via busca em outra aba, use `cellType: "Lookup"` + `lookup`:
	- `sheet`: nome da aba onde está a “tabela de referência”
	- `keyField`: campo do dataset atual que contém a chave (ex.: `productId`)
	- `lookupField`: campo da aba de referência que contém a chave (ex.: `id`)
	- `returnField`: campo da aba de referência que será retornado (ex.: `name`)
	- `engine`: `vlookup` (compatível) ou `xlookup` (mais flexível)
	- `matchMode`: `exact` (padrão) ou variantes

Importante:
	- Você NÃO precisa expor `keyField` como uma coluna visível para o Lookup funcionar.
		- Se `keyField` não existir nas colunas renderizadas da aba, o serviço usa o valor da linha como literal na fórmula (`VLOOKUP(2,...)`).
		- Se `keyField` existir como coluna, o serviço referencia a célula (`VLOOKUP(A2,...)`) — bom para cenários em que o usuário edita a planilha.
	- `engine: "vlookup"` exige que `returnField` esteja à direita de `lookupField` na aba de referência; caso contrário, use `engine: "xlookup"`.

4) Para charts, use `layout.charts[]` e aponte `categoryField`/`valueField` para campos existentes na aba.

Exemplo completo (baseado no seu caso):

```json
{
	"layout": {
		"freezeHeader": true,
		"autoSizeColumns": true,
		"groupBy": "supplier",
		"sheets": [
			{ "name": "Products", "dataSource": "products" },
			{ "name": "Purchases", "dataSource": "purchases" },
			{ "name": "Fleet", "dataSource": "fleet" },
			{ "name": "Logistics", "dataSource": "logistics" },
			{ "name": "FinanceSummary", "dataSource": "summary" }
		],
		"columns": [
			{ "field": "id", "title": "ID", "backgroundColor": "#F2F2F2" },
			{
				"field": "product",
				"title": "Product",
				"cellType": "Lookup",
				"lookup": {
					"sheet": "Products",
					"keyField": "productId",
					"lookupField": "id",
					"returnField": "name",
					"matchMode": "exact",
					"engine": "vlookup"
				}
			},
			{ "field": "supplier", "title": "Supplier", "backgroundColor": "#E8F4FF" },
			{ "field": "qty", "title": "Quantity", "format": "number" },
			{ "field": "price", "title": "Unit Price", "format": "currency", "textColor": "#003366" },
			{
				"field": "total",
				"title": "Total Cost",
				"formula": "qty * price",
				"aggregate": "sum",
				"format": "currency",
				"backgroundColor": "#FFF4E5",
				"conditionalFormatting": [
					{
						"operator": "greaterThan",
						"value": "10000",
						"backgroundColor": "#FFD6D6",
						"textColor": "#9A0511"
					}
				]
			},
			{ "field": "percent", "title": "% Supplier", "percentageOf": "total", "format": "percentage" }
		],
		"charts": [
			{
				"type": "column",
				"title": "Cost by Supplier",
				"sheet": "Purchases",
				"position": "H2",
				"categoryField": "supplier",
				"valueField": "total"
			},
			{
				"type": "pie",
				"title": "Fleet Cost Distribution",
				"sheet": "Fleet",
				"position": "H20",
				"categoryField": "vehicle",
				"valueField": "cost"
			}
		]
	},
	"data": {
		"products": [
			{ "id": 1, "name": "Server Dell R740" },
			{ "id": 2, "name": "Laptop Lenovo T14" },
			{ "id": 3, "name": "Cisco Switch 24p" }
		],
		"purchases": [
			{ "productId": 1, "supplier": "Dell", "qty": 2, "price": 15000 },
			{ "productId": 2, "supplier": "Lenovo", "qty": 5, "price": 7000 },
			{ "productId": 3, "supplier": "Cisco", "qty": 3, "price": 12000 },
			{ "productId": 2, "supplier": "Lenovo", "qty": 10, "price": 6800 }
		],
		"fleet": [
			{ "vehicle": "Truck Volvo FH", "cost": 120000 },
			{ "vehicle": "Van Mercedes Sprinter", "cost": 90000 },
			{ "vehicle": "Pickup Hilux", "cost": 75000 }
		],
		"logistics": [
			{ "route": "SP → RJ", "distance": 430, "cost": 3200 },
			{ "route": "SP → MG", "distance": 590, "cost": 4200 },
			{ "route": "SP → PR", "distance": 410, "cost": 3000 }
		],
		"summary": [
			{ "label": "Total Purchases", "value": "", "formula": "SUM(sheet:Purchases.total)" },
			{ "label": "Fleet Investment", "value": "", "formula": "SUM(sheet:Fleet.cost)" },
			{ "label": "Logistics Cost", "value": "", "formula": "SUM(sheet:Logistics.cost)" }
		]
	}
}
```

Nota sobre `summary`:
	- Se você quiser que a coluna `value` realmente calcule a fórmula, a forma recomendada é:
		- definir uma coluna `value` com `cellType: "Formula"` e colocar a expressão Excel dentro do próprio campo `value` (ex.: `"value": "=SUM(sheet:Purchases.total)"`), ou
		- usar `sheetFormula` na coluna (quando o layout for específico para essa aba).

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

