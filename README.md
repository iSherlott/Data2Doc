
# Data2Doc

API to generate documents from tabular data.

## Authentication

All endpoints are protected (Bearer token). Send:

- `Authorization: Bearer <token>`

## Endpoints

- `POST /generate/excel` → returns `.xlsx`
- `POST /generate/pdf` → returns `.pdf`
- `POST /generate/word` → returns `.docx`

Common query param:

- `id` (optional): base filename (without extension). If omitted, defaults to `X-Request-Id` header or `document`.

### POST /generate/excel

Body (`ExcelGenerateRequest`):

- `templateId` (optional)
- `layout` (optional): Excel-only options (sheets, freeze header, autosize columns, columns formatting)
- `data` (required): can be a single dataset or multiple datasets (for `layout.sheets`)

Single dataset example:

```json
{
	"templateId": "default",
	"layout": {
		"freezeHeader": true,
		"autoSizeColumns": true
	},
	"data": [
		{"name": "Ana", "salary": 1234.56},
		{"name": "Bruno", "salary": 2000}
	]
}
```

Multiple sheets example:

```json
{
	"layout": {
		"sheets": [
			{"name": "Employees", "dataSource": "employees"},
			{"name": "Departments", "dataSource": "departments"}
		]
	},
	"data": {
		"employees": [{"name": "Ana"}],
		"departments": [{"dept": "IT"}]
	}
}
```

### POST /generate/pdf

Body (`PDFGenerateRequest`):

- `templateId` (optional)
- `layout` (optional): PDF-only options (header image, footer pagination, margins/orientation)
- `data` (required): array/object of rows

Example:

```json
{
	"layout": {
		"footer": {
			"show": true,
			"alignment": "center",
			"pageNumber": {"enabled": true, "format": "Arabic"}
		}
	},
	"data": [
		{"name": "João", "age": 30},
		{"name": "María", "age": 28}
	]
}
```

### POST /generate/word

Body (`WordGenerateRequest`):

- `templateId` (optional)
- `layout` (optional): Word-only options (margins, page breaks, footer pagination)
- `data` (required): array/object of rows

Example:

```json
{
	"layout": {
		"pageBreak": {"enabled": true, "rowsPerPage": 25}
	},
	"data": [
		{"name": "Ana", "city": "São Paulo"},
		{"name": "Bruno", "city": "Porto"}
	]
}
```

## Curl examples

```bash
curl -X POST "http://localhost:8080/generate/pdf?id=report" \
	-H "Authorization: Bearer $TOKEN" \
	-H "Content-Type: application/json" \
	-o report.pdf \
	--data @payload.json
```

