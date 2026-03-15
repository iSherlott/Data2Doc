package data2doc_test

import (
	"archive/zip"
	"bytes"
	"io"
	"strconv"
	"strings"
	"testing"

	"Data2Doc/internal/models"
	"Data2Doc/internal/service"

	"github.com/xuri/excelize/v2"
)

func TestGenerateV2_Excel_WritesSheetsAndData(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			FreezeHeader:    true,
			AutoSizeColumns: true,
			Sheets: []models.SheetConfig{
				{Name: "Employees", DataSource: "employees"},
				{Name: "Departments", DataSource: "departments"},
			},
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Employee Name"},
				{Field: "salary", Title: "Salary", Format: models.ColFormatCurrency},
				{Field: "hired", Title: "Hire Date", Format: models.ColFormatDate},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {
				Items: []map[string]any{{"name": "Pedro Alves", "salary": 4500.50, "hired": "2022-01-10"}},
				Order: []string{"name", "salary", "hired"},
			},
			"departments": {
				Items: []map[string]any{{"department": "Financeiro", "employees": 12}},
				Order: []string{"department", "employees"},
			},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty xlsx bytes")
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	if idx, _ := f.GetSheetIndex("Employees"); idx < 0 {
		t.Fatalf("missing sheet Employees")
	}
	if idx, _ := f.GetSheetIndex("Departments"); idx < 0 {
		t.Fatalf("missing sheet Departments")
	}

	v, err := f.GetCellValue("Employees", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if v != "Employee Name" {
		t.Fatalf("unexpected header A1: %q", v)
	}
	v, err = f.GetCellValue("Employees", "A2")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if v != "Pedro Alves" {
		t.Fatalf("unexpected value A2: %q", v)
	}

	// When multiple sheets point to different datasets with different schemas,
	// the renderer should infer columns per-sheet if the configured layout.columns
	// do not match the dataset.
	v, err = f.GetCellValue("Departments", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if v != "department" {
		t.Fatalf("unexpected Departments header A1: %q", v)
	}
	v, err = f.GetCellValue("Departments", "A2")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if v != "Financeiro" {
		t.Fatalf("unexpected Departments value A2: %q", v)
	}
	v, err = f.GetCellValue("Departments", "B1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if v != "employees" {
		t.Fatalf("unexpected Departments header B1: %q", v)
	}
	v, err = f.GetCellValue("Departments", "B2")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	vv, err := strconv.ParseFloat(strings.ReplaceAll(v, ",", "."), 64)
	if err != nil {
		t.Fatalf("ParseFloat %q: %v", v, err)
	}
	if vv != 12 {
		t.Fatalf("unexpected Departments value B2: %v (raw=%q)", vv, v)
	}
}

func TestGenerateV2_Excel_Advanced_Features_WritesExpectedXML(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			FreezeHeader:    true,
			FreezeColumns:   1,
			HideEmptyRows:   true,
			MaxVisibleRows:  1,
			AutoSizeColumns: true,
			Columns: []models.ColumnConfig{
				{Field: "id", Title: "ID"},
				{Field: "secret", Title: "Secret", Hidden: true},
				{Field: "status", Title: "Status", CellType: models.ExcelCellSelect, Options: []string{"Active", "Inactive"}},
				{Field: "amount", Title: "Amount", CellType: models.ExcelCellCurrency, Locked: true},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"": {
				Items: []map[string]any{{"id": 1, "secret": "x", "status": "Active", "amount": 12.34}},
				Order: []string{"id", "secret", "status", "amount"},
			},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty xlsx bytes")
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	if !strings.Contains(sheetXML, "xSplit=\"1\"") || !strings.Contains(sheetXML, "ySplit=\"1\"") {
		t.Fatalf("expected panes with xSplit=1 and ySplit=1")
	}
	if !strings.Contains(sheetXML, "hidden=\"1\"") && !strings.Contains(sheetXML, "hidden=\"true\"") {
		t.Fatalf("expected hidden column in worksheet xml")
	}
	if !strings.Contains(sheetXML, "dataValidations") {
		t.Fatalf("expected dataValidations in worksheet xml")
	}
	if !strings.Contains(sheetXML, "Active") || !strings.Contains(sheetXML, "Inactive") {
		t.Fatalf("expected dropdown options in worksheet xml")
	}
}

func TestGenerateV2_Excel_Charts_WritesChartParts(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets: []models.SheetConfig{{Name: "Sales", DataSource: "sales"}},
			Columns: []models.ColumnConfig{
				{Field: "dept", Title: "Department"},
				{Field: "total", Title: "Total", Aggregate: "sum"},
			},
			Charts: []models.ExcelChartConfig{
				{
					Type:          models.ChartColumn,
					Title:         "Sales by Department",
					Sheet:         "Sales",
					Position:      "E2",
					CategoryField: "dept",
					ValueField:    "total",
				},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"sales": {
				Items: []map[string]any{{"dept": "A", "total": 10}, {"dept": "B", "total": 20}},
				Order: []string{"dept", "total"},
			},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty xlsx bytes")
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	chartParts := 0
	var chartXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/charts/") && strings.HasSuffix(f.Name, ".xml") {
			chartParts++
			if chartXML == "" {
				r, err := f.Open()
				if err != nil {
					t.Fatalf("open %s: %v", f.Name, err)
				}
				b, _ := io.ReadAll(r)
				_ = r.Close()
				chartXML = string(b)
			}
		}
	}
	if chartParts == 0 {
		t.Fatalf("expected at least one chart part under xl/charts/")
	}
	if chartXML == "" {
		t.Fatalf("expected readable chart xml")
	}
	if !strings.Contains(chartXML, "Sales") {
		t.Fatalf("expected chart xml to reference sheet 'Sales'")
	}
}

func TestGenerateV2_Excel_ConditionalFormatting_WritesWorksheetXML(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
				{Field: "amount", Title: "Amount", ConditionalFormatting: []models.ExcelConditionalFormattingRule{
					{Operator: models.ExcelCondGreaterThan, Value: "10", BackgroundColor: "#FEC7CE", TextColor: "#9A0511"},
				}},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"name": "A", "amount": 5}, {"name": "B", "amount": 20}}, Order: []string{"name", "amount"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}
	if !strings.Contains(sheetXML, "conditionalFormatting") {
		t.Fatalf("expected conditionalFormatting in worksheet xml")
	}
}

func TestGenerateV2_Excel_ValidationRange_WritesWorksheetXML(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets: []models.SheetConfig{{Name: "Products", DataSource: "products"}, {Name: "Orders", DataSource: "orders"}},
			Columns: []models.ColumnConfig{
				{Field: "product", Title: "Product", CellType: models.ExcelCellSelect, ValidationRange: "Products!A2:A3"},
				{Field: "qty", Title: "Qty", CellType: models.ExcelCellNumber},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"products": {Items: []map[string]any{{"name": "P1"}, {"name": "P2"}}, Order: []string{"name"}},
			"orders":   {Items: []map[string]any{{"product": "P1", "qty": 1}}, Order: []string{"product", "qty"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	ordersSheetXML := ""
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			xml := string(b)
			// The sheet order isn't stable across versions; just grab the one that contains the validation formula.
			if strings.Contains(xml, "dataValidations") && strings.Contains(xml, "Products") {
				ordersSheetXML = xml
				break
			}
		}
	}
	if ordersSheetXML == "" {
		t.Fatalf("expected worksheet xml with data validation referencing Products")
	}
}

func TestGenerateV2_Excel_CellTypeFormula_WritesFormula(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{
				{Field: "a", Title: "A", CellType: models.ExcelCellNumber},
				{Field: "b", Title: "B", CellType: models.ExcelCellFormula},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"a": 5, "b": "=A2*2"}}, Order: []string{"a", "b"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	fx, err := f.GetCellFormula("Sheet1", "B2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if !strings.Contains(fx, "A2*2") {
		t.Fatalf("expected formula to contain A2*2, got %q", fx)
	}
}

func TestGenerateV2_Excel_InferredColumns_LeadingEquals_WritesFormulaAndResolvesSheetTokens(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets: []models.SheetConfig{{Name: "Employees", DataSource: "employees"}, {Name: "FinanceSummary", DataSource: "summary"}},
			// Global columns apply to Employees, but not to FinanceSummary (it will infer label/value).
			Columns: []models.ColumnConfig{
				{Field: "employee", Title: "Employee"},
				{Field: "totalCompensation", Title: "Total", Format: models.ColFormatCurrency, Aggregate: "sum"},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {Items: []map[string]any{{"employee": "Ana", "totalCompensation": 1000}, {"employee": "Bruno", "totalCompensation": 2000}}, Order: []string{"employee", "totalCompensation"}},
			"summary":   {Items: []map[string]any{{"label": "Total Payroll", "value": "=SUM(sheet:Employees.totalCompensation)"}}, Order: []string{"label", "value"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	fx, err := f.GetCellFormula("FinanceSummary", "B2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if fx == "" {
		t.Fatalf("expected FinanceSummary!B2 to have a formula")
	}
	if strings.Contains(fx, "sheet:") {
		t.Fatalf("expected sheet tokens to be resolved, got %q", fx)
	}
	if !strings.Contains(fx, "Employees") {
		t.Fatalf("expected formula to reference Employees sheet, got %q", fx)
	}
	if !strings.Contains(strings.ToUpper(fx), "SUM(") {
		t.Fatalf("expected formula to contain SUM, got %q", fx)
	}
}

func TestGenerateV2_Excel_CellTypeLookup_WritesVlookupFormula(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			// Put Orders before Products to ensure lookups are resolved after all sheets are registered.
			Sheets: []models.SheetConfig{{Name: "Orders", DataSource: "orders"}, {Name: "Products", DataSource: "products"}},
			Columns: []models.ColumnConfig{
				{Field: "productId", Title: "Product ID", CellType: models.ExcelCellNumber},
				{Field: "productName", Title: "Product", CellType: models.ExcelCellLookup, Lookup: &models.ExcelLookupConfig{
					Sheet:       "Products",
					KeyField:    "productId",
					LookupField: "id",
					ReturnField: "name",
					MatchMode:   models.ExcelLookupMatchExact,
					Engine:      models.ExcelLookupEngineVLookup,
				}},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"products": {Items: []map[string]any{{"id": 1, "name": "P1"}, {"id": 2, "name": "P2"}}, Order: []string{"id", "name"}},
			"orders":   {Items: []map[string]any{{"productId": 2}}, Order: []string{"productId"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	fx, err := f.GetCellFormula("Orders", "B2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	strip := func(s string) string { return strings.ReplaceAll(strings.ReplaceAll(s, " ", ""), "\t", "") }
	got := strip(fx)
	if !strings.Contains(got, "VLOOKUP(") {
		t.Fatalf("expected VLOOKUP formula, got %q", fx)
	}
	if !strings.Contains(got, "VLOOKUP(A2") {
		t.Fatalf("expected VLOOKUP to reference key cell A2, got %q", fx)
	}
	if !strings.Contains(got, "'Products'!$A$2:$B$3") {
		t.Fatalf("expected VLOOKUP to reference Products range $A$2:$B$3, got %q", fx)
	}
	if !strings.Contains(got, ",2,FALSE)") {
		t.Fatalf("expected VLOOKUP colIndex=2 and exact match FALSE, got %q", fx)
	}
}

func TestGenerateV2_Excel_SheetsConfig_RemovesDefaultSheet1(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets: []models.SheetConfig{{Name: "Employees", DataSource: "employees"}},
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {Items: []map[string]any{{"name": "Ana"}}, Order: []string{"name"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	if idx, _ := f.GetSheetIndex("Employees"); idx < 0 {
		t.Fatalf("missing sheet Employees")
	}
	if idx, _ := f.GetSheetIndex("Sheet1"); idx >= 0 {
		t.Fatalf("did not expect default Sheet1 when layout.sheets is set")
	}
}

func TestGenerateV2_Excel_HideEmptyRows_HidesAllByDefault_AndHidesColsToXFD(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			HideEmptyRows:  true,
			MaxVisibleRows: 1,
			Columns: []models.ColumnConfig{
				{Field: "a", Title: "A"},
				{Field: "b", Title: "B"},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"a": "x", "b": "y"}}, Order: []string{"a", "b"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	// Rows hidden by default without enumerating 1,048,576 rows.
	if !strings.Contains(sheetXML, "zeroHeight=\"1\"") && !strings.Contains(sheetXML, "zeroHeight=\"true\"") {
		t.Fatalf("expected sheetFormatPr zeroHeight to be set")
	}

	// Columns after B should be hidden up to XFD (16384).
	if !strings.Contains(sheetXML, "max=\"16384\"") {
		t.Fatalf("expected hidden columns range up to XFD (max=16384) in worksheet xml")
	}
}

func TestGenerateV2_Excel_RowHeight_AutoAdjusts_ForWrappedText(t *testing.T) {
	svc := &service.DocumentService{}

	long := strings.Repeat("muito ", 25)

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{{Field: "txt", Title: "Texto", Width: 5}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"txt": long}}, Order: []string{"txt"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	// Expect an explicit height on the first data row (row r="2").
	// Depending on excelize version, this can appear as ht="..." with customHeight, or only ht="...".
	row2 := strings.Index(sheetXML, "<row")
	for row2 >= 0 {
		end := strings.Index(sheetXML[row2:], ">")
		if end < 0 {
			break
		}
		tag := sheetXML[row2 : row2+end+1]
		if strings.Contains(tag, `r="2"`) {
			if strings.Contains(tag, `ht="`) || strings.Contains(tag, `customHeight="1"`) || strings.Contains(tag, `customHeight="true"`) {
				return
			}
			t.Fatalf("expected row 2 to have explicit height/customHeight, got tag: %s", tag)
		}
		next := strings.Index(sheetXML[row2+end+1:], "<row")
		if next < 0 {
			break
		}
		row2 = row2 + end + 1 + next
	}

	t.Fatalf("expected to find row r=2 in worksheet xml")
}

func TestGenerateV2_Excel_SheetsConfig_AllowsSheetNamedSheet1_NoExtraSheets(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets:  []models.SheetConfig{{Name: "Sheet1", DataSource: "employees"}},
			Columns: []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {Items: []map[string]any{{"name": "Ana"}}, Order: []string{"name"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	list := f.GetSheetList()
	if len(list) != 1 || list[0] != "Sheet1" {
		t.Fatalf("expected only Sheet1, got %v", list)
	}
}

func TestGenerateV2_Excel_NoHideEmptyRows_DoesNotSetZeroHeight(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			HideEmptyRows: false,
			Columns:       []models.ColumnConfig{{Field: "a", Title: "A"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"a": "x"}}, Order: []string{"a"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	if strings.Contains(sheetXML, "zeroHeight=") {
		t.Fatalf("did not expect zeroHeight when HideEmptyRows/MaxVisibleRows are not set")
	}
}

func TestGenerateV2_Excel_HideEmptyRows_InferredColumns_HidesFromNextColumnToXFD(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			HideEmptyRows: true,
			// No layout.columns: infer columns from data order.
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"a": "x", "b": "y"}}, Order: []string{"a", "b"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	// After 2 inferred columns (A,B), hiding should start from column 3 (C) up to XFD (16384).
	if !strings.Contains(sheetXML, "max=\"16384\"") {
		t.Fatalf("expected hidden columns range up to XFD (max=16384) in worksheet xml")
	}
	if !strings.Contains(sheetXML, "min=\"3\"") {
		t.Fatalf("expected hidden columns to start at min=3 (column C) when 2 columns are generated")
	}
}

func TestGenerateV2_Excel_RowHeight_IsClampedToMax409Points(t *testing.T) {
	svc := &service.DocumentService{}

	// Force an absurd number of wrapped lines by using a tiny column width.
	veryLong := strings.Repeat("x", 20000)

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{{Field: "txt", Title: "Texto", Width: 1}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"txt": veryLong}}, Order: []string{"txt"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(gen.Bytes), int64(len(gen.Bytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var sheetXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml") {
			r, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", f.Name, err)
			}
			b, _ := io.ReadAll(r)
			_ = r.Close()
			sheetXML = string(b)
			break
		}
	}
	if sheetXML == "" {
		t.Fatalf("expected worksheet xml")
	}

	rowIdx := strings.Index(sheetXML, "<row")
	for rowIdx >= 0 {
		end := strings.Index(sheetXML[rowIdx:], ">")
		if end < 0 {
			break
		}
		tag := sheetXML[rowIdx : rowIdx+end+1]
		if strings.Contains(tag, `r="2"`) {
			// Excel stores row height in points as a float.
			pos := strings.Index(tag, `ht="`)
			if pos < 0 {
				t.Fatalf("expected row 2 to include ht attribute, got tag: %s", tag)
			}
			pos += len(`ht="`)
			endQ := strings.Index(tag[pos:], `"`)
			if endQ < 0 {
				t.Fatalf("malformed ht attribute in tag: %s", tag)
			}
			v := tag[pos : pos+endQ]
			ht, err := strconv.ParseFloat(v, 64)
			if err != nil {
				t.Fatalf("ParseFloat ht=%q: %v", v, err)
			}
			if ht > 409.0 {
				t.Fatalf("expected row height to be clamped to <=409, got %v", ht)
			}
			return
		}
		next := strings.Index(sheetXML[rowIdx+end+1:], "<row")
		if next < 0 {
			break
		}
		rowIdx = rowIdx + end + 1 + next
	}

	t.Fatalf("expected to find row r=2 in worksheet xml")
}

func TestGenerateV2_Excel_Stress_ManyRows_GeneratesAndIsReadable(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 2000)
	for i := 1; i <= 2000; i++ {
		items = append(items, map[string]any{
			"id":     i,
			"name":   "User",
			"salary": 1234.56,
			"active": true,
			"note":   "ok",
		})
	}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			FreezeHeader:    true,
			AutoSizeColumns: false,
			Columns: []models.ColumnConfig{
				{Field: "id", Title: "ID", Format: models.ColFormatNumber},
				{Field: "name", Title: "Name"},
				{Field: "salary", Title: "Salary", Format: models.ColFormatCurrency},
				{Field: "active", Title: "Active"},
				{Field: "note", Title: "Note"},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id", "name", "salary", "active", "note"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty xlsx bytes")
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	// Last row: header(1) + 2000 items => row 2001
	v, err := f.GetCellValue("Sheet1", "A2001")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	vv, err := strconv.ParseFloat(strings.ReplaceAll(v, ",", "."), 64)
	if err != nil {
		t.Fatalf("ParseFloat %q: %v", v, err)
	}
	if vv != 2000 {
		t.Fatalf("expected A2001 numeric value 2000, got %v (raw=%q)", vv, v)
	}
}

func TestGenerateV2_Excel_Calcs_FormulaPerRow_IsResolvedByFieldName(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{
				{Field: "price", Title: "Price", Format: models.ColFormatNumber},
				{Field: "qty", Title: "Qty", Format: models.ColFormatNumber},
				{Field: "total", Title: "Total", Formula: "price * qty", Format: models.ColFormatNumber},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{
			{"price": 10, "qty": 2},
			{"price": 5, "qty": 3},
		}, Order: []string{"price", "qty"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	strip := func(s string) string { return strings.ReplaceAll(s, " ", "") }

	fx, err := f.GetCellFormula("Sheet1", "C2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strip(fx) != "A2*B2" {
		t.Fatalf("unexpected C2 formula: %q", fx)
	}
	fx, err = f.GetCellFormula("Sheet1", "C3")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strip(fx) != "A3*B3" {
		t.Fatalf("unexpected C3 formula: %q", fx)
	}
}

func TestGenerateV2_Excel_Calcs_AggregateSum_AddsTotalRow(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
				{Field: "salary", Title: "Salary", Aggregate: "sum", Format: models.ColFormatNumber},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{
			{"name": "Ana", "salary": 10},
			{"name": "Pedro", "salary": 20},
		}, Order: []string{"name", "salary"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	label, err := f.GetCellValue("Sheet1", "A4")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "TOTAL" {
		t.Fatalf("expected total row label TOTAL, got %q", label)
	}

	fx, err := f.GetCellFormula("Sheet1", "B4")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "SUM(B2:B3)" {
		t.Fatalf("unexpected B4 total formula: %q", fx)
	}
}

func TestGenerateV2_Excel_Calcs_GroupBy_InsertsSubtotals_AndTotalsFromSubtotals(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			GroupBy: "department",
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
				{Field: "department", Title: "Department"},
				{Field: "salary", Title: "Salary", Aggregate: "sum", Format: models.ColFormatNumber},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{
			{"name": "Lucas", "department": "Technology", "salary": 7000},
			{"name": "Ana", "department": "Finance", "salary": 5000},
			{"name": "Pedro", "department": "Finance", "salary": 6000},
		}, Order: []string{"name", "department", "salary"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	// After sorting by department: Finance rows (2,3), subtotal at 4; Technology row at 5, subtotal at 6; TOTAL at 7.
	label, err := f.GetCellValue("Sheet1", "A4")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "Subtotal Finance" {
		t.Fatalf("expected subtotal label, got %q", label)
	}
	fx, err := f.GetCellFormula("Sheet1", "C4")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "SUM(C2:C3)" {
		t.Fatalf("unexpected subtotal Finance formula: %q", fx)
	}

	label, err = f.GetCellValue("Sheet1", "A6")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "Subtotal Technology" {
		t.Fatalf("expected subtotal label, got %q", label)
	}
	fx, err = f.GetCellFormula("Sheet1", "C6")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "SUM(C5:C5)" {
		t.Fatalf("unexpected subtotal Technology formula: %q", fx)
	}

	label, err = f.GetCellValue("Sheet1", "A7")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "TOTAL" {
		t.Fatalf("expected TOTAL label, got %q", label)
	}
	fx, err = f.GetCellFormula("Sheet1", "C7")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "SUM(C4,C6)" {
		t.Fatalf("unexpected TOTAL formula from subtotals: %q", fx)
	}
}

func TestGenerateV2_Excel_Calcs_GroupBy_MultiSheet_IgnoresSheetsWithoutField(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			GroupBy: "supplier",
			Sheets: []models.SheetConfig{
				{Name: "Purchases", DataSource: "purchases"},
				{Name: "Summary", DataSource: "summary"},
			},
			Columns: []models.ColumnConfig{
				{Field: "supplier", Title: "Supplier"},
				{Field: "qty", Title: "Qty", Format: models.ColFormatNumber},
				{Field: "price", Title: "Unit Price", Format: models.ColFormatNumber},
				{Field: "total", Title: "Total", Formula: "qty * price", Aggregate: "sum", Format: models.ColFormatNumber},
				{Field: "percent", Title: "% Supplier", PercentageOf: "total", Format: models.ColFormatPercentage},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"purchases": {
				Items: []map[string]any{
					{"supplier": "Dell", "qty": 2, "price": 15000},
					{"supplier": "HP", "qty": 4, "price": 900},
				},
				Order: []string{"supplier", "qty", "price"},
			},
			"summary": {
				Items: []map[string]any{{"label": "Total Purchases", "value": 0}},
				Order: []string{"label", "value"},
			},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	// Ensure the Summary sheet is present and did not fail due to missing groupBy field.
	if idx, _ := f.GetSheetIndex("Summary"); idx < 0 {
		t.Fatalf("missing sheet Summary")
	}

	// Purchases grouping should still happen.
	label, err := f.GetCellValue("Purchases", "A3")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "Subtotal Dell" {
		t.Fatalf("expected Purchases subtotal label, got %q", label)
	}

	// TOTAL row expected at row 6 in this 2-supplier scenario.
	label, err = f.GetCellValue("Purchases", "A6")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if label != "TOTAL" {
		t.Fatalf("expected Purchases TOTAL label, got %q", label)
	}

	fx, err := f.GetCellFormula("Purchases", "D6")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "SUM(D3,D5)" {
		t.Fatalf("unexpected Purchases TOTAL formula from subtotals: %q", fx)
	}

	fx, err = f.GetCellFormula("Purchases", "E2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "D2/$D$6" {
		t.Fatalf("unexpected Purchases percentage formula E2: %q", fx)
	}

	fx, err = f.GetCellFormula("Purchases", "E4")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "D4/$D$6" {
		t.Fatalf("unexpected Purchases percentage formula E4: %q", fx)
	}
}

func TestGenerateV2_Excel_Calcs_PercentageOf_UsesAbsoluteTotalReference(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
				{Field: "salary", Title: "Salary", Aggregate: "sum", Format: models.ColFormatNumber},
				{Field: "salaryPercentage", Title: "% of Total", PercentageOf: "salary"},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{
			{"name": "Ana", "salary": 10},
			{"name": "Pedro", "salary": 20},
		}, Order: []string{"name", "salary"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	// Total row should be row 4.
	fx, err := f.GetCellFormula("Sheet1", "C2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "B2/$B$4" {
		t.Fatalf("unexpected percentage formula C2: %q", fx)
	}
	fx, err = f.GetCellFormula("Sheet1", "C3")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	if strings.ReplaceAll(fx, " ", "") != "B3/$B$4" {
		t.Fatalf("unexpected percentage formula C3: %q", fx)
	}
}

func TestGenerateV2_Excel_Calcs_CrossSheet_SheetToken_IsResolvedToRange(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatExcel,
		Layout: &models.LayoutConfig{
			Sheets: []models.SheetConfig{
				{Name: "Employees", DataSource: "employees"},
				{Name: "Summary", DataSource: "summary"},
			},
			Columns: []models.ColumnConfig{
				{Field: "name", Title: "Name"},
				{Field: "salary", Title: "Salary", Aggregate: "sum", Format: models.ColFormatNumber},
				{Field: "employeeTotal", Title: "Employee Total", SheetFormula: "SUM(sheet:Employees.salary)"},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {
				Items: []map[string]any{{"name": "Ana", "salary": 1234}},
				Order: []string{"name", "salary"},
			},
			"summary": {
				Items: []map[string]any{{"employeeTotal": ""}},
				Order: []string{"employeeTotal"},
			},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	f, err := excelize.OpenReader(bytes.NewReader(gen.Bytes))
	if err != nil {
		t.Fatalf("excelize.OpenReader: %v", err)
	}
	defer func() { _ = f.Close() }()

	fx, err := f.GetCellFormula("Summary", "C2")
	if err != nil {
		t.Fatalf("GetCellFormula: %v", err)
	}
	// Employees sheet: salary is column B, and with 1 row the range is B2:B2.
	if strings.ReplaceAll(fx, " ", "") != "SUM('Employees'!B2:B2)" {
		t.Fatalf("unexpected Summary C2 formula: %q", fx)
	}
}
