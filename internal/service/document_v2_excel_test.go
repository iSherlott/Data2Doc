package service

import (
	"bytes"
	"testing"

	"Data2Doc/internal/models"

	"github.com/xuri/excelize/v2"
)

func TestGenerateV2_Excel_WritesSheetsAndData(t *testing.T) {
	svc := &DocumentService{}

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

	gen, err := svc.GenerateV2(req)
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

	if got := f.GetSheetName(0); got == "" {
		t.Fatalf("expected at least one sheet")
	}

	// Validate that our target sheets exist.
	if idx, _ := f.GetSheetIndex("Employees"); idx < 0 {
		t.Fatalf("missing sheet Employees")
	}
	if idx, _ := f.GetSheetIndex("Departments"); idx < 0 {
		t.Fatalf("missing sheet Departments")
	}

	// Validate data landed in Employees sheet.
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
}
