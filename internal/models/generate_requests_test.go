package models

import (
	"encoding/json"
	"testing"
)

func TestExcelGenerateRequest_PromotesSingleDatasetToDefault(t *testing.T) {
	var req ExcelGenerateRequest
	b := []byte(`{"data": {"rows": [{"name": "Ana"}]}}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err != nil {
		t.Fatalf("Validate: %v", err)
	}
	if req.Data.Default.IsEmpty() {
		t.Fatalf("expected default dataset to be populated")
	}
	if len(req.Data.Default.Items) != 1 {
		t.Fatalf("expected 1 item, got %d", len(req.Data.Default.Items))
	}
	if got := req.Data.Default.Items[0]["name"]; got != "Ana" {
		t.Fatalf("expected name=Ana, got=%v", got)
	}
}

func TestExcelGenerateRequest_MultipleDatasetsWithoutSheetsFails(t *testing.T) {
	var req ExcelGenerateRequest
	b := []byte(`{"data": {"a": [{"x": 1}], "b": [{"y": 2}]}}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err == nil {
		t.Fatalf("expected validation error")
	}
}

func TestExcelGenerateRequest_MultipleDatasetsWithSheetsOK(t *testing.T) {
	var req ExcelGenerateRequest
	b := []byte(`{
		"layout": {"sheets": [{"name": "A", "dataSource": "a"}, {"name": "B", "dataSource": "b"}]},
		"data": {"a": [{"x": 1}], "b": [{"y": 2}]}
	}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err != nil {
		t.Fatalf("Validate: %v", err)
	}
}
