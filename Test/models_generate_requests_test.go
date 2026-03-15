package data2doc_test

import (
	"encoding/json"
	"strings"
	"testing"

	"Data2Doc/internal/models"
)

func TestExcelGenerateRequest_PromotesSingleDatasetToDefault(t *testing.T) {
	var req models.ExcelGenerateRequest
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
	var req models.ExcelGenerateRequest
	b := []byte(`{"data": {"a": [{"x": 1}], "b": [{"y": 2}]}}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err == nil {
		t.Fatalf("expected validation error")
	}
}

func TestExcelGenerateRequest_MultipleDatasetsWithSheetsOK(t *testing.T) {
	var req models.ExcelGenerateRequest
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

func TestDataPayload_ObjectOfObjects_IsSources_WhenAllValuesAreComplex(t *testing.T) {
	var p models.DataPayload
	b := []byte(`{"a": {"x": 1}, "b": {"y": 2}}`)
	if err := json.Unmarshal(b, &p); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if !p.Default.IsEmpty() {
		t.Fatalf("expected default dataset to be empty")
	}
	if got := len(p.Sources); got != 2 {
		t.Fatalf("expected 2 sources, got %d", got)
	}
	a := p.Get("a")
	bb := p.Get("b")
	if a.IsEmpty() || bb.IsEmpty() {
		t.Fatalf("expected sources a and b to be non-empty")
	}
}

func TestDataPayload_ScalarObject_RemainsDefaultDataset(t *testing.T) {
	var p models.DataPayload
	b := []byte(`{"name": "Ana", "age": 30}`)
	if err := json.Unmarshal(b, &p); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if p.Default.IsEmpty() {
		t.Fatalf("expected default dataset to be populated")
	}
	if len(p.Sources) != 0 {
		t.Fatalf("expected no sources for scalar object")
	}
}

func TestExcelGenerateRequest_SheetsAcceptObjectDatasets(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{
		"layout": {"sheets": [{"name": "A", "dataSource": "a"}, {"name": "B", "dataSource": "b"}]},
		"data": {"a": {"x": 1}, "b": {"y": 2}}
	}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err != nil {
		t.Fatalf("Validate: %v", err)
	}
}

func TestExcelGenerateRequest_SheetsAcceptMixedArrayAndObjectDatasets(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{
		"layout": {"sheets": [{"name": "A", "dataSource": "a"}, {"name": "B", "dataSource": "b"}]},
		"data": {"a": [{"x": 1}], "b": {"y": 2}}
	}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err != nil {
		t.Fatalf("Validate: %v", err)
	}
}

func TestDataPayload_Null_IsEmpty(t *testing.T) {
	var p models.DataPayload
	b := []byte(`null`)
	if err := json.Unmarshal(b, &p); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if !p.IsEmpty() {
		t.Fatalf("expected payload to be empty")
	}
}

func TestDataPayload_EmptyObject_IsEmpty(t *testing.T) {
	var p models.DataPayload
	b := []byte(`{}`)
	if err := json.Unmarshal(b, &p); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if !p.Default.IsEmpty() {
		t.Fatalf("expected default dataset to be empty")
	}
	if len(p.Sources) != 0 {
		t.Fatalf("expected no sources")
	}
	if !p.IsEmpty() {
		t.Fatalf("expected payload to be empty")
	}
}

func TestDataPayload_ObjectOfEmptyObjects_DoesNotCreateSources(t *testing.T) {
	var p models.DataPayload
	b := []byte(`{"a": {}, "b": {}}`)
	if err := json.Unmarshal(b, &p); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if len(p.Sources) != 0 {
		t.Fatalf("expected empty sources when all datasets are empty objects")
	}
	if !p.IsEmpty() {
		t.Fatalf("expected payload to be empty")
	}
}

func TestExcelGenerateRequest_ObjectOfEmptyObjects_FailsAsDataRequired(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{"data": {"a": {}, "b": {}}}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err == nil {
		t.Fatalf("expected validation error")
	} else if !strings.Contains(strings.ToLower(err.Error()), "data is required") {
		t.Fatalf("expected data required error, got: %v", err)
	}
}

func TestExcelGenerateRequest_FreezeColumns_GT_InferredOrder_Fails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{
		"layout": {"freezeColumns": 3},
		"data": [{"a": 1, "b": 2}]
	}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err == nil {
		t.Fatalf("expected validation error")
	}
}

func TestExcelGenerateRequest_SheetRefersMissingDataSource_Fails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{
		"layout": {"sheets": [{"name": "A", "dataSource": "missing"}]},
		"data": {"a": [{"x": 1}]}
	}`)
	if err := json.Unmarshal(b, &req); err != nil {
		t.Fatalf("json.Unmarshal: %v", err)
	}
	if err := req.Validate(); err == nil {
		t.Fatalf("expected validation error")
	}
}

func TestDataPayload_WrongRootType_Number_Fails(t *testing.T) {
	var p models.DataPayload
	b := []byte(`123`)
	if err := json.Unmarshal(b, &p); err == nil {
		t.Fatalf("expected error")
	}
}

func TestDataPayload_ArrayOfNumbers_Fails(t *testing.T) {
	var p models.DataPayload
	b := []byte(`[1,2,3]`)
	if err := json.Unmarshal(b, &p); err == nil {
		t.Fatalf("expected error")
	}
}

func TestExcelGenerateRequest_WrongType_FreezeColumnsString_UnmarshalFails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{"layout": {"freezeColumns": "no"}, "data": [{"a": 1}]}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}

func TestExcelGenerateRequest_WrongType_ColumnFieldNumber_UnmarshalFails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{"layout": {"columns": [{"field": 123}]}, "data": [{"a": 1}]}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}

func TestExcelGenerateRequest_WrongType_OptionsString_UnmarshalFails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{"layout": {"columns": [{"field": "status", "cellType": "Select", "options": "Active"}]}, "data": [{"status": "Active"}]}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}

func TestExcelGenerateRequest_WrongType_DataString_UnmarshalFails(t *testing.T) {
	var req models.ExcelGenerateRequest
	b := []byte(`{"data": "oops"}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}

func TestPDFGenerateRequest_WrongType_DataString_UnmarshalFails(t *testing.T) {
	var req models.PDFGenerateRequest
	b := []byte(`{"data": "oops"}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}

func TestWordGenerateRequest_WrongType_PageBreakEnabledString_UnmarshalFails(t *testing.T) {
	var req models.WordGenerateRequest
	b := []byte(`{"layout": {"pageBreak": {"enabled": "true"}}, "data": [{"a": 1}]}`)
	if err := json.Unmarshal(b, &req); err == nil {
		t.Fatalf("expected json.Unmarshal error")
	}
}
