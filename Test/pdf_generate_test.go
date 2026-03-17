package data2doc_test

import (
	"bytes"
	"encoding/hex"
	"math"
	"strings"
	"testing"

	"Data2Doc/internal/models"
	"Data2Doc/internal/service"

	"golang.org/x/text/encoding/charmap"
)

func TestGenerateV2_PDF_AutoPagination_ByHeight(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 220)
	for i := 1; i <= 220; i++ {
		items = append(items, map[string]any{"id": i, "text": "Linha de teste para paginação"})
	}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			PageBreak:  &models.PageBreakConfig{Enabled: true}, // rowsPerPage omitted => automatic
			Columns: []models.ColumnConfig{
				{Field: "id", Title: "ID"},
				{Field: "text", Title: "Texto"},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id", "text"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	p := countPDFPages(gen.Bytes)
	if p <= 1 {
		t.Fatalf("expected multiple pages, got %d", p)
	}
}

func TestGenerateV2_PDF_RowsPerPage_PaginatesPredictably(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 200)
	for i := 1; i <= 200; i++ {
		items = append(items, map[string]any{"id": i, "text": "Linha " + string(rune('0'+(i%10)))})
	}

	rowsPerPage := 30
	expectedPages := int(math.Ceil(float64(len(items)) / float64(rowsPerPage)))

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			PageBreak:  &models.PageBreakConfig{Enabled: true, RowsPerPage: rowsPerPage},
			Columns:    []models.ColumnConfig{{Field: "id", Title: "ID"}, {Field: "text", Title: "Texto"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id", "text"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	p := countPDFPages(gen.Bytes)
	if p != expectedPages {
		t.Fatalf("expected %d pages, got %d", expectedPages, p)
	}
}

func TestGenerateV2_PDF_Encoding_Portuguese_Strings_AreNotMojibake(t *testing.T) {
	svc := &service.DocumentService{}

	samples := []string{"ação", "produção", "informação", "São Paulo", "João", "educação", "coração"}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{Columns: []models.ColumnConfig{{Field: "txt", Title: "Texto"}}},
		Data: models.DataPayload{Default: models.DynamicData{
			Items: func() []map[string]any {
				items := make([]map[string]any, 0, len(samples))
				for _, s := range samples {
					items = append(items, map[string]any{"txt": s})
				}
				return items
			}(),
			Order: []string{"txt"},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	streams := extractFlateStreams(gen.Bytes)
	if len(streams) == 0 {
		t.Fatalf("expected at least one flate stream")
	}
	joined := bytes.Join(streams, []byte("\n"))

	enc := charmap.ISO8859_1.NewEncoder()
	for _, s := range samples {
		iso, _ := enc.Bytes([]byte(s))
		escaped := []byte(pdfLiteralEscapedISO88591(s))
		raw := iso
		hexUpper := []byte("<" + strings.ToUpper(hex.EncodeToString(raw)) + ">")
		hexLower := []byte("<" + strings.ToLower(hex.EncodeToString(raw)) + ">")
		if !bytes.Contains(joined, escaped) && !bytes.Contains(joined, raw) && !bytes.Contains(joined, hexUpper) && !bytes.Contains(joined, hexLower) {
			t.Fatalf("expected PDF content to contain %q in some ISO8859-1 form", s)
		}
	}

	// Ensure common mojibake sequences are not present.
	if bytes.Contains(joined, []byte("S\\303\\243o")) || bytes.Contains(joined, []byte("S\\303\\203\\302\\243o")) {
		t.Fatalf("detected mojibake-like escape sequences in PDF stream")
	}
}

func TestGenerateV2_PDF_Encoding_UnicodePunctuation_IsSanitized(t *testing.T) {
	svc := &service.DocumentService{}

	input := "Análise — resumo… 2023–2024"
	// Expected: Unicode punctuation normalized to ASCII that is safe for ISO-8859-1.
	expected := "Análise - resumo... 2023-2024"

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{Columns: []models.ColumnConfig{{Field: "txt", Title: "Texto"}}},
		Data: models.DataPayload{Default: models.DynamicData{
			Order: []string{"txt"},
			Items: []map[string]any{{"txt": input}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	streams := extractFlateStreams(gen.Bytes)
	if len(streams) == 0 {
		t.Fatalf("expected at least one flate stream")
	}
	joined := bytes.Join(streams, []byte("\n"))

	// UTF-8 sequences for: em dash (E2 80 94), ellipsis (E2 80 A6), en dash (E2 80 93)
	if bytes.Contains(joined, []byte{0xE2, 0x80, 0x94}) || bytes.Contains(joined, []byte{0xE2, 0x80, 0xA6}) || bytes.Contains(joined, []byte{0xE2, 0x80, 0x93}) {
		t.Fatalf("unexpected UTF-8 unicode punctuation bytes in PDF stream")
	}

	escaped := []byte(pdfLiteralEscapedISO88591(expected))
	raw, _ := charmap.ISO8859_1.NewEncoder().Bytes([]byte(expected))
	hexUpper := []byte("<" + strings.ToUpper(hex.EncodeToString(raw)) + ">")
	hexLower := []byte("<" + strings.ToLower(hex.EncodeToString(raw)) + ">")
	if !bytes.Contains(joined, escaped) && !bytes.Contains(joined, raw) && !bytes.Contains(joined, hexUpper) && !bytes.Contains(joined, hexLower) {
		t.Fatalf("expected PDF content to contain sanitized text %q", expected)
	}
}

const tinyPNGBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII="

func TestGenerateV2_PDF_Blocks_MixedContent_Renders(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			PageBreak:  &models.PageBreakConfig{Enabled: true},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockSectionTitle, Content: "Relatório Financeiro"},
				{Type: models.PDFBlockText, Content: "Este relatório apresenta os dados consolidados da empresa."},
				{Type: models.PDFBlockSpacer, Height: 8},
				{Type: models.PDFBlockImage, Data: tinyPNGBase64, Width: 40, Height: 12, Alignment: models.AlignCenter},
				{Type: models.PDFBlockSpacer, Height: 8},
				{
					Type:       models.PDFBlockTable,
					DataSource: "employees",
					Columns: []models.PDFTableColumnConfig{
						{Field: "name", Title: "Nome"},
						{Field: "salary", Title: "Salário"},
					},
				},
				{Type: models.PDFBlockSpacer, Height: 8},
				{
					Type:          models.PDFBlockChart,
					ChartType:     models.ChartBar,
					DataSource:    "sales",
					CategoryField: "department",
					ValueField:    "total",
					Title:         "Vendas por Departamento",
					Height:        60,
				},
				{Type: models.PDFBlockText, Content: "Os resultados demonstram crescimento consistente."},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"employees": {Items: []map[string]any{{"name": "Pedro", "salary": 5000}, {"name": "Ana", "salary": 6200}}, Order: []string{"name", "salary"}},
			"sales":     {Items: []map[string]any{{"department": "Financeiro", "total": 120000}, {"department": "Tecnologia", "total": 180000}}, Order: []string{"department", "total"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty PDF")
	}
	if !bytes.Contains(gen.Bytes, []byte("/Subtype /Image")) && !bytes.Contains(gen.Bytes, []byte("/Subtype/Image")) {
		t.Fatalf("expected PDF to contain an embedded image XObject")
	}
	if p := countPDFPages(gen.Bytes); p < 1 {
		t.Fatalf("expected at least 1 PDF page, got %d", p)
	}
}

func TestGenerateV2_PDF_Blocks_PageBreak_ForcesNewPage(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockText, Content: "Antes"},
				{Type: models.PDFBlockPageBreak},
				{Type: models.PDFBlockText, Content: "Depois"},
			},
		},
		Data: models.DataPayload{},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if p := countPDFPages(gen.Bytes); p != 2 {
		t.Fatalf("expected 2 pages, got %d", p)
	}
}

func TestGenerateV2_PDF_Blocks_Index_RendersAndCreatesLinks(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockIndex, Content: "Índice"},
				{Type: models.PDFBlockSectionTitle, Content: "Seção A"},
				{Type: models.PDFBlockText, Content: "Conteúdo"},
			},
		},
		Data: models.DataPayload{},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if p := countPDFPages(gen.Bytes); p < 2 {
		t.Fatalf("expected at least 2 pages (Index + content), got %d", p)
	}
	if !bytes.Contains(gen.Bytes, []byte("/Subtype /Link")) && !bytes.Contains(gen.Bytes, []byte("/Subtype/Link")) {
		t.Fatalf("expected PDF to contain link annotations for Index entries")
	}
}

func TestGenerateV2_PDF_Blocks_DataSourceMissing_ReturnsError(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			Blocks: []models.PDFBlockConfig{{Type: models.PDFBlockTable, DataSource: "missing", Columns: []models.PDFTableColumnConfig{{Field: "a"}}}},
		},
		Data: models.DataPayload{},
	}

	_, err := svc.Generate(req)
	if err == nil {
		t.Fatalf("expected error")
	}
	if !strings.Contains(strings.ToLower(err.Error()), "datasource") {
		t.Fatalf("expected datasource error, got: %v", err)
	}
}

func TestGenerateV2_PDF_Blocks_ChartForcesNewPage_WhenNotEnoughSpace(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 90)
	for i := 1; i <= 90; i++ {
		items = append(items, map[string]any{"id": i, "txt": "Linha de teste"})
	}

	layout := &models.LayoutConfig{
		PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
		PageBreak:  &models.PageBreakConfig{Enabled: true, RowsPerPage: 30},
		Blocks: []models.PDFBlockConfig{{
			Type:       models.PDFBlockTable,
			DataSource: "default",
			Columns:    []models.PDFTableColumnConfig{{Field: "id", Title: "ID"}, {Field: "txt", Title: "Texto"}},
		}},
	}

	baseReq := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: layout,
		Data:   models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id", "txt"}}},
	}

	genA, err := svc.Generate(baseReq)
	if err != nil {
		t.Fatalf("GenerateV2 A: %v", err)
	}
	pagesA := countPDFPages(genA.Bytes)
	if pagesA != int(math.Ceil(90.0/30.0)) {
		t.Fatalf("expected 3 pages for table-only, got %d", pagesA)
	}

	layoutWithChart := *layout
	layoutWithChart.Blocks = append(layoutWithChart.Blocks, models.PDFBlockConfig{
		Type:          models.PDFBlockChart,
		ChartType:     models.ChartColumn,
		DataSource:    "default",
		CategoryField: "id",
		ValueField:    "id",
		Title:         "Grafico",
		Height:        260,
	})
	chartReq := baseReq
	chartReq.Layout = &layoutWithChart

	genB, err := svc.Generate(chartReq)
	if err != nil {
		t.Fatalf("GenerateV2 B: %v", err)
	}
	pagesB := countPDFPages(genB.Bytes)
	if pagesB < pagesA+1 {
		t.Fatalf("expected chart to add at least one page (A=%d, B=%d)", pagesA, pagesB)
	}
}

func TestGenerateV2_PDF_Blocks_PageBreak_WinsOverRowsPerPageLogic(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			PageBreak:  &models.PageBreakConfig{Enabled: true, RowsPerPage: 1},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockText, Content: "Start"},
				{
					Type:       models.PDFBlockTable,
					DataSource: "default",
					Columns:    []models.PDFTableColumnConfig{{Field: "id", Title: "ID"}},
				},
				{Type: models.PDFBlockPageBreak},
				{Type: models.PDFBlockText, Content: "After"},
			},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"id": 1}, {"id": 2}}, Order: []string{"id"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	// Table has 2 rows and rowsPerPage=1 => at least 2 pages; explicit PageBreak adds one more.
	if p := countPDFPages(gen.Bytes); p != 3 {
		t.Fatalf("expected 3 pages, got %d", p)
	}
}

func TestGenerateV2_PDF_Blocks_OnlyTextAndImage_AllowsEmptyData(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockSectionTitle, Content: "Título"},
				{Type: models.PDFBlockText, Content: "Sem dados"},
				{Type: models.PDFBlockSpacer, Height: 6},
				{Type: models.PDFBlockImage, Data: tinyPNGBase64, Width: 40, Height: 12, Alignment: models.AlignCenter},
				{Type: models.PDFBlockText, Content: "Fim"},
			},
		},
		Data: models.DataPayload{},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if len(gen.Bytes) == 0 {
		t.Fatalf("expected non-empty PDF")
	}
	if !bytes.Contains(gen.Bytes, []byte("/Subtype /Image")) && !bytes.Contains(gen.Bytes, []byte("/Subtype/Image")) {
		t.Fatalf("expected PDF to contain an embedded image XObject")
	}
}

func TestGenerateV2_PDF_Blocks_Image_InvalidData_Fails(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			Blocks: []models.PDFBlockConfig{{Type: models.PDFBlockImage, Data: "%%%", Width: 40, Height: 12}},
		},
		Data: models.DataPayload{},
	}

	_, err := svc.Generate(req)
	if err == nil {
		t.Fatalf("expected error")
	}
}

func TestGenerateV2_PDF_Stress_TableRowsPerPage_ProducesExpectedPages(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 600)
	for i := 1; i <= 600; i++ {
		items = append(items, map[string]any{"id": i, "txt": "linha"})
	}

	rowsPerPage := 50
	expectedPages := int(math.Ceil(float64(len(items)) / float64(rowsPerPage)))

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			PageBreak:  &models.PageBreakConfig{Enabled: true, RowsPerPage: rowsPerPage},
			Blocks: []models.PDFBlockConfig{{
				Type:       models.PDFBlockTable,
				DataSource: "default",
				Columns: []models.PDFTableColumnConfig{
					{Field: "id", Title: "ID"},
					{Field: "txt", Title: "Texto"},
				},
			}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id", "txt"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	got := countPDFPages(gen.Bytes)
	if got < expectedPages {
		t.Fatalf("expected at least %d pages, got %d", expectedPages, got)
	}
	if got > expectedPages*4 {
		t.Fatalf("expected page count to be within a reasonable bound (<=%d), got %d", expectedPages*4, got)
	}
}

func TestGenerateV2_PDF_Chart_WithNonNumericValues_Fails(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatPDF,
		Layout: &models.LayoutConfig{
			Blocks: []models.PDFBlockConfig{{
				Type:          models.PDFBlockChart,
				ChartType:     models.ChartColumn,
				DataSource:    "default",
				CategoryField: "dept",
				ValueField:    "total",
				Title:         "Bad",
				Height:        60,
			}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"dept": "A", "total": "abc"}}, Order: []string{"dept", "total"}}},
	}

	_, err := svc.Generate(req)
	if err == nil {
		t.Fatalf("expected error")
	}
}
