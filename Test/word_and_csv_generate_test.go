package data2doc_test

import (
	"math"
	"strconv"
	"strings"
	"testing"
	"unicode/utf8"

	"Data2Doc/internal/models"
	"Data2Doc/internal/service"
)

func mmToTwips(mm float64) int {
	if mm < 0 {
		mm = 0
	}
	return int(math.Round(mm * 1440.0 / 25.4))
}

func TestGenerateV2_CSV_IsUTF8AndContainsAccents(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatCSV,
		Data: models.DataPayload{Default: models.DynamicData{
			Order: []string{"city"},
			Items: []map[string]any{{"city": "São Paulo"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	if !utf8.Valid(gen.Bytes) {
		t.Fatalf("CSV output is not valid UTF-8")
	}
	if !strings.Contains(string(gen.Bytes), "São Paulo") {
		t.Fatalf("CSV output missing expected substring")
	}
}

func TestGenerateV2_Word_ContainsAccentsAndFooterConfig(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageOrientation: models.PageLandscape,
			PageMargin:      &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			Footer: &models.FooterConfig{
				Show:      ptrBool(true),
				Alignment: models.AlignRight,
				PageNumber: &models.PageNumberConfig{
					Enabled: true,
					Format:  models.PageNumTextPageNum,
				},
			},
			Columns: []models.ColumnConfig{{Field: "name", Title: "Nome"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{
			Order: []string{"name"},
			Items: []map[string]any{{"name": "João"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	if !utf8.ValidString(docXML) {
		t.Fatalf("document.xml is not valid UTF-8")
	}
	if !strings.Contains(docXML, "João") {
		t.Fatalf("document.xml missing expected accent string")
	}
	if !strings.Contains(docXML, `w:orient="landscape"`) {
		t.Fatalf("document.xml missing landscape orientation")
	}

	footerXML := readZipEntryString(t, gen.Bytes, "word/footer1.xml")
	if !strings.Contains(footerXML, `w:jc w:val="right"`) {
		t.Fatalf("footer1.xml missing right alignment")
	}
	if !strings.Contains(footerXML, "Page ") {
		t.Fatalf("footer1.xml missing 'Page ' prefix")
	}
}

func TestGenerateV2_Word_TableWidth_RespectsMargins_AndCentersByDefault(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageOrientation:       models.PagePortrait,
			PageMargin:            &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			WordIgnorePageMargins: false,
			// WordCenterContent nil => default true
			Columns: []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{
			Items: []map[string]any{{"name": "Ana"}},
			Order: []string{"name"},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")

	pageWtw := 11906
	m := mmToTwips(10)
	expectedW := pageWtw - m - m

	if !strings.Contains(docXML, `w:tblW w:type="dxa" w:w="`+strconv.Itoa(expectedW)+`"`) {
		t.Fatalf("expected table width %d twips", expectedW)
	}
	if !strings.Contains(docXML, `<w:jc w:val="center"/>`) {
		t.Fatalf("expected centered table")
	}
}

func TestGenerateV2_Word_TableWidth_IgnoreMargins_UsesFullPageWidth(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageOrientation:       models.PagePortrait,
			PageMargin:            &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			WordIgnorePageMargins: true,
			WordCenterContent:     ptrBool(true),
			Columns:               []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{
			Items: []map[string]any{{"name": "Ana"}},
			Order: []string{"name"},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")

	pageWtw := 11906
	if !strings.Contains(docXML, `w:tblW w:type="dxa" w:w="`+strconv.Itoa(pageWtw)+`"`) {
		t.Fatalf("expected full page table width %d twips", pageWtw)
	}
}

func TestGenerateV2_Word_CenterContent_False_DoesNotWriteTableJc(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageOrientation:   models.PagePortrait,
			WordCenterContent: ptrBool(false),
			Columns:           []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{
			Items: []map[string]any{{"name": "Ana"}},
			Order: []string{"name"},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	start := strings.Index(docXML, `<w:tblPr>`)
	if start >= 0 {
		end := strings.Index(docXML[start:], `</w:tblPr>`)
		if end > 0 {
			block := docXML[start : start+end]
			if strings.Contains(block, `<w:jc w:val="center"/>`) {
				t.Fatalf("did not expect table jc center")
			}
		}
	}
}

func TestGenerateV2_Word_Orientation_Landscape_WritesPgSzOrient(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageOrientation: models.PageLandscape,
			Columns:         []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{
			Items: []map[string]any{{"name": "Ana"}},
			Order: []string{"name"},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}
	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")

	if !strings.Contains(docXML, `w:orient="landscape"`) {
		t.Fatalf("expected landscape orientation in sectPr")
	}
}

func TestGenerateV2_Word_Blocks_PageBreak_WritesWbrPage(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
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

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Fatalf("expected explicit page break in document.xml")
	}
}

func TestGenerateV2_Word_Blocks_MixedContent_WritesImagesAndHeading(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageMargin: &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockSectionTitle, Content: "Relatório"},
				{Type: models.PDFBlockText, Content: "Introdução"},
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
					ChartType:     models.ChartColumn,
					DataSource:    "sales",
					CategoryField: "department",
					ValueField:    "total",
					Title:         "Vendas",
					Height:        80,
				},
				{Type: models.PDFBlockPageBreak},
				{Type: models.PDFBlockText, Content: "Conclusão"},
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

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	if !strings.Contains(docXML, `<w:pStyle w:val="Heading1"/>`) {
		t.Fatalf("expected Heading1 style for SectionTitle")
	}
	if !strings.Contains(docXML, `<w:br w:type="page"/>`) {
		t.Fatalf("expected explicit page break")
	}

	relsXML := readZipEntryString(t, gen.Bytes, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, `Target="media/image1.png"`) {
		t.Fatalf("expected first embedded image relationship")
	}
	if !strings.Contains(relsXML, `Target="media/image2.png"`) {
		t.Fatalf("expected chart PNG embedded as second image")
	}

	_ = readZipEntry(t, gen.Bytes, "word/media/image1.png")
	_ = readZipEntry(t, gen.Bytes, "word/media/image2.png")
}

func TestGenerateV2_Word_Blocks_MultipleImages_HeaderImagePlusBlockImagePlusChart(t *testing.T) {
	svc := &service.DocumentService{}

	dataURI := "data:image/png;base64," + tinyPNGBase64

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageMargin:  &models.PageMarginConfig{Top: 10, Right: 10, Bottom: 10, Left: 10},
			HeaderImage: &models.HeaderImageConfig{Data: dataURI, Height: 10},
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockText, Content: "Antes"},
				{Type: models.PDFBlockImage, Data: dataURI, Width: 40, Height: 12, Alignment: models.AlignCenter},
				{
					Type:          models.PDFBlockChart,
					ChartType:     models.ChartPie,
					DataSource:    "sales",
					CategoryField: "department",
					ValueField:    "total",
					Title:         "Vendas",
					Height:        60,
				},
			},
		},
		Data: models.DataPayload{Sources: map[string]models.DynamicData{
			"sales": {Items: []map[string]any{{"department": "Financeiro", "total": 120000}, {"department": "Tecnologia", "total": 180000}}, Order: []string{"department", "total"}},
		}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	relsXML := readZipEntryString(t, gen.Bytes, "word/_rels/document.xml.rels")
	if !strings.Contains(relsXML, `Target="media/image1.png"`) {
		t.Fatalf("expected header image relationship")
	}
	if !strings.Contains(relsXML, `Target="media/image2.png"`) {
		t.Fatalf("expected block image relationship")
	}
	if !strings.Contains(relsXML, `Target="media/image3.png"`) {
		t.Fatalf("expected chart image relationship")
	}

	_ = readZipEntry(t, gen.Bytes, "word/media/image1.png")
	_ = readZipEntry(t, gen.Bytes, "word/media/image2.png")
	_ = readZipEntry(t, gen.Bytes, "word/media/image3.png")
}

func TestGenerateV2_Word_InvalidHeaderImageDataURI_Fails(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			HeaderImage: &models.HeaderImageConfig{Data: "data:image/png;base64,%%%", Height: 10},
			Columns:     []models.ColumnConfig{{Field: "name", Title: "Name"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"name": "Ana"}}, Order: []string{"name"}}},
	}

	_, err := svc.Generate(req)
	if err == nil {
		t.Fatalf("expected error")
	}
}

func TestGenerateV2_Word_Blocks_OnlyText_AllowsEmptyData(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockSectionTitle, Content: "Título"},
				{Type: models.PDFBlockText, Content: "Sem dados"},
				{Type: models.PDFBlockSpacer, Height: 5},
				{Type: models.PDFBlockText, Content: "Fim"},
			},
		},
		Data: models.DataPayload{},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	if !strings.Contains(docXML, "Sem dados") {
		t.Fatalf("expected document.xml to contain block text")
	}
}

func TestGenerateV2_Word_Blocks_Image_InvalidDataURI_Fails(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			Blocks: []models.PDFBlockConfig{
				{Type: models.PDFBlockImage, Data: "data:image/png;base64,%%%", Width: 40, Height: 10},
			},
		},
		Data: models.DataPayload{},
	}

	_, err := svc.Generate(req)
	if err == nil {
		t.Fatalf("expected error")
	}
}

func TestGenerateV2_Word_Legacy_RowsPerPage_InsertsPageBreaks(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageBreak: &models.PageBreakConfig{Enabled: true, RowsPerPage: 1},
			Columns:   []models.ColumnConfig{{Field: "id", Title: "ID"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: []map[string]any{{"id": 1}, {"id": 2}, {"id": 3}}, Order: []string{"id"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	// 3 chunks => 2 page breaks between tables.
	if got := strings.Count(docXML, `<w:br w:type="page"/>`); got != 2 {
		t.Fatalf("expected 2 page breaks, got %d", got)
	}
	if got := strings.Count(docXML, "<w:tbl>"); got != 3 {
		t.Fatalf("expected 3 tables (one per page chunk), got %d", got)
	}
}

func TestGenerateV2_Word_Stress_Legacy_RowsPerPage_ManyRows(t *testing.T) {
	svc := &service.DocumentService{}

	items := make([]map[string]any, 0, 120)
	for i := 1; i <= 120; i++ {
		items = append(items, map[string]any{"id": i})
	}

	rowsPerPage := 10
	expectedTables := int(math.Ceil(float64(len(items)) / float64(rowsPerPage)))
	expectedBreaks := expectedTables - 1

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
		Layout: &models.LayoutConfig{
			PageBreak: &models.PageBreakConfig{Enabled: true, RowsPerPage: rowsPerPage},
			Columns:   []models.ColumnConfig{{Field: "id", Title: "ID"}},
		},
		Data: models.DataPayload{Default: models.DynamicData{Items: items, Order: []string{"id"}}},
	}

	gen, err := svc.Generate(req)
	if err != nil {
		t.Fatalf("GenerateV2: %v", err)
	}

	docXML := readZipEntryString(t, gen.Bytes, "word/document.xml")
	if got := strings.Count(docXML, "<w:tbl>"); got != expectedTables {
		t.Fatalf("expected %d tables, got %d", expectedTables, got)
	}
	if got := strings.Count(docXML, `<w:br w:type="page"/>`); got != expectedBreaks {
		t.Fatalf("expected %d page breaks, got %d", expectedBreaks, got)
	}
}

func TestGenerateV2_Word_Blocks_Chart_WithNonNumericValues_Fails(t *testing.T) {
	svc := &service.DocumentService{}

	req := models.DocumentRequest{
		Format: models.DocumentFormatWord,
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
