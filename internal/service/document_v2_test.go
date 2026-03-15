package service

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
	"unicode/utf8"

	"Data2Doc/internal/models"
)

func TestV2PDFText_ISO8859_1Bytes(t *testing.T) {
	got := []byte(v2PDFText("ação"))
	// "ação" in ISO-8859-1 bytes: a=0x61, ç=0xE7, ã=0xE3, o=0x6F
	want := []byte{0x61, 0xE7, 0xE3, 0x6F}
	if !bytes.Equal(got, want) {
		t.Fatalf("unexpected bytes: got=%v want=%v", got, want)
	}
}

func TestV2RenderCSV_IsUTF8AndContainsAccents(t *testing.T) {
	svc := &DocumentService{}
	req := models.DocumentRequest{}
	req.Data.Default = models.DynamicData{
		Order: []string{"city"},
		Items: []map[string]any{{"city": "São Paulo"}},
	}

	cols := []v2Column{{Title: "Cidade", Field: "city"}}
	b, err := svc.v2RenderCSV(req, cols)
	if err != nil {
		t.Fatalf("v2RenderCSV error: %v", err)
	}
	if !utf8.Valid(b) {
		t.Fatalf("CSV output is not valid UTF-8")
	}
	if !strings.Contains(string(b), "São Paulo") {
		t.Fatalf("CSV output missing expected substring")
	}
}

func TestV2RenderWord_ContainsAccentsAndFooterConfig(t *testing.T) {
	svc := &DocumentService{}
	req := models.DocumentRequest{}
	req.Data.Default = models.DynamicData{
		Order: []string{"name"},
		Items: []map[string]any{{"name": "João"}},
	}
	req.Layout = &models.LayoutConfig{
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
	}

	cols := []v2Column{{Title: "Nome", Field: "name"}}
	b, err := svc.v2RenderWord(req, cols)
	if err != nil {
		t.Fatalf("v2RenderWord error: %v", err)
	}

	docXML := readZipEntry(t, b, "word/document.xml")
	if !utf8.Valid(docXML) {
		t.Fatalf("document.xml is not valid UTF-8")
	}
	if !strings.Contains(string(docXML), "João") {
		t.Fatalf("document.xml missing expected accent string")
	}
	if !strings.Contains(string(docXML), "w:orient=\"landscape\"") {
		t.Fatalf("document.xml missing landscape orientation")
	}

	footerXML := readZipEntry(t, b, "word/footer1.xml")
	if !strings.Contains(string(footerXML), "w:jc w:val=\"right\"") {
		t.Fatalf("footer1.xml missing right alignment")
	}
	if !strings.Contains(string(footerXML), "Page ") {
		t.Fatalf("footer1.xml missing 'Page ' prefix")
	}
}

func readZipEntry(t *testing.T, zipBytes []byte, name string) []byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(zipBytes), int64(len(zipBytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	for _, f := range r.File {
		if f.Name != name {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", name, err)
		}
		defer rc.Close()
		b, err := io.ReadAll(rc)
		if err != nil {
			t.Fatalf("read %s: %v", name, err)
		}
		return b
	}
	t.Fatalf("missing entry %s", name)
	return nil
}

func ptrBool(v bool) *bool { return &v }
