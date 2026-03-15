package service

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"Data2Doc/internal/models"
	"Data2Doc/internal/templates"

	"github.com/jung-kurt/gofpdf"
	"github.com/nguyenthenguyen/docx"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
)

var (
	ErrUnsupportedDocumentType = errors.New("unsupported document type")
	ErrInvalidPayload          = errors.New("invalid payload; expected JSON object or array of objects")
	ErrEmptyPayload            = errors.New("payload has no keys; send at least one field (e.g. [{\"name\":\"pedro\",\"age\":20}])")
)

type DocumentService struct {
	templates *templates.Loader
}

func NewDocumentService(loader *templates.Loader) *DocumentService {
	return &DocumentService{templates: loader}
}

func IsTemplateNotFound(err error) bool {
	return errors.Is(err, templates.ErrTemplateNotFound)
}

func (s *DocumentService) GenerateFromPayload(docType models.DocumentType, baseName string, templateName string, payload []byte) (filename string, contentType string, data []byte, err error) {
	if !docType.IsValid() {
		return "", "", nil, ErrUnsupportedDocumentType
	}

	safeID := strings.TrimSpace(baseName)
	if safeID == "" {
		safeID = "document"
	}

	items, headers, err := parsePayload(payload)
	if err != nil {
		return "", "", nil, err
	}

	switch docType {
	case models.DocumentExcel:
		contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		filename = fmt.Sprintf("%s.xlsx", safeID)
		data, err = s.generateExcel(items, headers, templateName)
		return
	case models.DocumentWord:
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
		filename = fmt.Sprintf("%s.docx", safeID)
		data, err = s.generateWord(items, headers, templateName)
		return
	case models.DocumentPDF:
		contentType = "application/pdf"
		filename = fmt.Sprintf("%s.pdf", safeID)
		data, err = s.generatePDF(items, headers, templateName)
		return
	default:
		return "", "", nil, ErrUnsupportedDocumentType
	}
}

func parsePayload(payload []byte) (items []map[string]any, headers []string, err error) {
	dec := json.NewDecoder(bytes.NewReader(payload))
	dec.UseNumber()

	first, err := dec.Token()
	if err != nil {
		return nil, nil, err
	}

	seen := map[string]bool{}
	addHeaders := func(order []string) {
		for _, k := range order {
			if !seen[k] {
				seen[k] = true
				headers = append(headers, k)
			}
		}
	}

	switch t := first.(type) {
	case json.Delim:
		switch t {
		case '[':
			for dec.More() {
				obj, order, err := decodeOrderedObject(dec, false)
				if err != nil {
					return nil, nil, err
				}
				items = append(items, obj)
				addHeaders(order)
			}
			end, err := dec.Token()
			if err != nil {
				return nil, nil, err
			}
			if d, ok := end.(json.Delim); !ok || d != ']' {
				return nil, nil, ErrInvalidPayload
			}
		case '{':
			obj, order, err := decodeOrderedObject(dec, true)
			if err != nil {
				return nil, nil, err
			}
			items = []map[string]any{obj}
			addHeaders(order)
		default:
			return nil, nil, ErrInvalidPayload
		}
	default:
		return nil, nil, ErrInvalidPayload
	}

	if len(items) == 0 {
		return nil, nil, ErrInvalidPayload
	}
	allEmpty := true
	for _, obj := range items {
		if len(obj) > 0 {
			allEmpty = false
			break
		}
	}
	if allEmpty {
		return nil, nil, ErrEmptyPayload
	}

	return items, headers, nil
}

func decodeOrderedObject(dec *json.Decoder, alreadyOpened bool) (map[string]any, []string, error) {
	if !alreadyOpened {
		tok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		d, ok := tok.(json.Delim)
		if !ok || d != '{' {
			return nil, nil, ErrInvalidPayload
		}
	}

	obj := map[string]any{}
	order := make([]string, 0)

	for dec.More() {
		keyTok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		key, ok := keyTok.(string)
		if !ok {
			return nil, nil, ErrInvalidPayload
		}
		order = append(order, key)

		var raw json.RawMessage
		if err := dec.Decode(&raw); err != nil {
			return nil, nil, err
		}
		val, err := decodeAnyUseNumber(raw)
		if err != nil {
			return nil, nil, err
		}
		obj[key] = val
	}

	endTok, err := dec.Token()
	if err != nil {
		return nil, nil, err
	}
	end, ok := endTok.(json.Delim)
	if !ok || end != '}' {
		return nil, nil, ErrInvalidPayload
	}

	return obj, order, nil
}

func decodeAnyUseNumber(raw json.RawMessage) (any, error) {
	dec := json.NewDecoder(bytes.NewReader(raw))
	dec.UseNumber()
	var v any
	if err := dec.Decode(&v); err != nil {
		return nil, err
	}
	return v, nil
}

func stringifyCellValue(v any) string {
	switch vv := v.(type) {
	case nil:
		return ""
	case string:
		return vv
	case json.Number:
		return vv.String()
	case bool:
		if vv {
			return "true"
		}
		return "false"
	default:
		b, err := json.Marshal(v)
		if err != nil {
			return fmt.Sprintf("%v", v)
		}
		return string(b)
	}
}

func formatParagraph(obj map[string]any, headers []string) string {
	parts := make([]string, 0, len(headers))
	for _, h := range headers {
		val, ok := obj[h]
		if !ok {
			continue
		}
		parts = append(parts, fmt.Sprintf("%s: %s", h, stringifyCellValue(val)))
	}
	if len(parts) == 0 {
		b, _ := json.Marshal(obj)
		return string(b)
	}
	return strings.Join(parts, "; ")
}

func (s *DocumentService) generateExcel(items []map[string]any, headers []string, templateName string) ([]byte, error) {
	var (
		f   *excelize.File
		err error
	)

	if strings.TrimSpace(templateName) != "" {
		_, path, loadErr := s.templates.Load(models.DocumentExcel, templateName)
		if loadErr != nil {
			return nil, loadErr
		}
		f, err = excelize.OpenFile(path)
		if err != nil {
			return nil, err
		}
	} else {
		f = excelize.NewFile()
	}
	defer func() { _ = f.Close() }()

	sheet := f.GetSheetName(0)
	if sheet == "" {
		sheet = "Sheet1"
		_ = f.SetSheetName(f.GetSheetName(0), sheet)
	}

	// Header
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		_ = f.SetCellValue(sheet, cell, h)
	}
	if len(headers) > 0 {
		styleID, err := f.NewStyle(&excelize.Style{
			Font: &excelize.Font{Bold: true},
			Alignment: &excelize.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
		})
		if err == nil {
			startCell, _ := excelize.CoordinatesToCellName(1, 1)
			endCell, _ := excelize.CoordinatesToCellName(len(headers), 1)
			_ = f.SetCellStyle(sheet, startCell, endCell, styleID)
		}
	}
	// Rows
	for r, obj := range items {
		for c, h := range headers {
			cell, _ := excelize.CoordinatesToCellName(c+1, r+2)
			_ = f.SetCellValue(sheet, cell, stringifyCellValue(obj[h]))
		}
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (s *DocumentService) generateWord(items []map[string]any, headers []string, templateName string) ([]byte, error) {
	textParts := make([]string, 0, len(items))
	for _, obj := range items {
		textParts = append(textParts, formatParagraph(obj, headers))
	}
	contentText := strings.Join(textParts, "\n")
	contentXML := buildWordParagraphsXML(textParts)

	var out bytes.Buffer
	if strings.TrimSpace(templateName) != "" {
		_, path, loadErr := s.templates.Load(models.DocumentWord, templateName)
		if loadErr != nil {
			return nil, loadErr
		}
		r, err := docx.ReadDocxFile(path)
		if err != nil {
			return nil, err
		}
		defer func() { _ = r.Close() }()

		d := r.Editable()
		// Preferred convention: template includes {{content_xml}} under <w:body>.
		// Fallback: replace {{content}} with plain text.
		d.ReplaceRaw("{{content_xml}}", contentXML, -1)
		_ = d.Replace("{{content}}", contentText, -1)
		if err := d.Write(&out); err != nil {
			return nil, err
		}
		return out.Bytes(), nil
	}

	// No template: build minimal DOCX by using an embedded empty template from library isn't available.
	// We'll create a temp docx from a tiny base template stored in code.
	// Simplest approach: create a blank docx file by copying a packaged template.
	// Here we generate from a minimal template in-memory (created once on disk).
	tmp, err := os.CreateTemp("", "data2doc-blank-*.docx")
	if err != nil {
		return nil, err
	}
	tmpPath := tmp.Name()
	_ = tmp.Close()
	defer func() { _ = os.Remove(tmpPath) }()

	if err := writeMinimalDocxTemplate(tmpPath); err != nil {
		return nil, err
	}

	r, err := docx.ReadDocxFile(tmpPath)
	if err != nil {
		return nil, err
	}
	defer func() { _ = r.Close() }()

	d := r.Editable()
	d.ReplaceRaw("{{content_xml}}", contentXML, -1)
	_ = d.Replace("{{content}}", contentText, -1)
	if err := d.Write(&out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func (s *DocumentService) generatePDF(items []map[string]any, headers []string, templateName string) ([]byte, error) {
	if strings.TrimSpace(templateName) != "" {
		b, _, loadErr := s.templates.Load(models.DocumentPDF, templateName)
		if loadErr != nil {
			return nil, loadErr
		}
		return b, nil
	}

	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.AddPage()
	pdf.SetFont("Arial", "", 12)
	for i, obj := range items {
		if i > 0 {
			pdf.Ln(4)
		}
		pdf.MultiCell(0, 6, toPDFText(formatParagraph(obj, headers)), "", "L", false)
	}

	var out bytes.Buffer
	if err := pdf.Output(&out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func writeMinimalDocxTemplate(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	zw := zip.NewWriter(f)
	defer func() { _ = zw.Close() }()

	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
		"word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
		"word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
	    {{content_xml}}
    <w:sectPr/>
  </w:body>
</w:document>`,
	}

	for name, content := range files {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := io.WriteString(w, content); err != nil {
			return err
		}
	}

	return zw.Close()
}

func escapeXMLText(s string) string {
	// minimal escaping for WordprocessingML text nodes
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

func buildWordParagraphsXML(paragraphs []string) string {
	if len(paragraphs) == 0 {
		return ""
	}
	var b strings.Builder
	for _, p := range paragraphs {
		b.WriteString(`<w:p><w:r><w:t xml:space="preserve">`)
		b.WriteString(escapeXMLText(p))
		b.WriteString(`</w:t></w:r></w:p>`)
	}
	return b.String()
}

func toPDFText(s string) string {
	// gofpdf core fonts expect ISO-8859-1 encoded input.
	out, err := charmap.ISO8859_1.NewEncoder().String(s)
	if err != nil {
		return s
	}
	return out
}
