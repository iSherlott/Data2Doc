package service

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"math"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"Data2Doc/internal/models"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
)

type GeneratedDocument struct {
	Filename    string
	ContentType string
	Bytes       []byte
}

type v2Column struct {
	Field  string
	Title  string
	Width  float64
	Align  models.ColumnAlignmentEnum
	VAlign models.VerticalAlignmentEnum
	Format models.ColumnFormatEnum
}

func (s *DocumentService) GenerateV2(req models.DocumentRequest) (*GeneratedDocument, error) {
	if err := req.Validate(); err != nil {
		return nil, err
	}

	switch req.Format {
	case models.DocumentFormatExcel:
		b, err := s.v2RenderExcel(req)
		if err != nil {
			return nil, err
		}
		return &GeneratedDocument{Filename: "document.xlsx", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Bytes: b}, nil
	case models.DocumentFormatCSV:
		data := req.Data.Get("")
		cols := v2BuildColumns(req, data)
		b, err := s.v2RenderCSV(req, cols)
		if err != nil {
			return nil, err
		}
		return &GeneratedDocument{Filename: "document.csv", ContentType: "text/csv; charset=utf-8", Bytes: b}, nil
	case models.DocumentFormatPDF:
		data := req.Data.Get("")
		cols := v2BuildColumns(req, data)
		b, err := s.v2RenderPDF(req, cols)
		if err != nil {
			return nil, err
		}
		return &GeneratedDocument{Filename: "document.pdf", ContentType: "application/pdf", Bytes: b}, nil
	case models.DocumentFormatWord:
		data := req.Data.Get("")
		cols := v2BuildColumns(req, data)
		b, err := s.v2RenderWord(req, cols)
		if err != nil {
			return nil, err
		}
		return &GeneratedDocument{Filename: "document.docx", ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", Bytes: b}, nil
	default:
		return nil, fmt.Errorf("unsupported format")
	}
}

func v2BuildColumns(req models.DocumentRequest, data models.DynamicData) []v2Column {
	if req.Layout != nil && len(req.Layout.Columns) > 0 {
		out := make([]v2Column, 0, len(req.Layout.Columns))
		for _, c := range req.Layout.Columns {
			title := strings.TrimSpace(c.Title)
			if title == "" {
				title = c.Field
			}
			out = append(out, v2Column{Field: c.Field, Title: title, Width: c.Width, Align: c.Alignment, VAlign: c.VerticalAlignment, Format: c.Format})
		}
		return out
	}
	out := make([]v2Column, 0, len(data.Order))
	for _, k := range data.Order {
		out = append(out, v2Column{Field: k, Title: k})
	}
	return out
}

func v2ApplyDefaultFont(def models.StyleConfig, req models.DocumentRequest) models.StyleConfig {
	if req.Layout == nil || req.Layout.DefaultFont == nil {
		return def
	}
	df := req.Layout.DefaultFont
	if df.FontFamily != "" {
		def.FontFamily = df.FontFamily
	}
	if df.FontSize != 0 {
		def.FontSize = df.FontSize
	}
	if strings.TrimSpace(df.FontColor) != "" {
		def.FontColor = df.FontColor
	}
	return def
}

func v2StyleHeader(req models.DocumentRequest) models.StyleConfig {
	def := models.StyleConfig{FontFamily: models.FontCalibri, FontSize: 11, Bold: true, Alignment: models.AlignCenter, VerticalAlign: models.VAlignMiddle, Border: true, Background: "#FFFFFF", FontColor: "#000000"}
	def = v2ApplyDefaultFont(def, req)
	if req.Layout == nil || req.Layout.Header == nil {
		return def
	}
	return v2MergeStyle(def, *req.Layout.Header)
}

func v2StyleBody(req models.DocumentRequest) models.StyleConfig {
	def := models.StyleConfig{FontFamily: models.FontCalibri, FontSize: 11, Alignment: models.AlignLeft, VerticalAlign: models.VAlignMiddle, Border: true, ZebraStripe: false, ZebraColorOdd: "#F3F3F3", ZebraColorEven: "#FFFFFF", Background: "#FFFFFF", FontColor: "#000000"}
	def = v2ApplyDefaultFont(def, req)
	if req.Layout == nil || req.Layout.Body == nil {
		return def
	}
	return v2MergeStyle(def, *req.Layout.Body)
}

func v2StyleFooter(req models.DocumentRequest) models.StyleConfig {
	def := models.StyleConfig{FontFamily: models.FontCalibri, FontSize: 9, Alignment: models.AlignRight, VerticalAlign: models.VAlignMiddle, FontColor: "#000000"}
	def = v2ApplyDefaultFont(def, req)
	if req.Layout == nil || req.Layout.Footer == nil {
		return def
	}
	return v2MergeStyle(def, req.Layout.Footer.StyleConfig)
}

func v2MergeStyle(def, over models.StyleConfig) models.StyleConfig {
	out := def
	if over.FontFamily != "" {
		out.FontFamily = over.FontFamily
	}
	if over.FontSize != 0 {
		out.FontSize = over.FontSize
	}
	out.Bold = over.Bold
	out.Italic = over.Italic
	out.Underline = over.Underline
	if strings.TrimSpace(over.FontColor) != "" {
		out.FontColor = over.FontColor
	}
	if strings.TrimSpace(over.Background) != "" {
		out.Background = over.Background
	}
	if over.Alignment != "" {
		out.Alignment = over.Alignment
	}
	if over.VerticalAlign != "" {
		out.VerticalAlign = over.VerticalAlign
	}
	out.Border = over.Border
	out.ZebraStripe = over.ZebraStripe
	if strings.TrimSpace(over.ZebraColorOdd) != "" {
		out.ZebraColorOdd = over.ZebraColorOdd
	}
	if strings.TrimSpace(over.ZebraColorEven) != "" {
		out.ZebraColorEven = over.ZebraColorEven
	}
	return out
}

func v2AnyToString(v any) string {
	if v == nil {
		return ""
	}
	switch vv := v.(type) {
	case string:
		return vv
	case bool:
		if vv {
			return "true"
		}
		return "false"
	case int:
		return strconv.Itoa(vv)
	case int64:
		return fmt.Sprintf("%d", vv)
	case float64:
		if vv == math.Trunc(vv) {
			return fmt.Sprintf("%.0f", vv)
		}
		return fmt.Sprintf("%v", vv)
	case json.Number:
		return vv.String()
	default:
		b, err := json.Marshal(vv)
		if err == nil {
			return string(b)
		}
		return fmt.Sprintf("%v", vv)
	}
}

func v2NormalizeHexColor(s string) string {
	v := strings.TrimSpace(s)
	if v == "" {
		return ""
	}
	if v[0] != '#' {
		v = "#" + v
	}
	if len(v) != 7 {
		return "#000000"
	}
	return strings.ToUpper(v)
}

func (s *DocumentService) v2RenderExcel(req models.DocumentRequest) ([]byte, error) {
	var (
		f   *excelize.File
		err error
	)
	if strings.TrimSpace(req.TemplateID) != "" {
		_, path, loadErr := s.templates.Load(models.DocumentExcel, req.TemplateID)
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

	// Determine sheets.
	targetSheets := []models.SheetConfig{}
	if req.Layout != nil && len(req.Layout.Sheets) > 0 {
		targetSheets = req.Layout.Sheets
	} else {
		name := f.GetSheetName(0)
		if strings.TrimSpace(name) == "" {
			name = "Sheet1"
		}
		targetSheets = []models.SheetConfig{{Name: name}}
	}

	for i, sh := range targetSheets {
		sheetName := strings.TrimSpace(sh.Name)
		if sheetName == "" {
			sheetName = fmt.Sprintf("Sheet%d", i+1)
		}
		// Ensure sheet exists.
		idx, err := f.GetSheetIndex(sheetName)
		if err != nil || idx < 0 {
			idx, err = f.NewSheet(sheetName)
			if err != nil {
				return nil, err
			}
		}
		if i == 0 {
			f.SetActiveSheet(idx)
		}

		data := req.Data.Get(sh.DataSource)
		cols := v2BuildColumns(req, data)
		if err := v2ApplyExcelPageSetup(f, sheetName, req); err != nil {
			return nil, err
		}
		if req.Layout != nil && req.Layout.FreezeHeader {
			_ = f.SetPanes(sheetName, &excelize.Panes{
				Freeze:      true,
				Split:       false,
				XSplit:      0,
				YSplit:      1,
				TopLeftCell: "A2",
				ActivePane:  "bottomLeft",
				Selection:   []excelize.Selection{{SQRef: "A2", ActiveCell: "A2", Pane: "bottomLeft"}},
			})
		}

		if err := v2RenderExcelSheet(f, sheetName, req, cols, data); err != nil {
			return nil, err
		}
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func v2ApplyExcelPageSetup(f *excelize.File, sheet string, req models.DocumentRequest) error {
	if req.Layout == nil {
		return nil
	}
	// Orientation
	if req.Layout.PageOrientation != "" {
		ori := "portrait"
		if req.Layout.PageOrientation == models.PageLandscape {
			ori = "landscape"
		}
		if err := f.SetPageLayout(sheet, &excelize.PageLayoutOptions{Orientation: &ori}); err != nil {
			return err
		}
	}
	// Margins (mm -> inches)
	if req.Layout.PageMargin != nil {
		mmToIn := func(v float64) float64 { return v / 25.4 }
		top := mmToIn(req.Layout.PageMargin.Top)
		bottom := mmToIn(req.Layout.PageMargin.Bottom)
		left := mmToIn(req.Layout.PageMargin.Left)
		right := mmToIn(req.Layout.PageMargin.Right)
		if err := f.SetPageMargins(sheet, &excelize.PageLayoutMarginsOptions{Top: &top, Bottom: &bottom, Left: &left, Right: &right}); err != nil {
			return err
		}
	}
	return nil
}

func v2RenderExcelSheet(f *excelize.File, sheet string, req models.DocumentRequest, cols []v2Column, data models.DynamicData) error {
	hStyle := v2StyleHeader(req)
	bStyle := v2StyleBody(req)

	// Cache style IDs by key.
	styleCache := map[string]int{}
	getStyle := func(isHeader bool, fillHex string, hAlign string, vAlign string, numFmt *string) (int, error) {
		key := fmt.Sprintf("h=%v|fill=%s|ha=%s|va=%s|nf=%v", isHeader, fillHex, hAlign, vAlign, numFmt)
		if id, ok := styleCache[key]; ok {
			return id, nil
		}
		base := bStyle
		if isHeader {
			base = hStyle
		}
		fontColor := v2NormalizeHexColor(base.FontColor)
		bg := v2NormalizeHexColor(fillHex)
		st := &excelize.Style{
			Font: &excelize.Font{Bold: base.Bold, Italic: base.Italic, Underline: func() string {
				if base.Underline {
					return "single"
				}
				return ""
			}(), Size: float64(maxInt(1, base.FontSize)), Color: fontColor},
			Alignment:    &excelize.Alignment{Horizontal: hAlign, Vertical: vAlign, WrapText: true},
			Fill:         excelize.Fill{Type: "pattern", Color: []string{bg}, Pattern: 1},
			Border:       excelBorders(base.Border),
			CustomNumFmt: numFmt,
		}
		id, err := f.NewStyle(st)
		if err != nil {
			return 0, err
		}
		styleCache[key] = id
		return id, nil
	}

	// Compute widths if autosize.
	computedWidths := make([]float64, len(cols))
	if req.Layout != nil && req.Layout.AutoSizeColumns {
		for i, c := range cols {
			maxLen := utf8.RuneCountInString(c.Title)
			for _, item := range data.Items {
				l := utf8.RuneCountInString(v2AnyToString(item[c.Field]))
				if l > maxLen {
					maxLen = l
				}
			}
			w := float64(maxLen)*1.2 + 2
			if w < 8 {
				w = 8
			}
			if w > 60 {
				w = 60
			}
			computedWidths[i] = w
		}
	}

	for i, c := range cols {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		hID, err := getStyle(true, hStyle.Background, excelHAlign(hStyle.Alignment), excelVAlign(hStyle.VerticalAlign), nil)
		if err != nil {
			return err
		}
		_ = f.SetCellValue(sheet, cell, c.Title)
		_ = f.SetCellStyle(sheet, cell, cell, hID)

		width := c.Width
		if req.Layout != nil && req.Layout.AutoSizeColumns {
			width = computedWidths[i]
		}
		if width > 0 {
			colLetter, _ := excelize.ColumnNumberToName(i + 1)
			_ = f.SetColWidth(sheet, colLetter, colLetter, width)
		}
	}

	oddBg := bStyle.ZebraColorOdd
	evenBg := bStyle.ZebraColorEven
	for r, item := range data.Items {
		excelRow := r + 2
		for cIdx, c := range cols {
			cell, _ := excelize.CoordinatesToCellName(cIdx+1, excelRow)
			val := item[c.Field]
			// Set raw value for better formatting when possible.
			setVal := val
			if c.Format != "" {
				if parsed, ok := v2CoerceForExcel(c.Format, val); ok {
					setVal = parsed
				}
			}
			_ = f.SetCellValue(sheet, cell, setVal)

			fill := bStyle.Background
			if bStyle.ZebraStripe {
				if r%2 == 0 {
					fill = evenBg
				} else {
					fill = oddBg
				}
			}
			ha := excelHAlign(bStyle.Alignment)
			if c.Align != "" {
				ha = excelColAlign(c.Align)
			}
			va := excelVAlign(bStyle.VerticalAlign)
			if c.VAlign != "" {
				va = excelVAlign(c.VAlign)
			}
			nf := v2ExcelCustomNumFmt(c.Format)
			bID, err := getStyle(false, fill, ha, va, nf)
			if err != nil {
				return err
			}
			_ = f.SetCellStyle(sheet, cell, cell, bID)
		}
	}
	return nil
}

func v2ExcelCustomNumFmt(fmtEnum models.ColumnFormatEnum) *string {
	switch fmtEnum {
	case models.ColFormatCurrency:
		v := "#,##0.00"
		return &v
	case models.ColFormatNumber:
		v := "0.00"
		return &v
	case models.ColFormatPercentage:
		v := "0.00%"
		return &v
	case models.ColFormatDate:
		v := "yyyy-mm-dd"
		return &v
	case models.ColFormatDateTime:
		v := "yyyy-mm-dd hh:mm:ss"
		return &v
	default:
		return nil
	}
}

func v2CoerceForExcel(fmtEnum models.ColumnFormatEnum, v any) (any, bool) {
	if v == nil {
		return nil, false
	}
	switch fmtEnum {
	case models.ColFormatCurrency, models.ColFormatNumber, models.ColFormatPercentage:
		switch vv := v.(type) {
		case int:
			return float64(vv), true
		case int64:
			return float64(vv), true
		case float64:
			return vv, true
		case json.Number:
			if f, err := vv.Float64(); err == nil {
				return f, true
			}
		case string:
			if f, err := strconv.ParseFloat(strings.ReplaceAll(strings.TrimSpace(vv), ",", "."), 64); err == nil {
				return f, true
			}
		}
	case models.ColFormatDate, models.ColFormatDateTime:
		switch vv := v.(type) {
		case string:
			s := strings.TrimSpace(vv)
			if s == "" {
				return nil, false
			}
			// Excel date serial number: days since 1899-12-30.
			var t time.Time
			var err error
			layouts := []string{time.RFC3339, "2006-01-02", "2006-01-02 15:04:05", "02/01/2006"}
			for _, layout := range layouts {
				t, err = time.Parse(layout, s)
				if err == nil {
					break
				}
			}
			if err != nil {
				return nil, false
			}
			return v2ExcelTimeSerial(t), true
		}
	}
	return nil, false
}

func v2ExcelTimeSerial(t time.Time) float64 {
	// Excel date serial (1900 date system): days since 1899-12-30.
	// This matches common Excel conventions including the 1900 leap-year bug offset.
	base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	if t.Location() != time.UTC {
		t = t.UTC()
	}
	d := t.Sub(base).Hours() / 24
	return d
}

func (s *DocumentService) v2RenderCSV(req models.DocumentRequest, cols []v2Column) ([]byte, error) {
	var b bytes.Buffer
	w := csv.NewWriter(&b)

	head := make([]string, 0, len(cols))
	for _, c := range cols {
		head = append(head, c.Title)
	}
	if err := w.Write(head); err != nil {
		return nil, err
	}

	data := req.Data.Get("")
	for _, item := range data.Items {
		row := make([]string, 0, len(cols))
		for _, c := range cols {
			row = append(row, v2AnyToString(item[c.Field]))
		}
		if err := w.Write(row); err != nil {
			return nil, err
		}
	}

	w.Flush()
	if err := w.Error(); err != nil {
		return nil, err
	}
	return b.Bytes(), nil
}

func (s *DocumentService) v2RenderPDF(req models.DocumentRequest, cols []v2Column) ([]byte, error) {
	if strings.TrimSpace(req.TemplateID) != "" {
		b, _, loadErr := s.templates.Load(models.DocumentPDF, req.TemplateID)
		if loadErr != nil {
			return nil, loadErr
		}
		return b, nil
	}

	hStyle := v2StyleHeader(req)
	bStyle := v2StyleBody(req)
	fStyle := v2StyleFooter(req)
	data := req.Data.Get("")

	orient := "P"
	if req.Layout != nil && req.Layout.PageOrientation == models.PageLandscape {
		orient = "L"
	}
	pdf := gofpdf.New(orient, "mm", "A4", "")
	// Margins in mm
	if req.Layout != nil && req.Layout.PageMargin != nil {
		pdf.SetMargins(req.Layout.PageMargin.Left, req.Layout.PageMargin.Top, req.Layout.PageMargin.Right)
		pdf.SetAutoPageBreak(true, req.Layout.PageMargin.Bottom)
	} else {
		pdf.SetAutoPageBreak(true, 12)
	}
	pdf.AliasNbPages("")
	if req.Layout != nil && req.Layout.Footer != nil && req.Layout.Footer.PageNumber != nil && req.Layout.Footer.PageNumber.Enabled {
		align := req.Layout.Footer.Alignment
		if align == "" {
			align = models.AlignCenter
		}
		fmtEnum := req.Layout.Footer.PageNumber.Format
		if fmtEnum == "" {
			fmtEnum = models.PageNumArabic
		}
		show := true
		if req.Layout.Footer.Show != nil {
			show = *req.Layout.Footer.Show
		}
		if show {
			pdf.SetFooterFunc(func() {
				pdf.SetY(-10)
				v2SetPDFFont(pdf, fStyle)
				pdf.CellFormat(0, 10, v2FormatPageNumberPDF(pdf.PageNo(), fmtEnum), "", 0, v2PDFAlign(align), false, 0, "")
			})
		}
	}
	pdf.AddPage()

	if req.Layout != nil && req.Layout.HeaderImage != nil && strings.TrimSpace(req.Layout.HeaderImage.Data) != "" {
		if err := v2AddPDFHeaderImage(pdf, *req.Layout.HeaderImage); err != nil {
			return nil, err
		}
	}

	pageW, _ := pdf.GetPageSize()
	lm, _, rm, _ := pdf.GetMargins()
	usable := pageW - lm - rm
	widths := v2ComputePDFWidths(cols, usable)

	v2SetPDFFont(pdf, hStyle)
	v2SetPDFFill(pdf, hStyle.Background)
	v2SetPDFText(pdf, hStyle.FontColor)
	for i, c := range cols {
		pdf.CellFormat(widths[i], 8, v2PDFText(c.Title), v2PDFBorder(hStyle.Border), 0, v2PDFColAlign(c.Align), true, 0, "")
	}
	pdf.Ln(-1)

	rowsPerPage := 0
	if req.Layout != nil && req.Layout.PageBreak != nil && req.Layout.PageBreak.Enabled {
		rowsPerPage = req.Layout.PageBreak.RowsPerPage
	}
	printed := 0
	for r, item := range data.Items {
		fillColor := bStyle.Background
		if bStyle.ZebraStripe {
			if r%2 == 0 {
				fillColor = bStyle.ZebraColorEven
			} else {
				fillColor = bStyle.ZebraColorOdd
			}
		}
		v2SetPDFFont(pdf, bStyle)
		v2SetPDFFill(pdf, fillColor)
		v2SetPDFText(pdf, bStyle.FontColor)
		for i, c := range cols {
			pdf.CellFormat(widths[i], 7, v2PDFText(v2AnyToString(item[c.Field])), v2PDFBorder(bStyle.Border), 0, v2PDFColAlign(c.Align), true, 0, "")
		}
		pdf.Ln(-1)
		printed++
		if rowsPerPage > 0 && printed%rowsPerPage == 0 && r < len(data.Items)-1 {
			pdf.AddPage()
			if req.Layout != nil && req.Layout.HeaderImage != nil && strings.TrimSpace(req.Layout.HeaderImage.Data) != "" {
				if err := v2AddPDFHeaderImage(pdf, *req.Layout.HeaderImage); err != nil {
					return nil, err
				}
			}
			v2SetPDFFont(pdf, hStyle)
			v2SetPDFFill(pdf, hStyle.Background)
			v2SetPDFText(pdf, hStyle.FontColor)
			for i, c := range cols {
				pdf.CellFormat(widths[i], 8, v2PDFText(c.Title), v2PDFBorder(hStyle.Border), 0, v2PDFColAlign(c.Align), true, 0, "")
			}
			pdf.Ln(-1)
		}
	}

	var out bytes.Buffer
	if err := pdf.Output(&out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func (s *DocumentService) v2RenderWord(req models.DocumentRequest, cols []v2Column) ([]byte, error) {
	// If a docx template exists, treat it as a raw-docx template with {{table_xml}}.
	// (We keep simple behavior; full templating can be extended later.)
	// If no template, generate a minimal docx in-memory.
	includeFooter := false
	footerAlign := models.AlignCenter
	footerFmt := models.PageNumArabic
	if req.Layout != nil && req.Layout.Footer != nil && req.Layout.Footer.PageNumber != nil && req.Layout.Footer.PageNumber.Enabled {
		show := true
		if req.Layout.Footer.Show != nil {
			show = *req.Layout.Footer.Show
		}
		includeFooter = show
		if req.Layout.Footer.Alignment != "" {
			footerAlign = req.Layout.Footer.Alignment
		}
		if req.Layout.Footer.PageNumber.Format != "" {
			footerFmt = req.Layout.Footer.PageNumber.Format
		}
	}
	includeImage := req.Layout != nil && req.Layout.HeaderImage != nil && strings.TrimSpace(req.Layout.HeaderImage.Data) != ""

	// Template-based path (optional)
	if strings.TrimSpace(req.TemplateID) != "" {
		b, _, loadErr := s.templates.Load(models.DocumentWord, req.TemplateID)
		if loadErr == nil && len(b) > 0 {
			// We cannot safely edit relationships in arbitrary templates here.
			// Fall back to our generated docx to ensure table/footer/image work.
		}
	}

	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)

	writeZipString(zw, "[Content_Types].xml", v2WordContentTypesXML(includeFooter, includeImage, req))
	writeZipString(zw, "_rels/.rels", v2WordRootRelsXML())
	writeZipString(zw, "word/styles.xml", v2WordStylesXML())
	writeZipString(zw, "word/document.xml", v2WordDocumentXML(req, cols, includeFooter, includeImage))
	writeZipString(zw, "word/_rels/document.xml.rels", v2WordDocumentRelsXML(includeFooter, includeImage, req))
	if includeFooter {
		writeZipString(zw, "word/footer1.xml", v2WordFooterXML(v2StyleFooter(req), footerAlign, footerFmt))
	}
	if includeImage {
		imgBytes, ext, err := v2DecodeImageBytes(req.Layout.HeaderImage.Data)
		if err != nil {
			_ = zw.Close()
			return nil, err
		}
		w, _ := zw.Create("word/media/image1" + ext)
		_, _ = w.Write(imgBytes)
	}

	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// ---- PDF helpers ----

func v2FormatPageNumberPDF(pageNo int, fmtEnum models.PageNumberFormatEnum) string {
	if pageNo <= 0 {
		pageNo = 1
	}
	switch fmtEnum {
	case models.PageNumTextPageNum:
		return fmt.Sprintf("Page %d", pageNo)
	case models.PageNumRoman:
		return strings.ToLower(v2ToRoman(pageNo))
	case models.PageNumRomanUpper:
		return v2ToRoman(pageNo)
	case models.PageNumArabic, "":
		fallthrough
	default:
		return fmt.Sprintf("%d", pageNo)
	}
}

func v2ToRoman(n int) string {
	if n <= 0 {
		return "I"
	}
	if n > 3999 {
		return fmt.Sprintf("%d", n)
	}
	vals := []struct {
		v int
		s string
	}{
		{1000, "M"},
		{900, "CM"},
		{500, "D"},
		{400, "CD"},
		{100, "C"},
		{90, "XC"},
		{50, "L"},
		{40, "XL"},
		{10, "X"},
		{9, "IX"},
		{5, "V"},
		{4, "IV"},
		{1, "I"},
	}
	var b strings.Builder
	for _, it := range vals {
		for n >= it.v {
			b.WriteString(it.s)
			n -= it.v
		}
	}
	return b.String()
}

func v2ComputePDFWidths(cols []v2Column, usable float64) []float64 {
	widths := make([]float64, len(cols))
	var sum float64
	for i, c := range cols {
		w := c.Width
		if w <= 0 {
			w = 1
		}
		widths[i] = w
		sum += w
	}
	if sum > 0 {
		scale := usable / sum
		for i := range widths {
			widths[i] *= scale
		}
	}
	return widths
}

func v2SetPDFFont(pdf *gofpdf.Fpdf, s models.StyleConfig) {
	// gofpdf supports only core fonts (Courier, Helvetica, Times, Symbol, ZapfDingbats)
	// and a few aliases (e.g., Arial -> Helvetica). Map our requested fonts accordingly.
	var family string
	switch s.FontFamily {
	case models.FontCalibri:
		family = "Helvetica"
	case models.FontArial:
		// Arial is commonly treated as an alias of Helvetica in FPDF implementations.
		family = "Helvetica"
	case models.FontTimesNewRoman:
		family = "Times"
	case models.FontHelvetica:
		family = "Helvetica"
	default:
		family = string(s.FontFamily)
		if strings.TrimSpace(family) == "" {
			family = "Helvetica"
		}
		// Defensive fallback for unknown/unsupported names.
		switch strings.ToLower(strings.TrimSpace(family)) {
		case "calibri", "arial":
			family = "Helvetica"
		case "timesnewroman", "times new roman", "times":
			family = "Times"
		}
	}
	style := ""
	if s.Bold {
		style += "B"
	}
	if s.Italic {
		style += "I"
	}
	if s.Underline {
		style += "U"
	}
	size := float64(s.FontSize)
	if size <= 0 {
		size = 11
	}
	pdf.SetFont(family, style, size)
}

func v2HexRGB(hex string) (int, int, int) {
	v := v2NormalizeHexColor(hex)
	v = strings.TrimPrefix(v, "#")
	if len(v) != 6 {
		return 0, 0, 0
	}
	r, _ := strconv.ParseInt(v[0:2], 16, 32)
	g, _ := strconv.ParseInt(v[2:4], 16, 32)
	b, _ := strconv.ParseInt(v[4:6], 16, 32)
	return int(r), int(g), int(b)
}

func v2SetPDFFill(pdf *gofpdf.Fpdf, hex string) {
	r, g, b := v2HexRGB(hex)
	pdf.SetFillColor(r, g, b)
}

func v2SetPDFText(pdf *gofpdf.Fpdf, hex string) {
	r, g, b := v2HexRGB(hex)
	pdf.SetTextColor(r, g, b)
}

func v2PDFAlign(a models.AlignmentEnum) string {
	switch a {
	case models.AlignCenter:
		return "C"
	case models.AlignRight:
		return "R"
	default:
		return "L"
	}
}

func v2PDFColAlign(a models.ColumnAlignmentEnum) string {
	switch a {
	case models.ColAlignCenter:
		return "C"
	case models.ColAlignRight:
		return "R"
	default:
		return "L"
	}
}

func v2PDFBorder(enabled bool) string {
	if enabled {
		return "1"
	}
	return ""
}

func v2AddPDFHeaderImage(pdf *gofpdf.Fpdf, cfg models.HeaderImageConfig) error {
	imgBytes, _, err := v2DecodeImageBytes(cfg.Data)
	if err != nil {
		return err
	}
	ic, format, err := image.DecodeConfig(bytes.NewReader(imgBytes))
	if err != nil {
		return err
	}
	opt := gofpdf.ImageOptions{ImageType: strings.ToUpper(format), ReadDpi: true}
	name := "headerImage"
	pdf.RegisterImageOptionsReader(name, opt, bytes.NewReader(imgBytes))

	height := cfg.Height
	if height <= 0 {
		height = 12
	}
	pageW, _ := pdf.GetPageSize()
	lm, tm, rm, _ := pdf.GetMargins()
	usableW := pageW - lm - rm

	// Desired bounding box.
	boxW := usableW * 0.4
	if cfg.FillHeaderWidth {
		boxW = usableW
	}
	if boxW <= 0 {
		boxW = usableW
	}
	boxH := height

	keepAR := true
	if cfg.KeepAspectRatio != nil {
		keepAR = *cfg.KeepAspectRatio
	}
	fit := cfg.FitMode
	if fit == "" && cfg.Stretch {
		fit = models.ImageFitStretch
	}
	if fit == models.ImageFitStretch && keepAR {
		fit = models.ImageFitContain
	}

	ratio := 1.0
	if ic.Height > 0 {
		ratio = float64(ic.Width) / float64(ic.Height)
	}

	// Determine box x based on alignment.
	xBox := lm
	// Prefer explicit horizontalAlignment if provided; fall back to legacy position.
	switch cfg.HorizontalAlignment {
	case models.ColAlignRight:
		xBox = lm + (usableW - boxW)
	case models.ColAlignCenter:
		xBox = lm + (usableW-boxW)/2
	case models.ColAlignLeft, "":
		// handled below
	}
	if cfg.HorizontalAlignment == "" {
		switch cfg.Position {
		case models.ImageTopCenter:
			xBox = lm + (usableW-boxW)/2
		case models.ImageTopRight:
			xBox = lm + (usableW - boxW)
		default:
			xBox = lm
		}
	}

	// Compute draw sizes.
	drawW, drawH := boxW, boxH
	if keepAR {
		switch fit {
		case models.ImageFitCover:
			// Scale to cover the box (may overflow and be clipped).
			drawW = boxW
			drawH = drawW / ratio
			if drawH < boxH {
				drawH = boxH
				drawW = drawH * ratio
			}
		default:
			// Contain/center: scale to fit within box.
			drawW = boxW
			drawH = drawW / ratio
			if drawH > boxH {
				drawH = boxH
				drawW = drawH * ratio
			}
		}
	}

	// Image placement.
	imgX, imgY := xBox, tm
	if fit == models.ImageFitCover {
		// Center inside the box; clip to box.
		imgX = xBox + (boxW-drawW)/2
		imgY = tm + (boxH-drawH)/2
		pdf.ClipRect(xBox, tm, boxW, boxH, false)
		pdf.ImageOptions(name, imgX, imgY, drawW, drawH, false, opt, 0, "")
		pdf.ClipEnd()
	} else {
		pdf.ImageOptions(name, imgX, imgY, drawW, drawH, false, opt, 0, "")
	}

	pdf.Ln(boxH + 4)
	return nil
}

func v2PDFText(s string) string {
	out, err := charmap.ISO8859_1.NewEncoder().String(s)
	if err != nil {
		return s
	}
	return out
}

// ---- Excel helpers ----

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func excelHAlign(a models.AlignmentEnum) string {
	switch a {
	case models.AlignCenter:
		return "center"
	case models.AlignRight:
		return "right"
	default:
		return "left"
	}
}

func excelVAlign(a models.VerticalAlignmentEnum) string {
	switch a {
	case models.VAlignTop:
		return "top"
	case models.VAlignBottom:
		return "bottom"
	default:
		return "center"
	}
}

func excelColAlign(a models.ColumnAlignmentEnum) string {
	switch a {
	case models.ColAlignCenter:
		return "center"
	case models.ColAlignRight:
		return "right"
	default:
		return "left"
	}
}

func excelBorders(enabled bool) []excelize.Border {
	if !enabled {
		return nil
	}
	return []excelize.Border{{Type: "left", Color: "C0C0C0", Style: 1}, {Type: "top", Color: "C0C0C0", Style: 1}, {Type: "bottom", Color: "C0C0C0", Style: 1}, {Type: "right", Color: "C0C0C0", Style: 1}}
}

// ---- Word (docx zip) helpers ----

func writeZipString(zw *zip.Writer, name string, content string) {
	w, _ := zw.Create(name)
	_, _ = w.Write([]byte(content))
}

func v2WordContentTypesXML(includeFooter, includeImage bool, req models.DocumentRequest) string {
	imgDefault := ""
	if includeImage {
		_, ext, _ := v2DecodeImageBytes(req.Layout.HeaderImage.Data)
		ct := "image/png"
		if ext == ".jpg" || ext == ".jpeg" {
			ct = "image/jpeg"
		}
		imgDefault = fmt.Sprintf("\n  <Default Extension=\"%s\" ContentType=\"%s\"/>", strings.TrimPrefix(ext, "."), ct)
	}
	footerOverride := ""
	if includeFooter {
		footerOverride = "\n  <Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>"
	}
	return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>` + imgDefault + `
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` + footerOverride + `
</Types>`
}

func v2WordRootRelsXML() string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
}

func v2WordDocumentRelsXML(includeFooter, includeImage bool, req models.DocumentRequest) string {
	parts := []string{
		`  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
	}
	if includeFooter {
		parts = append(parts, `  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
	}
	if includeImage {
		_, ext, _ := v2DecodeImageBytes(req.Layout.HeaderImage.Data)
		parts = append(parts, fmt.Sprintf(`  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1%s"/>`, ext))
	}
	return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" + strings.Join(parts, "\n") + "\n</Relationships>"
}

func v2WordStylesXML() string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
		<w:name w:val="Normal"/>
	</w:style>
</w:styles>`
}

func v2WordFooterXML(style models.StyleConfig, align models.AlignmentEnum, fmtEnum models.PageNumberFormatEnum) string {
	_ = style
	jc := "center"
	switch align {
	case models.AlignLeft:
		jc = "left"
	case models.AlignRight:
		jc = "right"
	default:
		jc = "center"
	}
	instr := " PAGE "
	switch fmtEnum {
	case models.PageNumRoman:
		instr = " PAGE \\* roman "
	case models.PageNumRomanUpper:
		instr = " PAGE \\* ROMAN "
	case models.PageNumArabic, "":
		instr = " PAGE "
	}
	// TextPageNumber adds a static "Page " prefix plus an Arabic PAGE field.
	textPrefix := ""
	if fmtEnum == models.PageNumTextPageNum {
		textPrefix = `<w:r><w:t>Page </w:t></w:r>`
		instr = " PAGE "
	}
	return `<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<w:p>
		<w:pPr><w:jc w:val="` + jc + `"/></w:pPr>
		` + textPrefix + `
		<w:r><w:fldChar w:fldCharType="begin"/></w:r>
		<w:r><w:instrText xml:space="preserve">` + instr + `</w:instrText></w:r>
		<w:r><w:fldChar w:fldCharType="end"/></w:r>
	</w:p>
</w:ftr>`
}

func v2WordDocumentXML(req models.DocumentRequest, cols []v2Column, includeFooter, includeImage bool) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
`)
	if includeImage {
		sb.WriteString(v2WordImageParagraphXML(req))
	}
	sb.WriteString(v2WordTablesXML(req, cols))
	sb.WriteString(v2WordSectPrXML(req, includeFooter))
	sb.WriteString(`
	  </w:body>
</w:document>`)
	return sb.String()
}

func v2WordSectPrXML(req models.DocumentRequest, includeFooter bool) string {
	var sb strings.Builder
	sb.WriteString(`<w:sectPr>`)
	if includeFooter {
		sb.WriteString(`<w:footerReference w:type="default" r:id="rIdFooter1"/>`)
	}
	// Page size/orientation (A4 defaults).
	if req.Layout != nil && req.Layout.PageOrientation != "" {
		wTw, hTw, orient := v2WordPageSizeTwips(req.Layout.PageOrientation)
		sb.WriteString(fmt.Sprintf(`<w:pgSz w:w="%d" w:h="%d" w:orient="%s"/>`, wTw, hTw, orient))
	}
	// Margins (mm -> twips)
	if req.Layout != nil && req.Layout.PageMargin != nil {
		m := req.Layout.PageMargin
		sb.WriteString(fmt.Sprintf(`<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="0" w:footer="0" w:gutter="0"/>`, v2MMToTwips(m.Top), v2MMToTwips(m.Right), v2MMToTwips(m.Bottom), v2MMToTwips(m.Left)))
	}
	sb.WriteString(`</w:sectPr>`)
	return sb.String()
}

func v2MMToTwips(mm float64) int {
	// 1 inch = 25.4mm, 1 inch = 1440 twips
	if mm < 0 {
		mm = 0
	}
	return int(math.Round(mm * 1440.0 / 25.4))
}

func v2WordPageSizeTwips(orientation models.PageOrientationEnum) (wTwips int, hTwips int, orientAttr string) {
	// A4: 210mm x 297mm => 11906 x 16838 twips
	wTwips, hTwips = 11906, 16838
	orientAttr = "portrait"
	if orientation == models.PageLandscape {
		wTwips, hTwips = hTwips, wTwips
		orientAttr = "landscape"
	}
	return
}

func v2WordImageParagraphXML(req models.DocumentRequest) string {
	// Compute requested box size in EMU.
	boxWemu := int64(3048000) // ~3.33in default
	boxHemu := int64(457200)  // ~0.5in default
	jc := "center"

	if req.Layout != nil && req.Layout.HeaderImage != nil {
		cfg := req.Layout.HeaderImage
		if cfg.Position == models.ImageTopLeft {
			jc = "left"
		} else if cfg.Position == models.ImageTopRight {
			jc = "right"
		}
		if cfg.HorizontalAlignment == models.ColAlignLeft {
			jc = "left"
		} else if cfg.HorizontalAlignment == models.ColAlignRight {
			jc = "right"
		} else if cfg.HorizontalAlignment == models.ColAlignCenter {
			jc = "center"
		}

		heightMM := cfg.Height
		if heightMM <= 0 {
			heightMM = 12
		}
		boxHemu = v2MMToEMU(heightMM)

		// Width target.
		pageWtw, _, _ := v2WordPageSizeTwips(req.Layout.PageOrientation)
		pageWemu := v2TwipsToEMU(pageWtw)
		if cfg.FillHeaderWidth {
			boxWemu = pageWemu
		}

		keepAR := true
		if cfg.KeepAspectRatio != nil {
			keepAR = *cfg.KeepAspectRatio
		}
		fit := cfg.FitMode
		if fit == "" && cfg.Stretch {
			fit = models.ImageFitStretch
		}
		if fit == models.ImageFitStretch && keepAR {
			fit = models.ImageFitContain
		}
		if fit == models.ImageFitContain || fit == models.ImageFitCenter || fit == models.ImageFitCover {
			// Approximate contain/center/cover by contain (no cropping).
			imgBytes, _, err := v2DecodeImageBytes(cfg.Data)
			if err == nil {
				if ic, _, err2 := image.DecodeConfig(bytes.NewReader(imgBytes)); err2 == nil {
					ratio := float64(ic.Width) / float64(maxInt(1, ic.Height))
					w := int64(float64(boxHemu) * ratio)
					if w > boxWemu {
						w = boxWemu
						boxHemu = int64(float64(w) / ratio)
					}
					boxWemu = w
				}
			}
		}
	}

	return fmt.Sprintf(`<w:p>
	<w:pPr><w:jc w:val="%s"/></w:pPr>
	<w:r>
		<w:drawing>
			<wp:inline distT="0" distB="0" distL="0" distR="0">
				<wp:extent cx="%d" cy="%d"/>
				<wp:docPr id="1" name="HeaderImage"/>
				<a:graphic>
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic>
							<pic:nvPicPr>
								<pic:cNvPr id="0" name="image1"/>
								<pic:cNvPicPr/>
							</pic:nvPicPr>
							<pic:blipFill>
								<a:blip r:embed="rIdImage1"/>
								<a:stretch><a:fillRect/></a:stretch>
							</pic:blipFill>
							<pic:spPr>
								<a:xfrm>
									<a:off x="0" y="0"/>
									<a:ext cx="%d" cy="%d"/>
								</a:xfrm>
								<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
							</pic:spPr>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:inline>
		</w:drawing>
	</w:r>
</w:p>`, jc, boxWemu, boxHemu, boxWemu, boxHemu)
}

func v2MMToEMU(mm float64) int64 {
	// 1 inch = 25.4mm, 1 inch = 914400 EMU
	if mm < 0 {
		mm = 0
	}
	return int64(math.Round(mm * 914400.0 / 25.4))
}

func v2TwipsToEMU(twips int) int64 {
	// 1 inch = 1440 twips, 1 inch = 914400 EMU
	return int64(math.Round(float64(twips) * 914400.0 / 1440.0))
}

func v2WordTablesXML(req models.DocumentRequest, cols []v2Column) string {
	data := req.Data.Get("")
	rowsPerPage := 0
	if req.Layout != nil && req.Layout.PageBreak != nil && req.Layout.PageBreak.Enabled {
		rowsPerPage = req.Layout.PageBreak.RowsPerPage
	}
	if rowsPerPage <= 0 {
		return v2WordTableXML(req, cols, data.Items)
	}
	var sb strings.Builder
	for i := 0; i < len(data.Items); i += rowsPerPage {
		end := i + rowsPerPage
		if end > len(data.Items) {
			end = len(data.Items)
		}
		if i > 0 {
			sb.WriteString(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)
		}
		sb.WriteString(v2WordTableXML(req, cols, data.Items[i:end]))
	}
	return sb.String()
}

func v2WordTableXML(req models.DocumentRequest, cols []v2Column, items []map[string]any) string {
	var sb strings.Builder
	useBounds := true
	if req.Layout != nil && req.Layout.UsePageContentBounds != nil {
		useBounds = *req.Layout.UsePageContentBounds
	}
	tblW := `<w:tblW w:w="0" w:type="auto"/>`
	if !useBounds {
		pageWtw, _, _ := v2WordPageSizeTwips(models.PagePortrait)
		if req.Layout != nil && req.Layout.PageOrientation != "" {
			pageWtw, _, _ = v2WordPageSizeTwips(req.Layout.PageOrientation)
		}
		tblW = fmt.Sprintf(`<w:tblW w:type="dxa" w:w="%d"/>`, pageWtw)
	}
	sb.WriteString(`<w:tbl>
	  <w:tblPr>` + tblW + `</w:tblPr>
	  <w:tblGrid>`)
	for range cols {
		sb.WriteString(`<w:gridCol w:w="2400"/>`)
	}
	sb.WriteString(`</w:tblGrid>`)
	// Header row
	sb.WriteString(`<w:tr>`)
	for _, c := range cols {
		sb.WriteString(`<w:tc><w:tcPr>` + v2WordVAlignXML(v2StyleHeader(req).VerticalAlign) + `</w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>`)
		sb.WriteString(v2EscapeXML(c.Title))
		sb.WriteString(`</w:t></w:r></w:p></w:tc>`)
	}
	sb.WriteString(`</w:tr>`)
	bodyVA := v2StyleBody(req).VerticalAlign
	for _, item := range items {
		sb.WriteString(`<w:tr>`)
		for _, c := range cols {
			va := bodyVA
			if c.VAlign != "" {
				va = c.VAlign
			}
			sb.WriteString(`<w:tc><w:tcPr>` + v2WordVAlignXML(va) + `</w:tcPr><w:p><w:r><w:t xml:space="preserve">`)
			sb.WriteString(v2EscapeXML(v2AnyToString(item[c.Field])))
			sb.WriteString(`</w:t></w:r></w:p></w:tc>`)
		}
		sb.WriteString(`</w:tr>`)
	}
	sb.WriteString(`</w:tbl>`)
	return sb.String()
}

func v2WordVAlignXML(va models.VerticalAlignmentEnum) string {
	switch va {
	case models.VAlignTop:
		return `<w:vAlign w:val="top"/>`
	case models.VAlignBottom:
		return `<w:vAlign w:val="bottom"/>`
	default:
		return `<w:vAlign w:val="center"/>`
	}
}

func v2EscapeXML(s string) string {
	r := strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		"\"", "&quot;",
		"'", "&apos;",
	)
	return r.Replace(s)
}

func v2DecodeImageBytes(data string) ([]byte, string, error) {
	v := strings.TrimSpace(data)
	if v == "" {
		return nil, "", fmt.Errorf("headerImage.data is empty")
	}
	if strings.HasPrefix(v, "data:") {
		parts := strings.SplitN(v, ",", 2)
		if len(parts) != 2 {
			return nil, "", fmt.Errorf("invalid data URI")
		}
		meta := parts[0]
		payload := parts[1]
		b, err := base64.StdEncoding.DecodeString(payload)
		if err != nil {
			return nil, "", err
		}
		ext := ".png"
		if strings.Contains(meta, "image/jpeg") {
			ext = ".jpg"
		}
		return b, ext, nil
	}
	b, err := base64.StdEncoding.DecodeString(v)
	if err != nil {
		return nil, "", err
	}
	_, format, err := image.DecodeConfig(bytes.NewReader(b))
	if err != nil {
		return nil, "", err
	}
	switch strings.ToLower(format) {
	case "png":
		return b, ".png", nil
	case "jpeg":
		return b, ".jpg", nil
	default:
		return nil, "", fmt.Errorf("unsupported image format: %s", format)
	}
}
