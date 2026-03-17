package service

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"image"
	"image/color"
	"image/draw"
	_ "image/gif"
	_ "image/jpeg"
	"image/png"
	_ "image/png"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"Data2Doc/internal/models"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/charmap"
)

type DocumentService struct{}

func NewDocumentService() *DocumentService { return &DocumentService{} }

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

	// Excel calculation features
	Formula      string
	SheetFormula string
	Aggregate    string
	PercentageOf string

	// Excel-only features
	CellType              models.ExcelCellTypeEnum
	Options               []string
	ValidationRange       string
	Lookup                *models.ExcelLookupConfig
	ConditionalFormatting []models.ExcelConditionalFormattingRule
	BackgroundColor       string
	TextColor             string
	HeaderColor           string
	Hidden                bool
	Locked                bool
}

type pendingExcelLookup struct {
	Sheet   string
	Cell    string
	KeyCell string
	Config  models.ExcelLookupConfig
}

func v2ApplyPendingLookups(f *excelize.File, pending []pendingExcelLookup, reg *excelSheetRegistry) error {
	if f == nil || len(pending) == 0 {
		return nil
	}
	for i := range pending {
		p := pending[i]
		fx, err := v2BuildLookupFormula(p, reg)
		if err != nil {
			return err
		}
		if err := f.SetCellFormula(p.Sheet, p.Cell, fx); err != nil {
			return err
		}
	}
	return nil
}

func v2BuildLookupFormula(p pendingExcelLookup, reg *excelSheetRegistry) (string, error) {
	engine := p.Config.Engine
	if strings.TrimSpace(string(engine)) == "" {
		engine = models.ExcelLookupEngineVLookup
	}
	matchMode := p.Config.MatchMode
	if strings.TrimSpace(string(matchMode)) == "" {
		matchMode = models.ExcelLookupMatchExact
	}
	lookupSheet := strings.TrimSpace(p.Config.Sheet)
	if lookupSheet == "" {
		return "", fmt.Errorf("lookup.sheet is required")
	}
	meta, ok := reg.Get(lookupSheet)
	if !ok {
		return "", fmt.Errorf("unknown lookup sheet '%s'", lookupSheet)
	}
	lookupField := strings.ToLower(strings.TrimSpace(p.Config.LookupField))
	returnField := strings.ToLower(strings.TrimSpace(p.Config.ReturnField))
	lookupCol, ok := meta.FieldToCol[lookupField]
	if !ok {
		return "", fmt.Errorf("lookupField '%s' not found in sheet '%s'", p.Config.LookupField, meta.Name)
	}
	returnCol, ok := meta.FieldToCol[returnField]
	if !ok {
		return "", fmt.Errorf("returnField '%s' not found in sheet '%s'", p.Config.ReturnField, meta.Name)
	}
	lookupColNum, err := excelize.ColumnNameToNumber(lookupCol)
	if err != nil {
		return "", err
	}
	returnColNum, err := excelize.ColumnNameToNumber(returnCol)
	if err != nil {
		return "", err
	}

	start := maxInt(2, meta.DataStartRow)
	end := maxInt(start, meta.DataEndRow)
	qSheet := excelQuoteSheetName(meta.Name)

	switch engine {
	case models.ExcelLookupEngineXLookup:
		lookupArr := fmt.Sprintf("%s!$%s$%d:$%s$%d", qSheet, lookupCol, start, lookupCol, end)
		returnArr := fmt.Sprintf("%s!$%s$%d:$%s$%d", qSheet, returnCol, start, returnCol, end)
		mm := "0"
		switch matchMode {
		case models.ExcelLookupMatchExact:
			mm = "0"
		case models.ExcelLookupMatchLessEq, models.ExcelLookupMatchApprox:
			mm = "-1"
		case models.ExcelLookupMatchGreater:
			mm = "1"
		default:
			mm = "0"
		}
		return fmt.Sprintf("XLOOKUP(%s,%s,%s,\"\",%s)", p.KeyCell, lookupArr, returnArr, mm), nil
	case models.ExcelLookupEngineVLookup:
		if returnColNum < lookupColNum {
			return "", fmt.Errorf("lookup: returnField '%s' is left of lookupField '%s' on sheet '%s'; use engine 'xlookup' or reorder the lookup sheet columns", p.Config.ReturnField, p.Config.LookupField, meta.Name)
		}
		rangeEndColNum := maxInt(lookupColNum, returnColNum)
		rangeEndCol, _ := excelize.ColumnNumberToName(rangeEndColNum)
		rangeRef := fmt.Sprintf("%s!$%s$%d:$%s$%d", qSheet, lookupCol, start, rangeEndCol, end)
		colIndex := returnColNum - lookupColNum + 1
		approx := "FALSE"
		if matchMode != models.ExcelLookupMatchExact {
			approx = "TRUE"
		}
		return fmt.Sprintf("VLOOKUP(%s,%s,%d,%s)", p.KeyCell, rangeRef, colIndex, approx), nil
	default:
		return "", fmt.Errorf("lookup.engine is invalid")
	}
}

func (s *DocumentService) Generate(req models.DocumentRequest) (*GeneratedDocument, error) {
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
			out = append(out, v2Column{
				Field:                 c.Field,
				Title:                 title,
				Width:                 c.Width,
				Align:                 c.Alignment,
				VAlign:                c.VerticalAlignment,
				Format:                c.Format,
				Formula:               c.Formula,
				SheetFormula:          c.SheetFormula,
				Aggregate:             c.Aggregate,
				PercentageOf:          c.PercentageOf,
				CellType:              c.CellType,
				Options:               c.Options,
				ValidationRange:       c.ValidationRange,
				Lookup:                c.Lookup,
				ConditionalFormatting: c.ConditionalFormatting,
				BackgroundColor:       c.BackgroundColor,
				TextColor:             c.TextColor,
				HeaderColor:           c.HeaderColor,
				Hidden:                c.Hidden,
				Locked:                c.Locked,
			})
		}
		return out
	}
	out := make([]v2Column, 0, len(data.Order))
	for _, k := range data.Order {
		out = append(out, v2Column{Field: k, Title: k})
	}
	return out
}

func v2BuildExcelColumns(req models.DocumentRequest, data models.DynamicData) []v2Column {
	if req.Layout == nil || len(req.Layout.Columns) == 0 {
		return v2BuildColumns(req, data)
	}
	// For multi-sheet Excel, it is common that each sheet points to a different dataset.
	// If the global layout.columns isn't applicable to the dataset schema for this sheet,
	// infer columns from the dataset to avoid empty sheets with unrelated headers.
	if !v2ShouldUseConfiguredColumns(req.Layout.Columns, data) {
		return v2InferExcelColumns(data)
	}
	return v2BuildColumns(req, data)
}

func v2InferExcelColumns(data models.DynamicData) []v2Column {
	out := make([]v2Column, 0, len(data.Order))
	for _, k := range data.Order {
		col := v2Column{Field: k, Title: k}
		// Heuristic: if any row contains a string that looks like an Excel formula (starts with '='),
		// infer the column as Formula so it is written using SetCellFormula rather than as raw text.
		for i := range data.Items {
			row := data.Items[i]
			if row == nil {
				continue
			}
			val, ok := row[k]
			if !ok {
				// Best-effort: support case-insensitive key lookups when Order uses a different casing.
				for kk, vv := range row {
					if strings.EqualFold(kk, k) {
						val = vv
						ok = true
						break
					}
				}
			}
			if !ok {
				continue
			}
			s, isStr := val.(string)
			if !isStr {
				continue
			}
			st := strings.TrimSpace(s)
			if strings.HasPrefix(st, "=") && len(st) > 1 {
				col.CellType = models.ExcelCellFormula
				break
			}
		}
		out = append(out, col)
	}
	return out
}

func v2ShouldUseConfiguredColumns(cols []models.ColumnConfig, data models.DynamicData) bool {
	if len(cols) == 0 {
		return false
	}
	keys := v2DataKeySet(data)
	// If the dataset has no keys (shouldn't happen due to validation), keep configured columns.
	if len(keys) == 0 {
		return true
	}
	configuredFields := make(map[string]struct{}, len(cols))
	for i := range cols {
		f := strings.ToLower(strings.TrimSpace(cols[i].Field))
		if f != "" {
			configuredFields[f] = struct{}{}
		}
	}

	applicable := make(map[string]bool, len(cols))
	// Seed applicability based on direct data fields and non-row-dependent computed columns.
	for i := range cols {
		c := cols[i]
		field := strings.ToLower(strings.TrimSpace(c.Field))
		if field == "" {
			continue
		}
		if _, ok := keys[field]; ok {
			applicable[field] = true
			continue
		}
		if strings.TrimSpace(c.SheetFormula) != "" {
			applicable[field] = true
			continue
		}
		if c.CellType == models.ExcelCellLookup && c.Lookup != nil {
			k := strings.ToLower(strings.TrimSpace(c.Lookup.KeyField))
			if k != "" {
				if _, ok := keys[k]; ok {
					applicable[field] = true
					continue
				}
			}
		}
	}

	// Iteratively mark computed columns as applicable when their dependencies are available.
	changed := true
	for changed {
		changed = false
		for i := range cols {
			c := cols[i]
			field := strings.ToLower(strings.TrimSpace(c.Field))
			if field == "" || applicable[field] {
				continue
			}

			if strings.TrimSpace(c.Formula) != "" {
				ok := true
				for _, tok := range v2ExtractIdentTokens(c.Formula) {
					tl := strings.ToLower(tok)
					if tl == "" {
						continue
					}
					if _, inData := keys[tl]; inData {
						continue
					}
					// Only treat tokens that match known column fields as dependencies.
					if _, isField := configuredFields[tl]; isField {
						if applicable[tl] {
							continue
						}
						ok = false
						break
					}
					// Unknown tokens are treated as functions/constants (e.g., ROUND, IF).
				}
				if ok {
					applicable[field] = true
					changed = true
					continue
				}
			}

			if strings.TrimSpace(c.PercentageOf) != "" {
				ref := strings.ToLower(strings.TrimSpace(c.PercentageOf))
				if ref != "" {
					if _, inData := keys[ref]; inData || applicable[ref] {
						applicable[field] = true
						changed = true
						continue
					}
				}
			}
		}
	}

	applicableCount := 0
	for _, v := range applicable {
		if v {
			applicableCount++
		}
	}
	// Heuristic: require at least 2 applicable columns for multi-field datasets.
	// This keeps configured columns on sheets like Purchases, while allowing helper sheets
	// (Products/Fleet/Logistics/Summary) to infer columns naturally.
	if len(keys) == 1 {
		return applicableCount >= 1
	}
	return applicableCount >= 2
}

func v2DataKeySet(data models.DynamicData) map[string]struct{} {
	keys := map[string]struct{}{}
	for _, k := range data.Order {
		v := strings.ToLower(strings.TrimSpace(k))
		if v != "" {
			keys[v] = struct{}{}
		}
	}
	if len(keys) == 0 && len(data.Items) > 0 {
		for k := range data.Items[0] {
			v := strings.ToLower(strings.TrimSpace(k))
			if v != "" {
				keys[v] = struct{}{}
			}
		}
	}
	return keys
}

func v2ExtractIdentTokens(expr string) []string {
	out := make([]string, 0, 8)
	in := strings.TrimSpace(expr)
	for i := 0; i < len(in); {
		r := rune(in[i])
		if isIdentStart(r) {
			j := i + 1
			for j < len(in) {
				r2 := rune(in[j])
				if !isIdentContinue(r2) {
					break
				}
				j++
			}
			out = append(out, in[i:j])
			i = j
			continue
		}
		i++
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

func v2ExcelFormulaLiteral(v any) (string, bool) {
	if v == nil {
		return "", false
	}
	switch vv := v.(type) {
	case string:
		s := strings.TrimSpace(vv)
		if s == "" {
			return "", false
		}
		// Escape double quotes in Excel string literals by doubling.
		es := strings.ReplaceAll(s, "\"", "\"\"")
		return "\"" + es + "\"", true
	case bool:
		if vv {
			return "TRUE", true
		}
		return "FALSE", true
	case int:
		return strconv.Itoa(vv), true
	case int64:
		return fmt.Sprintf("%d", vv), true
	case float64:
		if vv == math.Trunc(vv) {
			return fmt.Sprintf("%.0f", vv), true
		}
		return fmt.Sprintf("%v", vv), true
	case json.Number:
		return vv.String(), true
	default:
		// Fallback: treat as string.
		s := strings.TrimSpace(v2AnyToString(vv))
		if s == "" {
			return "", false
		}
		es := strings.ReplaceAll(s, "\"", "\"\"")
		return "\"" + es + "\"", true
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
	f = excelize.NewFile()
	defer func() { _ = f.Close() }()

	reg := newExcelSheetRegistry()
	pending := make([]pendingExcelFormula, 0)
	pendingLookups := make([]pendingExcelLookup, 0)

	// Determine sheets.
	targetSheets := []models.SheetConfig{}
	if req.Layout != nil && len(req.Layout.Sheets) > 0 {
		targetSheets = req.Layout.Sheets
		// Excelize creates a default sheet (usually "Sheet1"). When layout.sheets is provided,
		// we must keep only the declared sheets and remove the default.
		if err := v2ExcelEnsureOnlySheets(f, targetSheets); err != nil {
			return nil, err
		}
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
		cols := v2BuildExcelColumns(req, data)
		if err := v2ApplyExcelPageSetup(f, sheetName, req); err != nil {
			return nil, err
		}
		if req.Layout != nil {
			freezeRows := 0
			if req.Layout.FreezeHeader {
				freezeRows = 1
			}
			freezeCols := req.Layout.FreezeColumns
			if freezeRows > 0 || freezeCols > 0 {
				colName, _ := excelize.ColumnNumberToName(maxInt(1, freezeCols+1))
				rowIdx := maxInt(1, freezeRows+1)
				topLeft := fmt.Sprintf("%s%d", colName, rowIdx)
				pane := "bottomLeft"
				if freezeRows > 0 && freezeCols > 0 {
					pane = "bottomRight"
				} else if freezeCols > 0 {
					pane = "topRight"
				}
				_ = f.SetPanes(sheetName, &excelize.Panes{
					Freeze:      true,
					Split:       false,
					XSplit:      freezeCols,
					YSplit:      freezeRows,
					TopLeftCell: topLeft,
					ActivePane:  pane,
					Selection:   []excelize.Selection{{SQRef: topLeft, ActiveCell: topLeft, Pane: pane}},
				})
			}
		}

		meta, pend, pendLookups, err := v2RenderExcelSheet(f, sheetName, req, cols, data)
		if err != nil {
			return nil, err
		}
		reg.Register(meta)
		pending = append(pending, pend...)
		pendingLookups = append(pendingLookups, pendLookups...)
	}

	// Apply pending lookups after all sheets are rendered/registered.
	if err := v2ApplyPendingLookups(f, pendingLookups, reg); err != nil {
		return nil, err
	}

	// Resolve any pending cross-sheet formulas after all sheets are rendered.
	for i := range pending {
		p := pending[i]
		resolved, err := v2ResolveSheetTokens(p.Formula, reg)
		if err != nil {
			return nil, err
		}
		if err := f.SetCellFormula(p.Sheet, p.Cell, resolved); err != nil {
			return nil, err
		}
	}

	if err := v2ApplyExcelCharts(f, req, reg); err != nil {
		return nil, err
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func v2ApplyExcelCharts(f *excelize.File, req models.DocumentRequest, reg *excelSheetRegistry) error {
	if f == nil || req.Layout == nil || len(req.Layout.Charts) == 0 {
		return nil
	}
	defaultSheet := ""
	if len(req.Layout.Sheets) > 0 {
		defaultSheet = strings.TrimSpace(req.Layout.Sheets[0].Name)
	}
	if defaultSheet == "" {
		defaultSheet = f.GetSheetName(0)
		if strings.TrimSpace(defaultSheet) == "" {
			defaultSheet = "Sheet1"
		}
	}

	toExcelizeType := func(t models.ChartTypeEnum) excelize.ChartType {
		switch t {
		case models.ChartBar:
			return excelize.Bar
		case models.ChartLine:
			return excelize.Line
		case models.ChartPie:
			return excelize.Pie
		case models.ChartArea:
			return excelize.Area
		case models.ChartColumn:
			return excelize.Col
		default:
			return excelize.Col
		}
	}

	for i := range req.Layout.Charts {
		ch := req.Layout.Charts[i]
		sheet := strings.TrimSpace(ch.Sheet)
		if sheet == "" {
			sheet = defaultSheet
		}
		pos := strings.TrimSpace(ch.Position)
		if pos == "" {
			pos = "E2"
		}
		meta, ok := reg.Get(sheet)
		if !ok {
			return fmt.Errorf("charts[%d]: unknown sheet '%s'", i, sheet)
		}

		catCol, ok := meta.FieldToCol[strings.ToLower(strings.TrimSpace(ch.CategoryField))]
		if !ok {
			return fmt.Errorf("charts[%d]: categoryField '%s' not found in sheet '%s'", i, ch.CategoryField, meta.Name)
		}
		valCol, ok := meta.FieldToCol[strings.ToLower(strings.TrimSpace(ch.ValueField))]
		if !ok {
			return fmt.Errorf("charts[%d]: valueField '%s' not found in sheet '%s'", i, ch.ValueField, meta.Name)
		}
		start := maxInt(2, meta.DataStartRow)
		end := meta.DataEndRow
		if end < start {
			return fmt.Errorf("charts[%d]: sheet '%s' has no data rows for chart", i, meta.Name)
		}

		qSheet := excelQuoteSheetName(meta.Name)
		categories := fmt.Sprintf("%s!$%s$%d:$%s$%d", qSheet, catCol, start, catCol, end)
		values := fmt.Sprintf("%s!$%s$%d:$%s$%d", qSheet, valCol, start, valCol, end)
		name := fmt.Sprintf("%s!$%s$1", qSheet, valCol)

		excelChart := &excelize.Chart{
			Type: toExcelizeType(ch.Type),
			Series: []excelize.ChartSeries{{
				Name:       name,
				Categories: categories,
				Values:     values,
			}},
		}
		if strings.TrimSpace(ch.Title) != "" {
			excelChart.Title = []excelize.RichTextRun{{Text: strings.TrimSpace(ch.Title)}}
		}

		if err := f.AddChart(meta.Name, pos, excelChart); err != nil {
			return fmt.Errorf("charts[%d]: %w", i, err)
		}
	}
	return nil
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

func v2RenderExcelSheet(f *excelize.File, sheet string, req models.DocumentRequest, cols []v2Column, data models.DynamicData) (excelSheetMeta, []pendingExcelFormula, []pendingExcelLookup, error) {
	hStyle := v2StyleHeader(req)
	bStyle := v2StyleBody(req)

	meta := excelSheetMeta{Name: sheet, FieldToCol: map[string]string{}, DataStartRow: 2}
	pending := make([]pendingExcelFormula, 0)
	pendingLookups := make([]pendingExcelLookup, 0)

	if req.Layout != nil {
		if req.Layout.FreezeColumns > 0 && req.Layout.FreezeColumns > len(cols) {
			return excelSheetMeta{}, nil, nil, fmt.Errorf("freezeColumns must be <= number of columns")
		}
		if req.Layout.MaxVisibleRows > 0 && len(data.Items) > req.Layout.MaxVisibleRows {
			return excelSheetMeta{}, nil, nil, fmt.Errorf("maxVisibleRows must be >= number of records")
		}
	}

	// Cache style IDs by key.
	styleCache := map[string]int{}
	getStyle := func(isHeader bool, fillHex string, fontHex string, hAlign string, vAlign string, numFmt *string, locked bool) (int, error) {
		key := fmt.Sprintf("h=%v|fill=%s|font=%s|ha=%s|va=%s|nf=%v|locked=%v", isHeader, fillHex, fontHex, hAlign, vAlign, numFmt, locked)
		if id, ok := styleCache[key]; ok {
			return id, nil
		}
		base := bStyle
		if isHeader {
			base = hStyle
		}
		fHex := strings.TrimSpace(fontHex)
		if fHex == "" {
			fHex = base.FontColor
		}
		fontColor := v2NormalizeHexColor(fHex)
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
			Protection:   &excelize.Protection{Locked: locked},
			CustomNumFmt: numFmt,
		}
		id, err := f.NewStyle(st)
		if err != nil {
			return 0, err
		}
		styleCache[key] = id
		return id, nil
	}

	// Build field->column mapping (case-insensitive).
	for i := range cols {
		colLetter, _ := excelize.ColumnNumberToName(i + 1)
		meta.FieldToCol[strings.ToLower(strings.TrimSpace(cols[i].Field))] = colLetter
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
		headerFill := hStyle.Background
		if strings.TrimSpace(c.HeaderColor) != "" {
			headerFill = c.HeaderColor
		}
		hID, err := getStyle(true, headerFill, "", excelHAlign(hStyle.Alignment), excelVAlign(hStyle.VerticalAlign), nil, true)
		if err != nil {
			return excelSheetMeta{}, nil, nil, err
		}
		_ = f.SetCellValue(sheet, cell, c.Title)
		_ = f.SetCellStyle(sheet, cell, cell, hID)

		if c.Hidden {
			colLetter, _ := excelize.ColumnNumberToName(i + 1)
			_ = f.SetColVisible(sheet, colLetter, false)
		}

		width := c.Width
		if req.Layout != nil && req.Layout.AutoSizeColumns {
			width = computedWidths[i]
		}
		if width > 0 {
			colLetter, _ := excelize.ColumnNumberToName(i + 1)
			_ = f.SetColWidth(sheet, colLetter, colLetter, width)
		}
	}

	// Optional grouping (layout.groupBy): sort and insert subtotal rows.
	groupBy := ""
	if req.Layout != nil {
		groupBy = strings.TrimSpace(req.Layout.GroupBy)
	}
	items := data.Items
	groupByKey := ""
	getGroupValue := func(item map[string]any) string {
		if groupBy == "" || item == nil {
			return ""
		}
		if groupByKey != "" {
			if v, ok := item[groupByKey]; ok {
				return v2AnyToString(v)
			}
		}
		for k, v := range item {
			if strings.EqualFold(k, groupBy) {
				return v2AnyToString(v)
			}
		}
		return ""
	}
	if groupBy != "" && len(items) > 0 {
		// groupBy is a layout-global setting, but in multi-sheet workbooks some sheets
		// may not contain this field (e.g., Summary/FuelSummary). In that case, we
		// skip grouping for the sheet instead of failing the whole render.
		found := false
		for _, it := range items {
			if it == nil {
				continue
			}
			if _, ok := it[groupBy]; ok {
				groupByKey = groupBy
				found = true
				break
			}
			for k := range it {
				if strings.EqualFold(k, groupBy) {
					groupByKey = k
					found = true
					break
				}
			}
			if found {
				break
			}
		}
		if !found {
			groupBy = ""
			groupByKey = ""
		}
	}
	if groupBy != "" && len(items) > 0 {
		cpy := make([]map[string]any, 0, len(items))
		cpy = append(cpy, items...)
		sort.SliceStable(cpy, func(i, j int) bool {
			a := getGroupValue(cpy[i])
			b := getGroupValue(cpy[j])
			return a < b
		})
		items = cpy
	}

	oddBg := bStyle.ZebraColorOdd
	evenBg := bStyle.ZebraColorEven
	dataRowMax := 1
	percentagePending := make([]struct {
		row          int
		colLetter    string
		refColLetter string
	}, 0)

	writeRow := func(excelRow int, item map[string]any, isSubtotalOrTotal bool, zebraIdx int) error {
		maxLines := 1
		for cIdx, c := range cols {
			cell, _ := excelize.CoordinatesToCellName(cIdx+1, excelRow)
			numFmtEnum := c.Format
			if strings.TrimSpace(c.PercentageOf) != "" && numFmtEnum == "" {
				numFmtEnum = models.ColFormatPercentage
			}

			if !isSubtotalOrTotal {
				val := any(nil)
				if item != nil {
					val = item[c.Field]
				}
				setVal := val
				// Formulas
				if strings.TrimSpace(c.SheetFormula) != "" {
					pending = append(pending, pendingExcelFormula{Sheet: sheet, Cell: cell, Formula: strings.TrimSpace(c.SheetFormula)})
					setVal = ""
				} else if strings.TrimSpace(c.Formula) != "" {
					rowFormula, err := v2ResolveRowFormula(strings.TrimSpace(c.Formula), meta.FieldToCol, excelRow)
					if err != nil {
						return err
					}
					if strings.Contains(rowFormula, "sheet:") {
						pending = append(pending, pendingExcelFormula{Sheet: sheet, Cell: cell, Formula: rowFormula})
					} else {
						_ = f.SetCellFormula(sheet, cell, rowFormula)
					}
					setVal = ""
				} else if strings.TrimSpace(c.PercentageOf) != "" {
					ref := strings.ToLower(strings.TrimSpace(c.PercentageOf))
					refCol, ok := meta.FieldToCol[ref]
					if !ok {
						return fmt.Errorf("column '%s': percentageOf references unknown field '%s'", c.Field, c.PercentageOf)
					}
					colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
					percentagePending = append(percentagePending, struct {
						row          int
						colLetter    string
						refColLetter string
					}{row: excelRow, colLetter: colLetter, refColLetter: refCol})
					setVal = ""
				} else {
					// Set raw value for better formatting when possible.
					didSetFormula := false
					switch c.CellType {
					case models.ExcelCellText:
						setVal = v2AnyToString(val)
					case models.ExcelCellNumber:
						numFmtEnum = models.ColFormatNumber
						if parsed, ok := v2CoerceForExcel(models.ColFormatNumber, val); ok {
							setVal = parsed
						}
					case models.ExcelCellCurrency:
						numFmtEnum = models.ColFormatCurrency
						if parsed, ok := v2CoerceForExcel(models.ColFormatCurrency, val); ok {
							setVal = parsed
						}
					case models.ExcelCellDate:
						numFmtEnum = models.ColFormatDate
						if s, ok := val.(string); ok {
							if tt, err := time.Parse("2006-01-02", strings.TrimSpace(s)); err == nil {
								setVal = tt
							}
						}
					case models.ExcelCellSelect:
						setVal = v2AnyToString(val)
					case models.ExcelCellFormula:
						fx := strings.TrimSpace(v2AnyToString(val))
						if strings.HasPrefix(fx, "=") {
							fx = strings.TrimSpace(fx[1:])
						}
						if fx != "" {
							if strings.Contains(fx, "sheet:") {
								pending = append(pending, pendingExcelFormula{Sheet: sheet, Cell: cell, Formula: fx})
							} else {
								_ = f.SetCellFormula(sheet, cell, fx)
							}
							didSetFormula = true
						}
						setVal = ""
					case models.ExcelCellLookup:
						if c.Lookup == nil {
							return fmt.Errorf("column '%s': lookup config is required when cellType is Lookup", c.Field)
						}
						keyField := strings.ToLower(strings.TrimSpace(c.Lookup.KeyField))
						keyExpr := ""
						if keyCol, ok := meta.FieldToCol[keyField]; ok {
							keyExpr = fmt.Sprintf("%s%d", keyCol, excelRow)
						} else {
							// If the keyField isn't a rendered column, use the row value as a literal.
							// This supports hidden keys (e.g., productId exists in data but not in columns).
							keyVal := any(nil)
							if item != nil {
								keyVal = item[c.Lookup.KeyField]
								if keyVal == nil {
									// best-effort case-insensitive lookup
									for k, v := range item {
										if strings.EqualFold(k, c.Lookup.KeyField) {
											keyVal = v
											break
										}
									}
								}
							}
							if lit, ok := v2ExcelFormulaLiteral(keyVal); ok {
								keyExpr = lit
							}
						}
						if strings.TrimSpace(keyExpr) == "" {
							// No key available for this row; don't error.
							setVal = v2AnyToString(val)
							break
						}
						pendingLookups = append(pendingLookups, pendingExcelLookup{Sheet: sheet, Cell: cell, KeyCell: keyExpr, Config: *c.Lookup})
						didSetFormula = true
						setVal = ""
					default:
						if c.Format != "" {
							if parsed, ok := v2CoerceForExcel(c.Format, val); ok {
								setVal = parsed
							}
						}
					}
					if !didSetFormula {
						_ = f.SetCellValue(sheet, cell, setVal)
					}
				}

				// Estimate row height based on the longest wrapped cell.
				colWidthChars := 12.0
				if req.Layout != nil {
					if req.Layout.AutoSizeColumns && cIdx < len(computedWidths) && computedWidths[cIdx] > 0 {
						colWidthChars = computedWidths[cIdx]
					} else if cols[cIdx].Width > 0 {
						colWidthChars = cols[cIdx].Width
					}
				}
				cellLines := v2ExcelWrappedLineCount(v2AnyToString(setVal), colWidthChars)
				if cellLines > maxLines {
					maxLines = cellLines
				}
			}

			fill := bStyle.Background
			if isSubtotalOrTotal {
				fill = hStyle.Background
			} else if bStyle.ZebraStripe {
				if zebraIdx%2 == 0 {
					fill = evenBg
				} else {
					fill = oddBg
				}
			}
			if !isSubtotalOrTotal && strings.TrimSpace(c.BackgroundColor) != "" {
				fill = c.BackgroundColor
			}
			fontOverride := ""
			if !isSubtotalOrTotal && strings.TrimSpace(c.TextColor) != "" {
				fontOverride = c.TextColor
			}
			ha := excelHAlign(bStyle.Alignment)
			if isSubtotalOrTotal {
				ha = excelHAlign(hStyle.Alignment)
			}
			if c.Align != "" {
				ha = excelColAlign(c.Align)
			}
			va := excelVAlign(bStyle.VerticalAlign)
			if isSubtotalOrTotal {
				va = excelVAlign(hStyle.VerticalAlign)
			}
			if c.VAlign != "" {
				va = excelVAlign(c.VAlign)
			}
			nf := v2ExcelCustomNumFmt(numFmtEnum)
			locked := c.Locked
			isHeaderStyle := false
			if isSubtotalOrTotal {
				isHeaderStyle = true
				locked = true
			}
			bID, err := getStyle(isHeaderStyle, fill, fontOverride, ha, va, nf, locked)
			if err != nil {
				return err
			}
			_ = f.SetCellStyle(sheet, cell, cell, bID)
		}
		if !isSubtotalOrTotal && maxLines > 1 {
			_ = f.SetRowHeight(sheet, excelRow, v2ExcelRowHeightPoints(bStyle.FontSize, maxLines))
		}
		return nil
	}

	currentRow := 2
	var (
		curGroup        string
		groupStartRow   int
		zebraDataIdx    int
		subtotalRows    []int
		hasAggregations bool
	)
	for _, c := range cols {
		if strings.ToLower(strings.TrimSpace(c.Aggregate)) == "sum" {
			hasAggregations = true
			break
		}
		if strings.TrimSpace(c.PercentageOf) != "" {
			hasAggregations = true
			break
		}
	}

	if groupBy != "" {
		curGroup = "__init__"
		groupStartRow = currentRow
	}

	for idx, item := range items {
		if groupBy != "" {
			g := getGroupValue(item)
			if idx == 0 {
				curGroup = g
				groupStartRow = currentRow
			} else if g != curGroup {
				// subtotal for previous group
				subRow := currentRow
				_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", subRow), fmt.Sprintf("Subtotal %s", curGroup))
				for cIdx, c := range cols {
					if strings.ToLower(strings.TrimSpace(c.Aggregate)) != "sum" {
						continue
					}
					colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
					formula := fmt.Sprintf("SUM(%s%d:%s%d)", colLetter, groupStartRow, colLetter, subRow-1)
					cell, _ := excelize.CoordinatesToCellName(cIdx+1, subRow)
					_ = f.SetCellFormula(sheet, cell, formula)
				}
				if err := writeRow(subRow, nil, true, zebraDataIdx); err != nil {
					return excelSheetMeta{}, nil, nil, err
				}
				subtotalRows = append(subtotalRows, subRow)
				currentRow++
				curGroup = g
				groupStartRow = currentRow
			}
		}

		if err := writeRow(currentRow, item, false, zebraDataIdx); err != nil {
			return excelSheetMeta{}, nil, nil, err
		}
		if currentRow > dataRowMax {
			dataRowMax = currentRow
		}
		zebraDataIdx++
		currentRow++
	}

	if groupBy != "" && len(items) > 0 {
		subRow := currentRow
		_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", subRow), fmt.Sprintf("Subtotal %s", curGroup))
		for cIdx, c := range cols {
			if strings.ToLower(strings.TrimSpace(c.Aggregate)) != "sum" {
				continue
			}
			colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
			formula := fmt.Sprintf("SUM(%s%d:%s%d)", colLetter, groupStartRow, colLetter, subRow-1)
			cell, _ := excelize.CoordinatesToCellName(cIdx+1, subRow)
			_ = f.SetCellFormula(sheet, cell, formula)
		}
		if err := writeRow(subRow, nil, true, zebraDataIdx); err != nil {
			return excelSheetMeta{}, nil, nil, err
		}
		subtotalRows = append(subtotalRows, subRow)
		currentRow++
	}

	showTotal := false
	if req.Layout != nil && req.Layout.ShowTotalRow {
		showTotal = true
	}
	if hasAggregations {
		showTotal = true
	}
	if showTotal {
		meta.TotalRow = currentRow
		_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", meta.TotalRow), "TOTAL")
		for cIdx, c := range cols {
			if strings.ToLower(strings.TrimSpace(c.Aggregate)) != "sum" {
				continue
			}
			colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
			var formula string
			if len(subtotalRows) > 0 {
				parts := make([]string, 0, len(subtotalRows))
				for _, rr := range subtotalRows {
					parts = append(parts, fmt.Sprintf("%s%d", colLetter, rr))
				}
				formula = fmt.Sprintf("SUM(%s)", strings.Join(parts, ","))
			} else {
				start := 2
				end := dataRowMax
				if end < start {
					end = start
				}
				formula = fmt.Sprintf("SUM(%s%d:%s%d)", colLetter, start, colLetter, end)
			}
			cell, _ := excelize.CoordinatesToCellName(cIdx+1, meta.TotalRow)
			_ = f.SetCellFormula(sheet, cell, formula)
		}
		if err := writeRow(meta.TotalRow, nil, true, zebraDataIdx); err != nil {
			return excelSheetMeta{}, nil, nil, err
		}
		currentRow++
	}

	// Apply percentageOf formulas after total row exists.
	if len(percentagePending) > 0 {
		if meta.TotalRow == 0 {
			return excelSheetMeta{}, nil, nil, fmt.Errorf("percentageOf requires a total row; enable layout.showTotalRow or add an aggregate")
		}
		for _, p := range percentagePending {
			cell := fmt.Sprintf("%s%d", p.colLetter, p.row)
			formula := fmt.Sprintf("%s%d/$%s$%d", p.refColLetter, p.row, p.refColLetter, meta.TotalRow)
			_ = f.SetCellFormula(sheet, cell, formula)
		}
	}

	meta.SubtotalRows = subtotalRows
	meta.HasSubtotals = len(subtotalRows) > 0
	meta.DataEndRow = maxInt(1, currentRow-1)
	meta.GeneratedRows = meta.DataEndRow
	if meta.TotalRow > 0 {
		meta.DataEndRow = meta.TotalRow - 1
		meta.GeneratedRows = meta.TotalRow
	}

	// Select/dropdown validation (per column).
	lastRow := maxInt(2, dataRowMax)
	for cIdx, c := range cols {
		if c.CellType != models.ExcelCellSelect {
			continue
		}
		colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
		sqref := fmt.Sprintf("%s2:%s%d", colLetter, colLetter, maxInt(2, lastRow))
		dv := excelize.NewDataValidation(true)
		dv.SetSqref(sqref)
		if strings.TrimSpace(c.ValidationRange) != "" {
			dv.SetSqrefDropList(strings.TrimSpace(c.ValidationRange))
		} else {
			if len(c.Options) == 0 {
				return excelSheetMeta{}, nil, nil, fmt.Errorf("column '%s': options or validationRange is required when cellType is Select", c.Field)
			}
			if err := dv.SetDropList(c.Options); err != nil {
				return excelSheetMeta{}, nil, nil, err
			}
		}
		dv.SetError(excelize.DataValidationErrorStyleStop, "Invalid value", "Select a value from the list")
		_ = f.AddDataValidation(sheet, dv)
	}

	// Conditional formatting (per column).
	if lastRow >= 2 {
		condStyleCache := map[string]int{}
		newCondStyle := func(bg string, fg string) (*int, error) {
			bg = strings.TrimSpace(bg)
			fg = strings.TrimSpace(fg)
			if bg == "" && fg == "" {
				return nil, nil
			}
			key := fmt.Sprintf("bg=%s|fg=%s", bg, fg)
			if id, ok := condStyleCache[key]; ok {
				return &id, nil
			}
			st := &excelize.Style{}
			if fg != "" {
				st.Font = &excelize.Font{Color: v2NormalizeHexColor(fg)}
			}
			if bg != "" {
				st.Fill = excelize.Fill{Type: "pattern", Color: []string{v2NormalizeHexColor(bg)}, Pattern: 1}
			}
			id, err := f.NewConditionalStyle(st)
			if err != nil {
				return nil, err
			}
			condStyleCache[key] = id
			return &id, nil
		}

		for cIdx, c := range cols {
			if len(c.ConditionalFormatting) == 0 {
				continue
			}
			colLetter, _ := excelize.ColumnNumberToName(cIdx + 1)
			rng := fmt.Sprintf("%s2:%s%d", colLetter, colLetter, lastRow)
			opts := make([]excelize.ConditionalFormatOptions, 0, len(c.ConditionalFormatting))
			for _, r := range c.ConditionalFormatting {
				fmtID, err := newCondStyle(r.BackgroundColor, r.TextColor)
				if err != nil {
					return excelSheetMeta{}, nil, nil, err
				}
				switch r.Operator {
				case models.ExcelCondGreaterThan:
					opts = append(opts, excelize.ConditionalFormatOptions{Type: "cell", Criteria: ">", Format: fmtID, Value: strings.TrimSpace(r.Value)})
				case models.ExcelCondLessThan:
					opts = append(opts, excelize.ConditionalFormatOptions{Type: "cell", Criteria: "<", Format: fmtID, Value: strings.TrimSpace(r.Value)})
				case models.ExcelCondEqual:
					opts = append(opts, excelize.ConditionalFormatOptions{Type: "cell", Criteria: "==", Format: fmtID, Value: strings.TrimSpace(r.Value)})
				case models.ExcelCondContainsText:
					opts = append(opts, excelize.ConditionalFormatOptions{Type: "text", Criteria: "containsText", Format: fmtID, Value: strings.TrimSpace(r.Value)})
				case models.ExcelCondFormula:
					opts = append(opts, excelize.ConditionalFormatOptions{Type: "formula", Criteria: strings.TrimSpace(r.Formula), Format: fmtID})
				default:
					return excelSheetMeta{}, nil, nil, fmt.Errorf("column '%s': unknown conditional operator '%s'", c.Field, r.Operator)
				}
			}
			if err := f.SetConditionalFormat(sheet, rng, opts); err != nil {
				return excelSheetMeta{}, nil, nil, err
			}
		}
	}

	// Hide rows and columns beyond visible limit.
	if req.Layout != nil {
		visibleDataRows := 0
		if req.Layout.MaxVisibleRows > 0 {
			visibleDataRows = req.Layout.MaxVisibleRows
		} else if req.Layout.HideEmptyRows {
			visibleDataRows = len(data.Items)
		}
		if visibleDataRows > 0 {
			// Efficiently hide all rows by default, then re-enable only the visible range.
			// This avoids writing per-row hidden state up to 1,048,576.
			zero := true
			_ = f.SetSheetProps(sheet, &excelize.SheetPropsOptions{ZeroHeight: &zero})

			lastVisibleRow := visibleDataRows + 1 // +1 header
			if meta.GeneratedRows > lastVisibleRow {
				lastVisibleRow = meta.GeneratedRows
			}
			for rr := 1; rr <= lastVisibleRow; rr++ {
				_ = f.SetRowVisible(sheet, rr, true)
			}

			// Hide all columns after the last generated column (up to XFD).
			if len(cols) > 0 && len(cols) < 16384 {
				startCol, _ := excelize.ColumnNumberToName(len(cols) + 1)
				_ = f.SetColVisible(sheet, startCol+":XFD", false)
			}
		}
	}
	return meta, pending, pendingLookups, nil
}

func v2ExcelWrappedLineCount(s string, widthChars float64) int {
	text := strings.ReplaceAll(s, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	if strings.TrimSpace(text) == "" {
		return 1
	}
	if widthChars <= 0 {
		widthChars = 12
	}
	width := int(math.Max(1, math.Floor(widthChars)))
	lines := 0
	for _, seg := range strings.Split(text, "\n") {
		r := utf8.RuneCountInString(seg)
		if r == 0 {
			lines += 1
			continue
		}
		lines += int(math.Ceil(float64(r) / float64(width)))
	}
	if lines < 1 {
		lines = 1
	}
	return lines
}

func v2ExcelRowHeightPoints(fontSize int, lines int) float64 {
	fs := float64(maxInt(1, fontSize))
	if lines < 1 {
		lines = 1
	}
	// A pragmatic approximation: Excel row height is in points; line height is typically ~1.2-1.3x font size.
	lineH := fs*1.25 + 1
	h := lineH * float64(lines)
	if h < 0 {
		return 0
	}
	// Guardrail: avoid absurd heights.
	if h > 409 {
		h = 409
	}
	return h
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
	hStyle := v2StyleHeader(req)
	bStyle := v2StyleBody(req)
	fStyle := v2StyleFooter(req)

	// Builder mode with Index (TOC) requires a two-pass render so we can compute
	// page numbers and create internal links.
	if req.Layout != nil && len(req.Layout.Blocks) > 0 {
		if v2PDFHasIndexBlock(req.Layout.Blocks) {
			return v2PDFRenderBlocksWithIndex(req, hStyle, bStyle, fStyle)
		}
	}

	pdf := v2PDFNew(req, fStyle)
	pdf.AddPage()

	// Builder mode: render ordered blocks sequentially.
	if req.Layout != nil && len(req.Layout.Blocks) > 0 {
		if err := v2PDFRenderBlocks(pdf, req, hStyle, bStyle); err != nil {
			return nil, err
		}
		var out bytes.Buffer
		if err := pdf.Output(&out); err != nil {
			return nil, err
		}
		return out.Bytes(), nil
	}

	// Legacy mode: render a single table from the default dataset.
	data := req.Data.Get("")

	pageW, _ := pdf.GetPageSize()
	lm, _, rm, _ := pdf.GetMargins()
	usable := pageW - lm - rm
	widths := v2ComputePDFWidths(cols, usable)

	if req.Layout != nil && req.Layout.Spacing != nil && req.Layout.Spacing.TableSpacing > 0 {
		pdf.Ln(req.Layout.Spacing.TableSpacing)
	}

	renderHeader := func() {
		v2PDFDrawTableHeader(pdf, req, cols, widths, hStyle)
	}

	newPageWithHeader := func() error {
		pdf.AddPage()
		renderHeader()
		return nil
	}

	ensurePage := func(needH float64) error {
		_, pageH := pdf.GetPageSize()
		_, _, _, bm := pdf.GetMargins()
		bottomY := pageH - bm
		if pdf.GetY()+needH <= bottomY {
			return nil
		}
		return newPageWithHeader()
	}

	// Print initial header (may auto-break later for long docs)
	renderHeader()

	rowsPerPage := 0
	if req.Layout != nil && req.Layout.PageBreak != nil && req.Layout.PageBreak.Enabled {
		rowsPerPage = req.Layout.PageBreak.RowsPerPage
	}
	printed := 0
	for r, item := range data.Items {
		// Fixed rowsPerPage mode.
		if rowsPerPage > 0 && printed > 0 && printed%rowsPerPage == 0 {
			if err := newPageWithHeader(); err != nil {
				return nil, err
			}
		}

		fillColor := bStyle.Background
		if bStyle.ZebraStripe {
			if r%2 == 0 {
				fillColor = bStyle.ZebraColorEven
			} else {
				fillColor = bStyle.ZebraColorOdd
			}
		}
		rowH := v2PDFRowHeight(pdf, req, cols, widths, item, bStyle)
		if err := ensurePage(rowH); err != nil {
			return nil, err
		}
		if err := v2PDFDrawTableRow(pdf, req, cols, widths, item, bStyle, fillColor); err != nil {
			return nil, err
		}
		printed++
	}

	var out bytes.Buffer
	if err := pdf.Output(&out); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func (s *DocumentService) v2RenderWord(req models.DocumentRequest, cols []v2Column) ([]byte, error) {
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

	// Render body and collect images (header image + layout.blocks images/charts).
	images := make([]v2WordImagePart, 0, 4)
	body := new(strings.Builder)
	nextImage := 1
	nextDocPrID := 1

	// Optional header image.
	if req.Layout != nil && req.Layout.HeaderImage != nil && strings.TrimSpace(req.Layout.HeaderImage.Data) != "" {
		imgBytes, ext, err := v2DecodeImageBytes(req.Layout.HeaderImage.Data)
		if err != nil {
			return nil, err
		}
		relID := fmt.Sprintf("rIdImage%d", nextImage)
		images = append(images, v2WordImagePart{RelID: relID, Ext: ext, Path: fmt.Sprintf("word/media/image%d%s", nextImage, ext), Bytes: imgBytes})
		body.WriteString(v2WordImageParagraphXML(req))
		nextImage++
		nextDocPrID++
	}

	// Builder mode: layout.blocks. If present, ignore the legacy columns-only renderer.
	if req.Layout != nil && len(req.Layout.Blocks) > 0 {
		blocksXML, blockImages, err := v2WordBlocksBodyXML(req, nextImage, nextDocPrID)
		if err != nil {
			return nil, err
		}
		body.WriteString(blocksXML)
		images = append(images, blockImages...)
	} else {
		body.WriteString(v2WordTablesXML(req, cols))
	}

	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)

	writeZipString(zw, "[Content_Types].xml", v2WordContentTypesXML(includeFooter, images))
	writeZipString(zw, "_rels/.rels", v2WordRootRelsXML())
	writeZipString(zw, "word/styles.xml", v2WordStylesXML())
	writeZipString(zw, "word/document.xml", v2WordDocumentXML(req, body.String(), includeFooter))
	writeZipString(zw, "word/_rels/document.xml.rels", v2WordDocumentRelsXML(includeFooter, images))
	if includeFooter {
		writeZipString(zw, "word/footer1.xml", v2WordFooterXML(v2StyleFooter(req), footerAlign, footerFmt))
	}
	for _, img := range images {
		w, _ := zw.Create(img.Path)
		_, _ = w.Write(img.Bytes)
	}

	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// ---- PDF helpers ----

func v2PDFNew(req models.DocumentRequest, fStyle models.StyleConfig) *gofpdf.Fpdf {
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
	// Ensure header image and top spacing are applied on every page, including auto page breaks.
	var headerCfg *models.HeaderImageConfig
	if req.Layout != nil && req.Layout.HeaderImage != nil && strings.TrimSpace(req.Layout.HeaderImage.Data) != "" {
		cc := *req.Layout.HeaderImage
		headerCfg = &cc
	}
	paragraphSpacing := 0.0
	if req.Layout != nil && req.Layout.Spacing != nil {
		paragraphSpacing = req.Layout.Spacing.ParagraphSpacing
	}
	if headerCfg != nil || paragraphSpacing > 0 {
		pdf.SetHeaderFunc(func() {
			if headerCfg != nil {
				_ = v2AddPDFHeaderImage(pdf, *headerCfg)
			}
			if paragraphSpacing > 0 {
				pdf.Ln(paragraphSpacing)
			}
		})
	}
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
	return pdf
}

func v2PDFHasIndexBlock(blocks []models.PDFBlockConfig) bool {
	count := 0
	for i := range blocks {
		if blocks[i].Type == models.PDFBlockIndex {
			count++
		}
	}
	return count > 0
}

type v2PDFIndexEntry struct {
	Title        string
	PageNo       int
	SectionIndex int
}

func v2PDFIndexTitle(blocks []models.PDFBlockConfig) string {
	for i := range blocks {
		if blocks[i].Type == models.PDFBlockIndex {
			if t := strings.TrimSpace(blocks[i].Content); t != "" {
				return t
			}
			break
		}
	}
	return "Índice"
}

func v2PDFValidateIndexBlocks(blocks []models.PDFBlockConfig) error {
	idx := -1
	count := 0
	for i := range blocks {
		if blocks[i].Type == models.PDFBlockIndex {
			count++
			if idx < 0 {
				idx = i
			}
		}
	}
	if count == 0 {
		return nil
	}
	if count > 1 {
		return fmt.Errorf("only one Index block is supported")
	}
	// Keep it simple/deterministic: Index must be the first block.
	if idx != 0 {
		return fmt.Errorf("Index block must be the first block")
	}
	return nil
}

func v2PDFIndexPageCount(req models.DocumentRequest, fStyle models.StyleConfig, title string, entries int) int {
	if entries <= 0 {
		return 1
	}
	probe := v2PDFNew(req, fStyle)
	probe.AddPage()

	// Title
	titleStyle := v2StyleBody(req)
	titleStyle.FontSize = 16
	titleStyle.Bold = true
	v2SetPDFFont(probe, titleStyle)
	lm, _, _, _ := probe.GetMargins()
	probe.SetX(lm)
	usableW := v2PDFUsableWidth(probe)
	probe.MultiCell(usableW, v2PDFLineHeightMM(titleStyle.FontSize), v2PDFText(title), "", "L", false)
	probe.Ln(4)

	entryStyle := v2StyleBody(req)
	entryStyle.FontSize = 12
	entryStyle.Bold = false
	v2SetPDFFont(probe, entryStyle)
	entryLH := v2PDFLineHeightMM(entryStyle.FontSize)

	availFirst := v2PDFBottomY(probe) - probe.GetY()
	perFirst := int(math.Floor(availFirst / entryLH))
	if perFirst < 1 {
		perFirst = 1
	}

	probe.AddPage()
	v2SetPDFFont(probe, entryStyle)
	availOther := v2PDFBottomY(probe) - probe.GetY()
	perOther := int(math.Floor(availOther / entryLH))
	if perOther < 1 {
		perOther = 1
	}

	if entries <= perFirst {
		return 1
	}
	remaining := entries - perFirst
	extra := int(math.Ceil(float64(remaining) / float64(perOther)))
	return 1 + extra
}

func v2PDFRenderIndex(pdf *gofpdf.Fpdf, req models.DocumentRequest, bStyle models.StyleConfig, title string, entries []v2PDFIndexEntry, pageOffset int) (map[int]int, error) {
	if pdf == nil {
		return nil, fmt.Errorf("pdf is nil")
	}
	links := map[int]int{}

	lm, _, _, _ := pdf.GetMargins()
	usableW := v2PDFUsableWidth(pdf)

	// Title
	titleStyle := bStyle
	titleStyle.FontSize = 16
	titleStyle.Bold = true
	v2SetPDFFont(pdf, titleStyle)
	pdf.SetX(lm)
	pdf.MultiCell(usableW, v2PDFLineHeightMM(titleStyle.FontSize), v2PDFText(title), "", "L", false)
	pdf.Ln(4)

	// Entries
	entryStyle := bStyle
	entryStyle.FontSize = 12
	entryStyle.Bold = false
	v2SetPDFFont(pdf, entryStyle)
	v2SetPDFText(pdf, entryStyle.FontColor)
	entryLH := v2PDFLineHeightMM(entryStyle.FontSize)

	titleW := usableW - 15
	numW := 15.0
	if titleW < 20 {
		titleW = usableW
		numW = 0
	}

	for i := range entries {
		if pdf.GetY()+entryLH > v2PDFBottomY(pdf) {
			pdf.AddPage()
			v2SetPDFFont(pdf, entryStyle)
			v2SetPDFText(pdf, entryStyle.FontColor)
		}
		pno := entries[i].PageNo + pageOffset
		if pno <= 0 {
			pno = 1
		}
		linkID := pdf.AddLink()
		links[entries[i].SectionIndex] = linkID

		pdf.SetX(lm)
		pdf.CellFormat(titleW, entryLH, v2PDFText(entries[i].Title), "", 0, "L", false, linkID, "")
		if numW > 0 {
			pdf.CellFormat(numW, entryLH, fmt.Sprintf("%d", pno), "", 1, "R", false, 0, "")
		} else {
			pdf.Ln(entryLH)
		}
	}
	return links, nil
}

func v2PDFRenderBlocksWithIndex(req models.DocumentRequest, hStyle, bStyle, fStyle models.StyleConfig) ([]byte, error) {
	if req.Layout == nil || len(req.Layout.Blocks) == 0 {
		return nil, fmt.Errorf("layout.blocks is required")
	}
	if err := v2PDFValidateIndexBlocks(req.Layout.Blocks); err != nil {
		return nil, err
	}

	// First pass: render content (skipping Index) to compute section page numbers.
	pdf1 := v2PDFNew(req, fStyle)
	pdf1.AddPage()
	entries := make([]v2PDFIndexEntry, 0, 16)
	if err := v2PDFRenderBlocksInternal(pdf1, req, hStyle, bStyle, v2PDFRenderBlocksOptions{SkipIndex: true, CollectIndex: &entries}); err != nil {
		return nil, err
	}

	idxTitle := v2PDFIndexTitle(req.Layout.Blocks)
	idxPages := v2PDFIndexPageCount(req, fStyle, idxTitle, len(entries))

	// Second pass: render Index pages, then render content (skipping Index) while setting internal links.
	// We render in a short loop to guarantee that the page offset used in the Index matches
	// the actual number of Index pages created.
	for attempt := 0; attempt < 3; attempt++ {
		pdf2 := v2PDFNew(req, fStyle)
		pdf2.AddPage()
		linkIDs, err := v2PDFRenderIndex(pdf2, req, bStyle, idxTitle, entries, idxPages)
		if err != nil {
			return nil, err
		}
		actualIdxPages := pdf2.PageNo()
		if actualIdxPages != idxPages {
			idxPages = actualIdxPages
			continue
		}

		// Force content to start after the Index pages.
		pdf2.AddPage()
		if err := v2PDFRenderBlocksInternal(pdf2, req, hStyle, bStyle, v2PDFRenderBlocksOptions{SkipIndex: true, SectionLinks: linkIDs}); err != nil {
			return nil, err
		}

		var out bytes.Buffer
		if err := pdf2.Output(&out); err != nil {
			return nil, err
		}
		return out.Bytes(), nil
	}
	return nil, fmt.Errorf("failed to render Index with stable page count")
}

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
	case models.AlignJustify:
		return "J"
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

func v2PDFLineHeightMM(fontSize int) float64 {
	// 1pt ~= 0.3528mm
	if fontSize <= 0 {
		fontSize = 11
	}
	h := float64(fontSize) * 0.3528 * 1.25
	if h < 4.0 {
		h = 4.0
	}
	return h
}

func v2PDFSplitLines(pdf *gofpdf.Fpdf, txt string, w float64) []string {
	if w <= 0 {
		return []string{txt}
	}
	lines := pdf.SplitLines([]byte(txt), w)
	out := make([]string, 0, len(lines))
	for _, b := range lines {
		out = append(out, string(b))
	}
	if len(out) == 0 {
		return []string{""}
	}
	return out
}

func v2PDFDrawTableHeader(pdf *gofpdf.Fpdf, req models.DocumentRequest, cols []v2Column, widths []float64, hStyle models.StyleConfig) {
	v2SetPDFFont(pdf, hStyle)
	v2SetPDFFill(pdf, hStyle.Background)
	v2SetPDFText(pdf, hStyle.FontColor)

	lineH := v2PDFLineHeightMM(hStyle.FontSize)
	pad := 1.0
	rowH := lineH + 2*pad

	x0 := pdf.GetX()
	y0 := pdf.GetY()
	x := x0
	for i, c := range cols {
		w := widths[i]
		style := "F"
		if hStyle.Border {
			style = "FD"
		}
		pdf.Rect(x, y0, w, rowH, style)
		pdf.SetXY(x+pad, y0+pad)
		pdf.MultiCell(w-2*pad, lineH, v2PDFText(c.Title), "", v2PDFColAlign(c.Align), false)
		x += w
		pdf.SetXY(x, y0)
	}
	pdf.SetXY(x0, y0+rowH)
}

func v2PDFRowHeight(pdf *gofpdf.Fpdf, req models.DocumentRequest, cols []v2Column, widths []float64, item map[string]any, bStyle models.StyleConfig) float64 {
	v2SetPDFFont(pdf, bStyle)
	lineH := v2PDFLineHeightMM(bStyle.FontSize)
	pad := 1.0
	maxLines := 1
	for i, c := range cols {
		w := widths[i] - 2*pad
		if w < 1 {
			w = 1
		}
		cellTxt := v2PDFText(v2AnyToString(item[c.Field]))
		lines := v2PDFSplitLines(pdf, cellTxt, w)
		if len(lines) > maxLines {
			maxLines = len(lines)
		}
	}
	return float64(maxLines)*lineH + 2*pad
}

func v2PDFDrawTableRow(pdf *gofpdf.Fpdf, req models.DocumentRequest, cols []v2Column, widths []float64, item map[string]any, bStyle models.StyleConfig, fillHex string) error {
	v2SetPDFFont(pdf, bStyle)
	v2SetPDFFill(pdf, fillHex)
	v2SetPDFText(pdf, bStyle.FontColor)

	lineH := v2PDFLineHeightMM(bStyle.FontSize)
	pad := 1.0

	// Pre-split all cells to get a stable row height.
	cellLines := make([][]string, len(cols))
	maxLines := 1
	for i, c := range cols {
		w := widths[i] - 2*pad
		if w < 1 {
			w = 1
		}
		cellTxt := v2PDFText(v2AnyToString(item[c.Field]))
		lines := v2PDFSplitLines(pdf, cellTxt, w)
		cellLines[i] = lines
		if len(lines) > maxLines {
			maxLines = len(lines)
		}
	}
	rowH := float64(maxLines)*lineH + 2*pad

	x0 := pdf.GetX()
	y0 := pdf.GetY()
	x := x0
	for i, c := range cols {
		w := widths[i]
		style := "F"
		if bStyle.Border {
			style = "FD"
		}
		pdf.Rect(x, y0, w, rowH, style)
		pdf.SetXY(x+pad, y0+pad)
		// Join lines with \n so MultiCell renders them.
		pdf.MultiCell(w-2*pad, lineH, strings.Join(cellLines[i], "\n"), "", v2PDFColAlign(c.Align), false)
		x += w
		pdf.SetXY(x, y0)
	}
	pdf.SetXY(x0, y0+rowH)
	return nil
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
	if s == "" {
		return ""
	}
	// gofpdf core fonts expect ISO-8859-1 encoded text. If we accidentally pass UTF-8 bytes
	// (e.g. after a failed encoding), the output becomes mojibake like "Ã¡" or "â€”".
	// Normalize common Unicode punctuation to Latin-1/ASCII and then force ISO-8859-1.
	n := v2PDFUnicodeReplacer.Replace(s)
	out, err := charmap.ISO8859_1.NewEncoder().String(n)
	if err == nil {
		return out
	}

	// Fallback: replace any remaining non-Latin-1 runes.
	var b strings.Builder
	b.Grow(len(n))
	for _, r := range n {
		switch r {
		case '\u2014', '\u2013', '\u2212':
			b.WriteByte('-')
		case '\u201C', '\u201D':
			b.WriteByte('"')
		case '\u2018', '\u2019':
			b.WriteByte('\'')
		case '\u2026':
			b.WriteString("...")
		case '\u2022':
			b.WriteByte('*')
		case '\u00A0':
			b.WriteByte(' ')
		case '\u20AC':
			b.WriteString("EUR")
		default:
			if r <= 0xFF {
				b.WriteRune(r)
			} else {
				b.WriteByte('?')
			}
		}
	}
	out, _ = charmap.ISO8859_1.NewEncoder().String(b.String())
	return out
}

var v2PDFUnicodeReplacer = strings.NewReplacer(
	"\u2014", "-", // em dash
	"\u2013", "-", // en dash
	"\u2212", "-", // minus
	"\u201C", "\"",
	"\u201D", "\"",
	"\u2018", "'",
	"\u2019", "'",
	"\u2026", "...",
	"\u2022", "*",
	"\u00A0", " ",
	"\u20AC", "EUR",
)

// ---- PDF block builder helpers ----

func v2AnyToFloat64(v any) (float64, bool) {
	if v == nil {
		return 0, false
	}
	switch vv := v.(type) {
	case float64:
		return vv, true
	case float32:
		return float64(vv), true
	case int:
		return float64(vv), true
	case int64:
		return float64(vv), true
	case json.Number:
		if f, err := vv.Float64(); err == nil {
			return f, true
		}
		if i, err := vv.Int64(); err == nil {
			return float64(i), true
		}
		return 0, false
	case string:
		if strings.TrimSpace(vv) == "" {
			return 0, false
		}
		if f, err := strconv.ParseFloat(strings.TrimSpace(vv), 64); err == nil {
			return f, true
		}
		return 0, false
	default:
		return 0, false
	}
}

func v2PDFMergeTextStyle(base models.StyleConfig, over *models.PDFTextStyleConfig) (models.StyleConfig, float64) {
	out := base
	lineSpacing := 0.0
	if over == nil {
		return out, lineSpacing
	}
	if over.FontSize > 0 {
		out.FontSize = over.FontSize
	}
	if over.Bold != nil {
		out.Bold = *over.Bold
	}
	if over.Italic != nil {
		out.Italic = *over.Italic
	}
	if over.Alignment != "" {
		out.Alignment = over.Alignment
	}
	if over.LineSpacing > 0 {
		lineSpacing = over.LineSpacing
	}
	return out, lineSpacing
}

func v2PDFUsableWidth(pdf *gofpdf.Fpdf) float64 {
	pageW, _ := pdf.GetPageSize()
	lm, _, rm, _ := pdf.GetMargins()
	return pageW - lm - rm
}

func v2PDFBottomY(pdf *gofpdf.Fpdf) float64 {
	_, pageH := pdf.GetPageSize()
	_, _, _, bm := pdf.GetMargins()
	return pageH - bm
}

func v2PDFEnsurePage(pdf *gofpdf.Fpdf, needH float64) {
	if needH <= 0 {
		return
	}
	if pdf.GetY()+needH <= v2PDFBottomY(pdf) {
		return
	}
	pdf.AddPage()
}

func v2PDFRenderBlocks(pdf *gofpdf.Fpdf, req models.DocumentRequest, hStyle, bStyle models.StyleConfig) error {
	return v2PDFRenderBlocksInternal(pdf, req, hStyle, bStyle, v2PDFRenderBlocksOptions{})
}

type v2PDFRenderBlocksOptions struct {
	SkipIndex    bool
	SectionLinks map[int]int
	CollectIndex *[]v2PDFIndexEntry
}

func v2PDFRenderBlocksInternal(pdf *gofpdf.Fpdf, req models.DocumentRequest, hStyle, bStyle models.StyleConfig, opt v2PDFRenderBlocksOptions) error {
	lm, _, _, _ := pdf.GetMargins()
	usableW := v2PDFUsableWidth(pdf)
	spacingPara := 0.0
	spacingTable := 0.0
	if req.Layout != nil && req.Layout.Spacing != nil {
		spacingPara = req.Layout.Spacing.ParagraphSpacing
		spacingTable = req.Layout.Spacing.TableSpacing
	}

	sectionIndex := 0

	for i := range req.Layout.Blocks {
		blk := req.Layout.Blocks[i]
		switch blk.Type {
		case models.PDFBlockSpacer:
			if blk.Height > 0 {
				v2PDFEnsurePage(pdf, blk.Height)
				pdf.Ln(blk.Height)
			}
			continue
		case models.PDFBlockPageBreak:
			pdf.AddPage()
			continue
		case models.PDFBlockIndex:
			if opt.SkipIndex {
				continue
			}
			return fmt.Errorf("blocks[%d]: Index must be handled by the PDF index renderer", i)
		case models.PDFBlockSectionTitle:
			sectionIndex++
			// Default title style.
			base := bStyle
			base.FontSize = 16
			base.Bold = true
			style, lineSpacing := v2PDFMergeTextStyle(base, blk.Style)
			lh := v2PDFLineHeightMM(style.FontSize)
			if lineSpacing > 0 {
				lh *= lineSpacing
			}
			v2SetPDFFont(pdf, style)
			pdf.SetX(lm)
			lines := v2PDFSplitLines(pdf, v2PDFText(blk.Content), usableW)
			v2PDFEnsurePage(pdf, float64(len(lines))*lh)
			if opt.CollectIndex != nil {
				*opt.CollectIndex = append(*opt.CollectIndex, v2PDFIndexEntry{Title: blk.Content, PageNo: pdf.PageNo(), SectionIndex: sectionIndex})
			}
			if opt.SectionLinks != nil {
				if linkID, ok := opt.SectionLinks[sectionIndex]; ok {
					pdf.SetLink(linkID, pdf.GetY(), pdf.PageNo())
				}
			}
			pdf.MultiCell(usableW, lh, v2PDFText(blk.Content), "", v2PDFAlign(style.Alignment), false)
			if spacingPara > 0 {
				pdf.Ln(spacingPara)
			}
			continue
		case models.PDFBlockText:
			style, lineSpacing := v2PDFMergeTextStyle(bStyle, blk.Style)
			lh := v2PDFLineHeightMM(style.FontSize)
			if lineSpacing > 0 {
				lh *= lineSpacing
			}
			v2SetPDFFont(pdf, style)
			v2SetPDFText(pdf, style.FontColor)
			pdf.SetX(lm)
			pdf.MultiCell(usableW, lh, v2PDFText(blk.Content), "", v2PDFAlign(style.Alignment), false)
			if spacingPara > 0 {
				pdf.Ln(spacingPara)
			}
			continue
		case models.PDFBlockTable:
			if err := v2PDFRenderTableBlock(pdf, req, blk, hStyle, bStyle); err != nil {
				return fmt.Errorf("blocks[%d] table: %w", i, err)
			}
			if spacingTable > 0 {
				pdf.Ln(spacingTable)
			}
			continue
		case models.PDFBlockChart:
			if err := v2PDFRenderChartBlock(pdf, req, blk, bStyle); err != nil {
				return fmt.Errorf("blocks[%d] chart: %w", i, err)
			}
			if spacingTable > 0 {
				pdf.Ln(spacingTable)
			}
			continue
		case models.PDFBlockImage:
			if err := v2PDFRenderImageBlock(pdf, req, blk); err != nil {
				return fmt.Errorf("blocks[%d] image: %w", i, err)
			}
			if spacingTable > 0 {
				pdf.Ln(spacingTable)
			}
			continue
		default:
			return fmt.Errorf("blocks[%d]: unsupported block type", i)
		}
	}
	return nil
}

func v2PDFRenderTableBlock(pdf *gofpdf.Fpdf, req models.DocumentRequest, blk models.PDFBlockConfig, hStyle, bStyle models.StyleConfig) error {
	data := req.Data.Get(blk.DataSource)
	if data.IsEmpty() {
		name := strings.TrimSpace(blk.DataSource)
		if name == "" {
			name = "default"
		}
		return fmt.Errorf("dataSource '%s' is empty or missing", name)
	}

	cols := make([]v2Column, 0, len(blk.Columns))
	order := make([]string, 0, len(blk.Columns))
	for _, c := range blk.Columns {
		t := strings.TrimSpace(c.Title)
		if t == "" {
			t = c.Field
		}
		cols = append(cols, v2Column{Field: c.Field, Title: t, Width: 1, Align: models.ColAlignLeft, VAlign: models.VAlignMiddle})
		order = append(order, c.Field)
	}
	// If dataset has no stable order (e.g., was sources-only), apply requested column order.
	if len(order) > 0 {
		data.Order = order
	}

	usableW := v2PDFUsableWidth(pdf)
	widths := v2ComputePDFWidths(cols, usableW)

	renderHeader := func() {
		v2PDFDrawTableHeader(pdf, req, cols, widths, hStyle)
	}

	// Ensure there's room for the header row before starting.
	pad := 1.0
	needHeaderH := v2PDFLineHeightMM(hStyle.FontSize) + 2*pad
	v2PDFEnsurePage(pdf, needHeaderH)
	renderHeader()

	newPageWithHeader := func() {
		pdf.AddPage()
		renderHeader()
	}

	ensureRow := func(needH float64) {
		if pdf.GetY()+needH <= v2PDFBottomY(pdf) {
			return
		}
		newPageWithHeader()
	}

	rowsPerPage := 0
	if req.Layout != nil && req.Layout.PageBreak != nil && req.Layout.PageBreak.Enabled {
		rowsPerPage = req.Layout.PageBreak.RowsPerPage
	}
	printed := 0
	for r, item := range data.Items {
		if rowsPerPage > 0 && printed > 0 && printed%rowsPerPage == 0 {
			newPageWithHeader()
		}
		fillColor := bStyle.Background
		if bStyle.ZebraStripe {
			if r%2 == 0 {
				fillColor = bStyle.ZebraColorEven
			} else {
				fillColor = bStyle.ZebraColorOdd
			}
		}
		rowH := v2PDFRowHeight(pdf, req, cols, widths, item, bStyle)
		ensureRow(rowH)
		if err := v2PDFDrawTableRow(pdf, req, cols, widths, item, bStyle, fillColor); err != nil {
			return err
		}
		printed++
	}
	return nil
}

func v2PDFRenderChartBlock(pdf *gofpdf.Fpdf, req models.DocumentRequest, blk models.PDFBlockConfig, bStyle models.StyleConfig) error {
	data := req.Data.Get(blk.DataSource)
	if data.IsEmpty() {
		name := strings.TrimSpace(blk.DataSource)
		if name == "" {
			name = "default"
		}
		return fmt.Errorf("dataSource '%s' is empty or missing", name)
	}

	lm, _, _, _ := pdf.GetMargins()
	usableW := v2PDFUsableWidth(pdf)
	width := usableW
	if blk.Width > 0 && blk.Width < width {
		width = blk.Width
	}
	height := 70.0
	if blk.Height > 0 {
		height = blk.Height
	}
	// Title area height.
	titleH := 0.0
	if strings.TrimSpace(blk.Title) != "" {
		titleH = v2PDFLineHeightMM(12) + 2
	}
	// Ensure block fits; charts are indivisible.
	v2PDFEnsurePage(pdf, titleH+height+6)

	// Render title.
	if strings.TrimSpace(blk.Title) != "" {
		titleStyle := bStyle
		titleStyle.FontSize = 12
		titleStyle.Bold = true
		v2SetPDFFont(pdf, titleStyle)
		pdf.SetX(lm)
		pdf.MultiCell(width, v2PDFLineHeightMM(12), v2PDFText(blk.Title), "", "L", false)
		pdf.Ln(2)
	}

	// Plot box.
	x := lm
	y := pdf.GetY()
	if height > v2PDFBottomY(pdf)-y {
		height = v2PDFBottomY(pdf) - y
		if height < 20 {
			height = 20
		}
	}
	if width > usableW {
		width = usableW
	}

	seriesCats, seriesVals := v2PDFChartSeries(data, blk.CategoryField, blk.ValueField)
	if len(seriesCats) == 0 {
		return fmt.Errorf("no data points found for categoryField/valueField")
	}
	v2PDFDrawChart(pdf, blk.ChartType, x, y, width, height, seriesCats, seriesVals, bStyle)
	// Move cursor below chart.
	pdf.SetXY(lm, y+height)
	return nil
}

func v2PDFChartSeries(data models.DynamicData, categoryField, valueField string) ([]string, []float64) {
	cats := make([]string, 0, len(data.Items))
	vals := make([]float64, 0, len(data.Items))
	for _, it := range data.Items {
		cat := v2AnyToString(it[categoryField])
		v, ok := v2AnyToFloat64(it[valueField])
		if !ok {
			continue
		}
		cats = append(cats, cat)
		vals = append(vals, v)
	}
	return cats, vals
}

func v2PDFDrawChart(pdf *gofpdf.Fpdf, chartType models.ChartTypeEnum, x, y, w, h float64, cats []string, vals []float64, style models.StyleConfig) {
	// Chart border
	pdf.SetDrawColor(30, 30, 30)
	pdf.Rect(x, y, w, h, "D")

	palette := []string{"#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948", "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC"}

	switch chartType {
	case models.ChartPie:
		v2PDFDrawPieChart(pdf, x, y, w, h, cats, vals, palette)
	case models.ChartLine:
		v2PDFDrawLineChart(pdf, x, y, w, h, cats, vals, false, palette)
	case models.ChartArea:
		v2PDFDrawLineChart(pdf, x, y, w, h, cats, vals, true, palette)
	case models.ChartBar:
		v2PDFDrawBarChart(pdf, x, y, w, h, cats, vals, true, palette)
	case models.ChartColumn:
		v2PDFDrawBarChart(pdf, x, y, w, h, cats, vals, false, palette)
	default:
		// Fallback: draw nothing.
	}

	_ = style
}

func v2PDFDrawBarChart(pdf *gofpdf.Fpdf, x, y, w, h float64, cats []string, vals []float64, horizontal bool, palette []string) {
	n := len(vals)
	if n == 0 {
		return
	}
	maxVal := 0.0
	for _, v := range vals {
		if v > maxVal {
			maxVal = v
		}
	}
	if maxVal <= 0 {
		maxVal = 1
	}

	labelFontSize := 8
	pdf.SetFont("Helvetica", "", float64(labelFontSize))

	if horizontal {
		leftPad := 30.0
		plotW := w - leftPad - 6
		plotH := h - 6
		if plotW < 10 {
			plotW = 10
		}
		slotH := plotH / float64(n)
		barH := slotH * 0.6
		if barH < 2 {
			barH = 2
		}
		for i := 0; i < n; i++ {
			v := vals[i]
			barW := (v / maxVal) * plotW
			bx := x + leftPad
			by := y + 3 + float64(i)*slotH + (slotH-barH)/2
			v2SetPDFFill(pdf, palette[i%len(palette)])
			pdf.Rect(bx, by, barW, barH, "F")
			// Label
			lab := cats[i]
			if utf8.RuneCountInString(lab) > 18 {
				lab = string([]rune(lab)[:18]) + "…"
			}
			pdf.Text(x+2, by+barH, v2PDFText(lab))
		}
		return
	}

	bottomPad := 12.0
	plotH := h - bottomPad - 4
	plotW := w - 6
	if plotH < 10 {
		plotH = 10
	}
	slotW := plotW / float64(n)
	barW := slotW * 0.6
	if barW < 2 {
		barW = 2
	}
	for i := 0; i < n; i++ {
		v := vals[i]
		barH := (v / maxVal) * plotH
		bx := x + 3 + float64(i)*slotW + (slotW-barW)/2
		by := y + 3 + (plotH - barH)
		v2SetPDFFill(pdf, palette[i%len(palette)])
		pdf.Rect(bx, by, barW, barH, "F")
		lab := cats[i]
		if utf8.RuneCountInString(lab) > 10 {
			lab = string([]rune(lab)[:10]) + "…"
		}
		pdf.Text(bx, y+h-3, v2PDFText(lab))
	}
}

func v2PDFDrawLineChart(pdf *gofpdf.Fpdf, x, y, w, h float64, cats []string, vals []float64, fill bool, palette []string) {
	n := len(vals)
	if n == 0 {
		return
	}
	maxVal := 0.0
	minVal := vals[0]
	for _, v := range vals {
		if v > maxVal {
			maxVal = v
		}
		if v < minVal {
			minVal = v
		}
	}
	if maxVal == minVal {
		maxVal = minVal + 1
	}
	leftPad := 6.0
	bottomPad := 10.0
	plotW := w - leftPad - 4
	plotH := h - bottomPad - 4
	if plotW < 10 {
		plotW = 10
	}
	if plotH < 10 {
		plotH = 10
	}
	x0 := x + leftPad
	y0 := y + 3 + plotH

	stepX := 0.0
	if n > 1 {
		stepX = plotW / float64(n-1)
	}
	points := make([]gofpdf.PointType, 0, n)
	for i := 0; i < n; i++ {
		norm := (vals[i] - minVal) / (maxVal - minVal)
		px := x0 + float64(i)*stepX
		py := y0 - norm*plotH
		points = append(points, gofpdf.PointType{X: px, Y: py})
	}

	// Optional area fill
	if fill {
		poly := make([]gofpdf.PointType, 0, len(points)+2)
		poly = append(poly, gofpdf.PointType{X: points[0].X, Y: y0})
		poly = append(poly, points...)
		poly = append(poly, gofpdf.PointType{X: points[len(points)-1].X, Y: y0})
		v2SetPDFFill(pdf, palette[0])
		pdf.SetAlpha(0.20, "Normal")
		pdf.Polygon(poly, "F")
		pdf.SetAlpha(1.0, "Normal")
	}

	// Line stroke
	pdf.SetDrawColor(10, 10, 10)
	for i := 0; i+1 < len(points); i++ {
		pdf.Line(points[i].X, points[i].Y, points[i+1].X, points[i+1].Y)
	}
}

func v2PDFDrawPieChart(pdf *gofpdf.Fpdf, x, y, w, h float64, cats []string, vals []float64, palette []string) {
	total := 0.0
	for _, v := range vals {
		if v > 0 {
			total += v
		}
	}
	if total <= 0 {
		return
	}
	r := math.Min(w, h)/2 - 6
	if r < 10 {
		r = 10
	}
	cx := x + w/2
	cy := y + h/2
	start := -math.Pi / 2
	for i, v := range vals {
		if v <= 0 {
			continue
		}
		frac := v / total
		end := start + frac*2*math.Pi
		pts := make([]gofpdf.PointType, 0, 64)
		pts = append(pts, gofpdf.PointType{X: cx, Y: cy})
		step := 6.0 * (math.Pi / 180)
		for a := start; a <= end; a += step {
			pts = append(pts, gofpdf.PointType{X: cx + r*math.Cos(a), Y: cy + r*math.Sin(a)})
		}
		pts = append(pts, gofpdf.PointType{X: cx + r*math.Cos(end), Y: cy + r*math.Sin(end)})
		v2SetPDFFill(pdf, palette[i%len(palette)])
		pdf.Polygon(pts, "F")
		start = end
	}
	// Outline circle
	pdf.SetDrawColor(30, 30, 30)
	pdf.Ellipse(cx, cy, r, r, 0, "D")
	_ = cats
}

func v2PDFRenderImageBlock(pdf *gofpdf.Fpdf, req models.DocumentRequest, blk models.PDFBlockConfig) error {
	imgBytes, ext, err := v2DecodeImageBytes(blk.Data)
	if err != nil {
		return err
	}
	opt := gofpdf.ImageOptions{ImageType: strings.TrimPrefix(strings.ToUpper(ext), "."), ReadDpi: true}

	ic, format, err := image.DecodeConfig(bytes.NewReader(imgBytes))
	if err != nil {
		return err
	}
	if format != "" {
		opt.ImageType = strings.ToUpper(format)
	}
	name := fmt.Sprintf("img_%d", time.Now().UnixNano())
	pdf.RegisterImageOptionsReader(name, opt, bytes.NewReader(imgBytes))

	usableW := v2PDFUsableWidth(pdf)
	pageW, _ := pdf.GetPageSize()
	lm, _, rm, _ := pdf.GetMargins()
	_ = pageW
	_ = rm

	// Determine dimensions (mm). Defaults are conservative.
	w := blk.Width
	h := blk.Height
	if w <= 0 && h <= 0 {
		h = 24
	}
	ratio := 1.0
	if ic.Height > 0 {
		ratio = float64(ic.Width) / float64(ic.Height)
	}
	if w <= 0 && h > 0 {
		w = h * ratio
	}
	if h <= 0 && w > 0 {
		h = w / ratio
	}
	// Indivisible block: ensure fits before placement.
	v2PDFEnsurePage(pdf, h)

	// Clamp to available space.
	if w > usableW {
		scale := usableW / w
		w *= scale
		h *= scale
	}
	availH := v2PDFBottomY(pdf) - pdf.GetY()
	if h > availH {
		scale := availH / h
		w *= scale
		h *= scale
	}

	x := lm
	switch blk.Alignment {
	case models.AlignCenter:
		x = lm + (usableW-w)/2
	case models.AlignRight:
		x = lm + (usableW - w)
	}
	y := pdf.GetY()
	pdf.ImageOptions(name, x, y, w, h, false, opt, 0, "")
	pdf.SetXY(lm, y+h)
	return nil
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

func v2ExcelEnsureOnlySheets(f *excelize.File, targetSheets []models.SheetConfig) error {
	if f == nil {
		return nil
	}
	if len(targetSheets) == 0 {
		return nil
	}

	// Normalize requested sheet names exactly as v2RenderExcel does.
	sheetNames := make([]string, 0, len(targetSheets))
	for i, sh := range targetSheets {
		name := strings.TrimSpace(sh.Name)
		if name == "" {
			name = fmt.Sprintf("Sheet%d", i+1)
		}
		sheetNames = append(sheetNames, name)
	}

	desired := map[string]bool{}
	for _, n := range sheetNames {
		desired[strings.ToLower(n)] = true
	}

	// Ensure desired sheets exist.
	for i, name := range sheetNames {
		idx, err := f.GetSheetIndex(name)
		if err != nil || idx < 0 {
			idx, err = f.NewSheet(name)
			if err != nil {
				return err
			}
			if i == 0 {
				f.SetActiveSheet(idx)
			}
		} else if i == 0 {
			f.SetActiveSheet(idx)
		}
	}

	// Remove any sheets not declared in layout.sheets (including the default Sheet1).
	for _, existing := range f.GetSheetList() {
		if !desired[strings.ToLower(existing)] {
			_ = f.DeleteSheet(existing)
		}
	}
	return nil
}

// ---- Word (docx zip) helpers ----

type v2WordImagePart struct {
	RelID string
	Ext   string
	Path  string
	Bytes []byte
}

func writeZipString(zw *zip.Writer, name string, content string) {
	w, _ := zw.Create(name)
	_, _ = w.Write([]byte(content))
}

func v2WordContentTypesXML(includeFooter bool, images []v2WordImagePart) string {
	footerOverride := ""
	if includeFooter {
		footerOverride = "\n  <Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>"
	}

	seen := map[string]bool{}
	imgDefaults := make([]string, 0, 2)
	for _, img := range images {
		ext := strings.ToLower(strings.TrimPrefix(img.Ext, "."))
		if ext == "" {
			continue
		}
		if seen[ext] {
			continue
		}
		seen[ext] = true
		ct := "image/png"
		if ext == "jpg" || ext == "jpeg" {
			ct = "image/jpeg"
		}
		imgDefaults = append(imgDefaults, fmt.Sprintf("\n  <Default Extension=\"%s\" ContentType=\"%s\"/>", ext, ct))
	}

	return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>` + strings.Join(imgDefaults, "") + `
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

func v2WordDocumentRelsXML(includeFooter bool, images []v2WordImagePart) string {
	parts := []string{
		`  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
	}
	if includeFooter {
		parts = append(parts, `  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
	}
	for _, img := range images {
		target := strings.TrimPrefix(img.Path, "word/")
		parts = append(parts, fmt.Sprintf(`  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>`, img.RelID, target))
	}
	return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" + strings.Join(parts, "\n") + "\n</Relationships>"
}

func v2WordStylesXML() string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
		<w:name w:val="Normal"/>
	</w:style>
	<w:style w:type="paragraph" w:styleId="Heading1">
		<w:name w:val="heading 1"/>
		<w:basedOn w:val="Normal"/>
		<w:uiPriority w:val="9"/>
		<w:qFormat/>
		<w:pPr>
			<w:spacing w:before="240" w:after="120"/>
		</w:pPr>
		<w:rPr>
			<w:b/>
			<w:sz w:val="32"/>
		</w:rPr>
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

func v2WordDocumentXML(req models.DocumentRequest, bodyXML string, includeFooter bool) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
`)
	sb.WriteString(bodyXML)
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

func v2WordBlocksBodyXML(req models.DocumentRequest, startImageIndex, startDocPrID int) (string, []v2WordImagePart, error) {
	if req.Layout == nil || len(req.Layout.Blocks) == 0 {
		return "", nil, nil
	}
	images := make([]v2WordImagePart, 0, 4)
	nextImage := maxInt(1, startImageIndex)
	nextDocPrID := maxInt(1, startDocPrID)

	rowsPerPage := 0
	if req.Layout.PageBreak != nil && req.Layout.PageBreak.Enabled {
		rowsPerPage = req.Layout.PageBreak.RowsPerPage
	}

	var sb strings.Builder
	for i := range req.Layout.Blocks {
		blk := req.Layout.Blocks[i]
		switch blk.Type {
		case models.PDFBlockIndex:
			title := strings.TrimSpace(blk.Content)
			if title == "" {
				title = "Índice"
			}
			sb.WriteString(v2WordBoldTextParagraphXML(title))
			sb.WriteString(v2WordTOCFieldXML())
			sb.WriteString(v2WordPageBreakParagraphXML())
		case models.PDFBlockSectionTitle:
			sb.WriteString(v2WordHeading1ParagraphXML(blk.Content))
		case models.PDFBlockText:
			sb.WriteString(v2WordTextParagraphXML(blk.Content))
		case models.PDFBlockSpacer:
			sb.WriteString(v2WordSpacerParagraphXML(blk.Height))
		case models.PDFBlockPageBreak:
			sb.WriteString(v2WordPageBreakParagraphXML())
		case models.PDFBlockTable:
			data := req.Data.Get(blk.DataSource)
			if data.IsEmpty() {
				name := strings.TrimSpace(blk.DataSource)
				if name == "" {
					name = "default"
				}
				return "", nil, fmt.Errorf("blocks[%d] table: dataSource '%s' is empty or missing", i, name)
			}
			cols := make([]v2Column, 0, len(blk.Columns))
			for _, c := range blk.Columns {
				t := strings.TrimSpace(c.Title)
				if t == "" {
					t = c.Field
				}
				cols = append(cols, v2Column{Field: c.Field, Title: t, Width: 1, Align: models.ColAlignLeft, VAlign: models.VAlignMiddle})
			}
			sb.WriteString(v2WordTableWithPageBreaksXML(req, cols, data.Items, rowsPerPage))
		case models.PDFBlockImage:
			imgBytes, ext, err := v2DecodeImageBytes(blk.Data)
			if err != nil {
				return "", nil, fmt.Errorf("blocks[%d] image: %w", i, err)
			}
			relID := fmt.Sprintf("rIdImage%d", nextImage)
			path := fmt.Sprintf("word/media/image%d%s", nextImage, ext)
			images = append(images, v2WordImagePart{RelID: relID, Ext: ext, Path: path, Bytes: imgBytes})

			jc := v2WordJcFromAlign(blk.Alignment)
			cx, cy := v2WordComputeBlockImageBoxEMU(imgBytes, blk.Width, blk.Height, 120, 60)
			sb.WriteString(v2WordInlineImageParagraphXML(jc, relID, nextDocPrID, fmt.Sprintf("Image%d", nextImage), cx, cy))
			nextImage++
			nextDocPrID++
		case models.PDFBlockChart:
			data := req.Data.Get(blk.DataSource)
			if data.IsEmpty() {
				name := strings.TrimSpace(blk.DataSource)
				if name == "" {
					name = "default"
				}
				return "", nil, fmt.Errorf("blocks[%d] chart: dataSource '%s' is empty or missing", i, name)
			}
			cats, vals := v2PDFChartSeries(data, blk.CategoryField, blk.ValueField)
			if len(cats) == 0 {
				return "", nil, fmt.Errorf("blocks[%d] chart: no data points found for categoryField/valueField", i)
			}
			if strings.TrimSpace(blk.Title) != "" {
				sb.WriteString(v2WordBoldTextParagraphXML(blk.Title))
			}
			wpx, hpx := v2WordChartPixelSize(blk.Width, blk.Height)
			pngBytes, err := v2WordRenderChartPNG(blk.ChartType, cats, vals, wpx, hpx)
			if err != nil {
				return "", nil, fmt.Errorf("blocks[%d] chart: %w", i, err)
			}
			relID := fmt.Sprintf("rIdImage%d", nextImage)
			path := fmt.Sprintf("word/media/image%d.png", nextImage)
			images = append(images, v2WordImagePart{RelID: relID, Ext: ".png", Path: path, Bytes: pngBytes})
			jc := v2WordJcFromAlign(models.AlignCenter)
			cx, cy := v2WordComputeChartBoxEMU(blk.Width, blk.Height)
			sb.WriteString(v2WordInlineImageParagraphXML(jc, relID, nextDocPrID, fmt.Sprintf("Chart%d", nextImage), cx, cy))
			nextImage++
			nextDocPrID++
		default:
			return "", nil, fmt.Errorf("blocks[%d]: unsupported block type", i)
		}
	}
	return sb.String(), images, nil
}

func v2WordTextParagraphXML(text string) string {
	return `<w:p><w:r><w:t xml:space="preserve">` + v2EscapeXML(text) + `</w:t></w:r></w:p>`
}

func v2WordBoldTextParagraphXML(text string) string {
	return `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">` + v2EscapeXML(text) + `</w:t></w:r></w:p>`
}

func v2WordHeading1ParagraphXML(text string) string {
	return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">` + v2EscapeXML(text) + `</w:t></w:r></w:p>`
}

func v2WordSpacerParagraphXML(heightMM float64) string {
	if heightMM <= 0 {
		return `<w:p/>`
	}
	afterTw := v2MMToTwips(heightMM)
	return fmt.Sprintf(`<w:p><w:pPr><w:spacing w:after="%d"/></w:pPr></w:p>`, afterTw)
}

func v2WordPageBreakParagraphXML() string {
	return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`
}

func v2WordTOCFieldXML() string {
	// Word will populate the Table of Contents when fields are updated (usually on open/print).
	// \h enables hyperlinks; page numbers are included by default.
	return `<w:p><w:fldSimple w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u"><w:r><w:t xml:space="preserve"> </w:t></w:r></w:fldSimple></w:p>`
}

func v2WordInlineImageParagraphXML(jc, relID string, docPrID int, name string, cxEMU, cyEMU int64) string {
	if strings.TrimSpace(jc) == "" {
		jc = "center"
	}
	if docPrID <= 0 {
		docPrID = 1
	}
	if strings.TrimSpace(name) == "" {
		name = "Image"
	}
	return fmt.Sprintf(`<w:p>
	<w:pPr><w:jc w:val="%s"/></w:pPr>
	<w:r>
		<w:drawing>
			<wp:inline distT="0" distB="0" distL="0" distR="0">
				<wp:extent cx="%d" cy="%d"/>
				<wp:docPr id="%d" name="%s"/>
				<a:graphic>
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic>
							<pic:nvPicPr>
								<pic:cNvPr id="0" name="%s"/>
								<pic:cNvPicPr/>
							</pic:nvPicPr>
							<pic:blipFill>
								<a:blip r:embed="%s"/>
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
</w:p>`, jc, cxEMU, cyEMU, docPrID, v2EscapeXML(name), v2EscapeXML(name), relID, cxEMU, cyEMU)
}

func v2WordJcFromAlign(a models.AlignmentEnum) string {
	switch a {
	case models.AlignLeft:
		return "left"
	case models.AlignRight:
		return "right"
	default:
		return "center"
	}
}

func v2WordComputeBlockImageBoxEMU(imgBytes []byte, widthMM, heightMM float64, defaultWmm, defaultHmm float64) (cxEMU, cyEMU int64) {
	// Defaults
	w := widthMM
	h := heightMM
	if w <= 0 {
		w = defaultWmm
	}
	if h <= 0 {
		h = defaultHmm
	}

	// Preserve aspect ratio when exactly one dimension is specified.
	if (widthMM > 0 && heightMM <= 0) || (heightMM > 0 && widthMM <= 0) {
		if ic, _, err := image.DecodeConfig(bytes.NewReader(imgBytes)); err == nil {
			ratio := float64(ic.Width) / float64(maxInt(1, ic.Height))
			if widthMM > 0 && heightMM <= 0 {
				h = widthMM / ratio
			}
			if heightMM > 0 && widthMM <= 0 {
				w = heightMM * ratio
			}
		}
	}

	return v2MMToEMU(w), v2MMToEMU(h)
}

func v2WordComputeChartBoxEMU(widthMM, heightMM float64) (cxEMU, cyEMU int64) {
	w := widthMM
	h := heightMM
	if w <= 0 {
		w = 160
	}
	if h <= 0 {
		h = 90
	}
	return v2MMToEMU(w), v2MMToEMU(h)
}

func v2WordChartPixelSize(widthMM, heightMM float64) (wpx, hpx int) {
	// 96 DPI approximation.
	mmToPx := func(mm float64) int {
		if mm <= 0 {
			return 0
		}
		return int(math.Round(mm / 25.4 * 96.0))
	}
	wpx = mmToPx(widthMM)
	hpx = mmToPx(heightMM)
	if wpx <= 0 {
		wpx = 900
	}
	if hpx <= 0 {
		hpx = 480
	}
	if wpx < 240 {
		wpx = 240
	}
	if hpx < 140 {
		hpx = 140
	}
	if wpx > 2000 {
		wpx = 2000
	}
	if hpx > 1200 {
		hpx = 1200
	}
	return
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

func v2WordTableWithPageBreaksXML(req models.DocumentRequest, cols []v2Column, items []map[string]any, rowsPerPage int) string {
	if rowsPerPage <= 0 {
		return v2WordTableXML(req, cols, items)
	}
	var sb strings.Builder
	for i := 0; i < len(items); i += rowsPerPage {
		end := i + rowsPerPage
		if end > len(items) {
			end = len(items)
		}
		if i > 0 {
			sb.WriteString(v2WordPageBreakParagraphXML())
		}
		sb.WriteString(v2WordTableXML(req, cols, items[i:end]))
	}
	return sb.String()
}

func v2WordTableXML(req models.DocumentRequest, cols []v2Column, items []map[string]any) string {
	var sb strings.Builder
	pageOrientation := models.PagePortrait
	if req.Layout != nil && req.Layout.PageOrientation != "" {
		pageOrientation = req.Layout.PageOrientation
	}
	pageWtw, _, _ := v2WordPageSizeTwips(pageOrientation)

	// Defaults
	marginLeftMM, marginRightMM := 25.4, 25.4
	if req.Layout != nil && req.Layout.PageMargin != nil {
		marginLeftMM = req.Layout.PageMargin.Left
		marginRightMM = req.Layout.PageMargin.Right
	}
	marginLeftTw := v2MMToTwips(marginLeftMM)
	marginRightTw := v2MMToTwips(marginRightMM)
	contentWtw := pageWtw - marginLeftTw - marginRightTw
	if contentWtw < 0 {
		contentWtw = 0
	}

	ignoreMargins := false
	if req.Layout != nil {
		ignoreMargins = req.Layout.WordIgnorePageMargins
		// Back-compat: UsePageContentBounds=false meant "ignore margins".
		if req.Layout.UsePageContentBounds != nil && !*req.Layout.UsePageContentBounds {
			ignoreMargins = true
		}
	}
	centerContent := true
	if req.Layout != nil && req.Layout.WordCenterContent != nil {
		centerContent = *req.Layout.WordCenterContent
	}

	tableWtw := contentWtw
	if ignoreMargins {
		tableWtw = pageWtw
	}

	jc := ""
	if centerContent {
		jc = `<w:jc w:val="center"/>`
	}

	tblW := fmt.Sprintf(`<w:tblW w:type="dxa" w:w="%d"/>`, tableWtw)
	sb.WriteString(`<w:tbl>
	  <w:tblPr>` + tblW + jc + `</w:tblPr>
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

func v2WordRenderChartPNG(chartType models.ChartTypeEnum, cats []string, vals []float64, widthPx, heightPx int) ([]byte, error) {
	if widthPx <= 0 {
		widthPx = 900
	}
	if heightPx <= 0 {
		heightPx = 480
	}
	img := image.NewRGBA(image.Rect(0, 0, widthPx, heightPx))
	// White background
	draw.Draw(img, img.Bounds(), &image.Uniform{C: color.RGBA{R: 255, G: 255, B: 255, A: 255}}, image.Point{}, draw.Src)

	// Border
	v2DrawRect(img, 0, 0, widthPx-1, heightPx-1, color.RGBA{R: 30, G: 30, B: 30, A: 255})

	palette := []color.RGBA{
		v2HexToRGBA("#4E79A7"), v2HexToRGBA("#F28E2B"), v2HexToRGBA("#E15759"), v2HexToRGBA("#76B7B2"), v2HexToRGBA("#59A14F"),
		v2HexToRGBA("#EDC948"), v2HexToRGBA("#B07AA1"), v2HexToRGBA("#FF9DA7"), v2HexToRGBA("#9C755F"), v2HexToRGBA("#BAB0AC"),
	}

	switch chartType {
	case models.ChartPie:
		v2DrawPieChart(img, vals, palette)
	case models.ChartLine:
		v2DrawLineOrAreaChart(img, vals, palette[0], false)
	case models.ChartArea:
		v2DrawLineOrAreaChart(img, vals, palette[0], true)
	case models.ChartBar:
		v2DrawBarChart(img, vals, palette, true)
	case models.ChartColumn:
		v2DrawBarChart(img, vals, palette, false)
	default:
		// Fallback: treat as column.
		v2DrawBarChart(img, vals, palette, false)
	}

	_ = cats

	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func v2HexToRGBA(hex string) color.RGBA {
	s := v2NormalizeHexColor(hex)
	if len(s) != 7 {
		return color.RGBA{R: 0, G: 0, B: 0, A: 255}
	}
	parse := func(i int) uint8 {
		v, err := strconv.ParseUint(s[i:i+2], 16, 8)
		if err != nil {
			return 0
		}
		return uint8(v)
	}
	return color.RGBA{R: parse(1), G: parse(3), B: parse(5), A: 255}
}

func v2DrawRect(img *image.RGBA, x0, y0, x1, y1 int, col color.RGBA) {
	for x := x0; x <= x1; x++ {
		img.SetRGBA(x, y0, col)
		img.SetRGBA(x, y1, col)
	}
	for y := y0; y <= y1; y++ {
		img.SetRGBA(x0, y, col)
		img.SetRGBA(x1, y, col)
	}
}

func v2FillRect(img *image.RGBA, x0, y0, x1, y1 int, col color.RGBA) {
	if x0 > x1 {
		x0, x1 = x1, x0
	}
	if y0 > y1 {
		y0, y1 = y1, y0
	}
	b := img.Bounds()
	if x0 < b.Min.X {
		x0 = b.Min.X
	}
	if y0 < b.Min.Y {
		y0 = b.Min.Y
	}
	if x1 >= b.Max.X {
		x1 = b.Max.X - 1
	}
	if y1 >= b.Max.Y {
		y1 = b.Max.Y - 1
	}
	for y := y0; y <= y1; y++ {
		for x := x0; x <= x1; x++ {
			img.SetRGBA(x, y, col)
		}
	}
}

func v2DrawLine(img *image.RGBA, x0, y0, x1, y1 int, col color.RGBA) {
	dx := int(math.Abs(float64(x1 - x0)))
	sx := -1
	if x0 < x1 {
		sx = 1
	}
	dy := -int(math.Abs(float64(y1 - y0)))
	sy := -1
	if y0 < y1 {
		sy = 1
	}
	err := dx + dy
	for {
		img.SetRGBA(x0, y0, col)
		if x0 == x1 && y0 == y1 {
			break
		}
		e2 := 2 * err
		if e2 >= dy {
			err += dy
			x0 += sx
		}
		if e2 <= dx {
			err += dx
			y0 += sy
		}
	}
}

func v2DrawBarChart(img *image.RGBA, vals []float64, palette []color.RGBA, horizontal bool) {
	n := len(vals)
	if n == 0 {
		return
	}
	maxVal := 0.0
	for _, v := range vals {
		if v > maxVal {
			maxVal = v
		}
	}
	if maxVal <= 0 {
		maxVal = 1
	}

	b := img.Bounds()
	padL, padR, padT, padB := 20, 20, 20, 20
	plotW := (b.Dx() - padL - padR)
	plotH := (b.Dy() - padT - padB)
	if plotW < 10 {
		plotW = 10
	}
	if plotH < 10 {
		plotH = 10
	}

	if horizontal {
		slotH := float64(plotH) / float64(n)
		barH := int(math.Max(2, slotH*0.6))
		for i := 0; i < n; i++ {
			v := vals[i]
			barW := int(math.Round((v / maxVal) * float64(plotW)))
			x0 := padL
			yc := padT + int(float64(i)*slotH+slotH/2)
			y0 := yc - barH/2
			v2FillRect(img, x0, y0, x0+barW, y0+barH, palette[i%len(palette)])
		}
		return
	}

	slotW := float64(plotW) / float64(n)
	barW := int(math.Max(2, slotW*0.6))
	for i := 0; i < n; i++ {
		v := vals[i]
		barH := int(math.Round((v / maxVal) * float64(plotH)))
		xc := padL + int(float64(i)*slotW+slotW/2)
		x0 := xc - barW/2
		y0 := padT + (plotH - barH)
		v2FillRect(img, x0, y0, x0+barW, y0+barH, palette[i%len(palette)])
	}
}

func v2DrawLineOrAreaChart(img *image.RGBA, vals []float64, col color.RGBA, fill bool) {
	n := len(vals)
	if n == 0 {
		return
	}
	maxVal := vals[0]
	minVal := vals[0]
	for _, v := range vals {
		if v > maxVal {
			maxVal = v
		}
		if v < minVal {
			minVal = v
		}
	}
	if maxVal == minVal {
		maxVal = minVal + 1
	}

	b := img.Bounds()
	padL, padR, padT, padB := 20, 20, 20, 30
	plotW := (b.Dx() - padL - padR)
	plotH := (b.Dy() - padT - padB)
	if plotW < 10 {
		plotW = 10
	}
	if plotH < 10 {
		plotH = 10
	}

	points := make([]image.Point, 0, n)
	stepX := 0.0
	if n > 1 {
		stepX = float64(plotW) / float64(n-1)
	}
	for i := 0; i < n; i++ {
		norm := (vals[i] - minVal) / (maxVal - minVal)
		x := padL + int(math.Round(float64(i)*stepX))
		y := padT + int(math.Round(float64(plotH)*(1-norm)))
		points = append(points, image.Point{X: x, Y: y})
	}

	if fill {
		fillCol := col
		fillCol.A = 50
		// Simple vertical-fill under polyline.
		for i := 0; i+1 < len(points); i++ {
			p0 := points[i]
			p1 := points[i+1]
			x0, x1 := p0.X, p1.X
			if x0 > x1 {
				x0, x1 = x1, x0
				p0, p1 = p1, p0
			}
			dx := maxInt(1, x1-x0)
			for x := x0; x <= x1; x++ {
				t := float64(x-x0) / float64(dx)
				y := int(math.Round(float64(p0.Y) + (float64(p1.Y-p0.Y) * t)))
				v2DrawLine(img, x, y, x, padT+plotH, fillCol)
			}
		}
	}

	stroke := color.RGBA{R: 10, G: 10, B: 10, A: 255}
	for i := 0; i+1 < len(points); i++ {
		v2DrawLine(img, points[i].X, points[i].Y, points[i+1].X, points[i+1].Y, stroke)
	}
}

func v2DrawPieChart(img *image.RGBA, vals []float64, palette []color.RGBA) {
	total := 0.0
	for _, v := range vals {
		if v > 0 {
			total += v
		}
	}
	if total <= 0 {
		return
	}

	b := img.Bounds()
	cx := b.Dx() / 2
	cy := b.Dy() / 2
	r := int(math.Min(float64(b.Dx()), float64(b.Dy()))/2.0) - 20
	if r < 10 {
		r = 10
	}

	// Precompute cumulative angle ranges.
	type seg struct {
		start float64
		end   float64
		col   color.RGBA
	}
	segs := make([]seg, 0, len(vals))
	start := -math.Pi / 2
	for i, v := range vals {
		if v <= 0 {
			continue
		}
		span := (v / total) * 2 * math.Pi
		segs = append(segs, seg{start: start, end: start + span, col: palette[i%len(palette)]})
		start += span
	}

	for y := cy - r; y <= cy+r; y++ {
		for x := cx - r; x <= cx+r; x++ {
			dx := float64(x - cx)
			dy := float64(y - cy)
			if dx*dx+dy*dy > float64(r*r) {
				continue
			}
			ang := math.Atan2(dy, dx)
			// Normalize to [-pi/2, 3pi/2)
			for ang < -math.Pi/2 {
				ang += 2 * math.Pi
			}
			for ang >= 3*math.Pi/2 {
				ang -= 2 * math.Pi
			}
			for _, s := range segs {
				if ang >= s.start && ang < s.end {
					img.SetRGBA(x, y, s.col)
					break
				}
			}
		}
	}
}
