package models

import (
	"fmt"
	"strings"
)

// ExcelGenerateRequest is the request body for POST /generate/excel.
// It intentionally only exposes options relevant to Excel generation.
type ExcelGenerateRequest struct {
	Layout *ExcelLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data   DataPayload        `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *ExcelGenerateRequest) Validate() error {
	r.normalizeDefaultDataset()

	if err := r.Layout.Validate(); err != nil {
		return err
	}
	if r.Layout != nil && len(r.Layout.Sheets) > 0 {
		seen := map[string]bool{}
		for i := range r.Layout.Sheets {
			sh := r.Layout.Sheets[i]
			name := strings.TrimSpace(sh.Name)
			if name == "" {
				return fmt.Errorf("sheets[%d].name is required", i)
			}
			if seen[strings.ToLower(name)] {
				return fmt.Errorf("sheets[%d].name '%s' is duplicated", i, name)
			}
			seen[strings.ToLower(name)] = true

			ds := r.Data.Get(sh.DataSource)
			if ds.IsEmpty() {
				dsName := strings.TrimSpace(sh.DataSource)
				if dsName == "" {
					dsName = "default"
				}
				return fmt.Errorf("dataSource '%s' (sheet '%s') is empty or missing", dsName, name)
			}
		}
	}
	// For Excel without sheets, the renderer uses the "default" dataset.
	if (r.Layout == nil || len(r.Layout.Sheets) == 0) && r.Data.Default.IsEmpty() {
		if len(r.Data.Sources) > 0 {
			return fmt.Errorf("layout.sheets is required when data is an object of datasets; otherwise send data as an array/object (default dataset)")
		}
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	if r.Data.IsEmpty() {
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	if r.Layout != nil {
		validateDataset := func(dsName string, data DynamicData) error {
			if r.Layout.MaxVisibleRows > 0 && len(data.Items) > r.Layout.MaxVisibleRows {
				return fmt.Errorf("maxVisibleRows must be >= number of records (dataset '%s')", dsName)
			}
			if r.Layout.FreezeColumns > 0 {
				colCount := len(r.Layout.Columns)
				if colCount == 0 {
					if len(data.Order) > 0 {
						colCount = len(data.Order)
					} else if len(data.Items) > 0 {
						colCount = len(data.Items[0])
					}
				}
				if colCount > 0 && r.Layout.FreezeColumns > colCount {
					return fmt.Errorf("freezeColumns must be <= number of columns (dataset '%s')", dsName)
				}
			}
			return nil
		}

		if len(r.Layout.Sheets) > 0 {
			for _, sh := range r.Layout.Sheets {
				name := strings.TrimSpace(sh.DataSource)
				if name == "" {
					name = "default"
				}
				if err := validateDataset(name, r.Data.Get(sh.DataSource)); err != nil {
					return err
				}
			}
		} else {
			if err := validateDataset("default", r.Data.Default); err != nil {
				return err
			}
		}
	}
	return nil
}

func (r *ExcelGenerateRequest) normalizeDefaultDataset() {
	if r == nil {
		return
	}
	if !r.Data.Default.IsEmpty() {
		return
	}
	if len(r.Data.Sources) != 1 {
		return
	}
	for _, d := range r.Data.Sources {
		if !d.IsEmpty() {
			r.Data.Default = d
			return
		}
	}
}

func (r ExcelGenerateRequest) ToDocumentRequest() DocumentRequest {
	out := DocumentRequest{
		Format: DocumentFormatExcel,
		Layout: nil,
		Data:   r.Data,
	}
	if r.Layout != nil {
		out.Layout = r.Layout.ToLayoutConfig()
	}
	return out
}

// PDFGenerateRequest is the request body for POST /generate/pdf.
// It intentionally only exposes options relevant to PDF generation.
type PDFGenerateRequest struct {
	Layout *PDFLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data   DataPayload      `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *PDFGenerateRequest) Validate() error {
	r.normalizeDefaultDataset()

	if err := r.Layout.Validate(); err != nil {
		return err
	}

	// Builder mode: allow data to be empty when no block requires a dataset.
	blocks := []PDFBlockConfig(nil)
	if r.Layout != nil {
		blocks = r.Layout.Blocks
	}
	if len(blocks) > 0 {
		needsData := false
		for i := range blocks {
			b := blocks[i]
			switch b.Type {
			case PDFBlockTable, PDFBlockChart:
				needsData = true
				ds := r.Data.Get(b.DataSource)
				if ds.IsEmpty() {
					dsName := strings.TrimSpace(b.DataSource)
					if dsName == "" {
						dsName = "default"
					}
					return fmt.Errorf("dataSource '%s' (blocks[%d]) is empty or missing", dsName, i)
				}
			}
		}
		if needsData {
			return nil
		}
		// No dataset-backed blocks => data can be empty.
		return nil
	}

	// Legacy mode: expects a single dataset as default.
	if r.Data.Default.IsEmpty() {
		if len(r.Data.Sources) > 0 {
			return fmt.Errorf("default dataset is required when layout.blocks is omitted; otherwise send layout.blocks and reference dataSource datasets")
		}
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	return nil
}

func (r *PDFGenerateRequest) normalizeDefaultDataset() {
	if r == nil {
		return
	}
	if !r.Data.Default.IsEmpty() {
		return
	}
	if len(r.Data.Sources) != 1 {
		return
	}
	for _, d := range r.Data.Sources {
		if !d.IsEmpty() {
			r.Data.Default = d
			return
		}
	}
}

func (r PDFGenerateRequest) ToDocumentRequest() DocumentRequest {
	out := DocumentRequest{
		Format: DocumentFormatPDF,
		Layout: nil,
		Data:   r.Data,
	}
	if r.Layout != nil {
		out.Layout = r.Layout.ToLayoutConfig()
	}
	return out
}

// WordGenerateRequest is the request body for POST /generate/word.
// It intentionally only exposes options relevant to Word (DOCX) generation.
type WordGenerateRequest struct {
	Layout *WordLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data   DataPayload       `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *WordGenerateRequest) Validate() error {
	r.normalizeDefaultDataset()

	if err := r.Layout.Validate(); err != nil {
		return err
	}

	// Builder mode: allow data to be empty when no block requires a dataset.
	blocks := []PDFBlockConfig(nil)
	if r.Layout != nil {
		blocks = r.Layout.Blocks
	}
	if len(blocks) > 0 {
		needsData := false
		for i := range blocks {
			b := blocks[i]
			switch b.Type {
			case PDFBlockTable, PDFBlockChart:
				needsData = true
				ds := r.Data.Get(b.DataSource)
				if ds.IsEmpty() {
					dsName := strings.TrimSpace(b.DataSource)
					if dsName == "" {
						dsName = "default"
					}
					return fmt.Errorf("dataSource '%s' (blocks[%d]) is empty or missing", dsName, i)
				}
			}
		}
		if needsData {
			return nil
		}
		return nil
	}

	// Legacy mode: expects a single dataset as default.
	if r.Data.Default.IsEmpty() {
		if len(r.Data.Sources) > 0 {
			return fmt.Errorf("default dataset is required when layout.blocks is omitted; otherwise send layout.blocks and reference dataSource datasets")
		}
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	return nil
}

func (r *WordGenerateRequest) normalizeDefaultDataset() {
	if r == nil {
		return
	}
	if !r.Data.Default.IsEmpty() {
		return
	}
	if len(r.Data.Sources) != 1 {
		return
	}
	for _, d := range r.Data.Sources {
		if !d.IsEmpty() {
			r.Data.Default = d
			return
		}
	}
}

func (r WordGenerateRequest) ToDocumentRequest() DocumentRequest {
	out := DocumentRequest{
		Format: DocumentFormatWord,
		Layout: nil,
		Data:   r.Data,
	}
	if r.Layout != nil {
		out.Layout = r.Layout.ToLayoutConfig()
	}
	return out
}

type ExcelColumnConfig struct {
	Field             string                `json:"field" xml:"Field" example:"name"`
	Title             string                `json:"title,omitempty" xml:"Title,omitempty" example:"Name"`
	Width             float64               `json:"width,omitempty" xml:"Width,omitempty" example:"20"` // characters
	Alignment         ColumnAlignmentEnum   `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"left"`
	VerticalAlignment VerticalAlignmentEnum `json:"verticalAlignment,omitempty" xml:"VerticalAlignment,omitempty" example:"Middle"`
	Format            ColumnFormatEnum      `json:"format,omitempty" xml:"Format,omitempty" example:"currency"`

	// Excel calculation features
	Formula      string `json:"formula,omitempty" xml:"Formula,omitempty" example:"price * qty"`
	SheetFormula string `json:"sheetFormula,omitempty" xml:"SheetFormula,omitempty" example:"SUM(Employees!B2:B100)"`
	Aggregate    string `json:"aggregate,omitempty" xml:"Aggregate,omitempty" example:"sum"`
	PercentageOf string `json:"percentageOf,omitempty" xml:"PercentageOf,omitempty" example:"salary"`

	CellType              ExcelCellTypeEnum                `json:"cellType,omitempty" xml:"CellType,omitempty" example:"Select"`
	Options               []string                         `json:"options,omitempty" xml:"Options>Option,omitempty"`
	ValidationRange       string                           `json:"validationRange,omitempty" xml:"ValidationRange,omitempty" example:"Products!A2:A50"`
	Lookup                *ExcelLookupConfig               `json:"lookup,omitempty" xml:"Lookup,omitempty"`
	ConditionalFormatting []ExcelConditionalFormattingRule `json:"conditionalFormatting,omitempty" xml:"ConditionalFormatting>Rule,omitempty"`
	BackgroundColor       string                           `json:"backgroundColor,omitempty" xml:"BackgroundColor,omitempty" example:"#FFFFFF"`
	TextColor             string                           `json:"textColor,omitempty" xml:"TextColor,omitempty" example:"#000000"`
	HeaderColor           string                           `json:"headerColor,omitempty" xml:"HeaderColor,omitempty" example:"#FFFFFF"`
	Hidden                bool                             `json:"hidden,omitempty" xml:"Hidden,omitempty"`
	Locked                bool                             `json:"locked,omitempty" xml:"Locked,omitempty"`
}

func (c ExcelColumnConfig) Validate() error {
	if strings.TrimSpace(c.Field) == "" {
		return fmt.Errorf("column.field is required")
	}
	if strings.TrimSpace(c.Formula) != "" && strings.TrimSpace(c.SheetFormula) != "" {
		return fmt.Errorf("column.formula and column.sheetFormula are mutually exclusive")
	}
	if strings.TrimSpace(c.Aggregate) != "" {
		agg := strings.ToLower(strings.TrimSpace(c.Aggregate))
		if agg != "sum" {
			return fmt.Errorf("column.aggregate is invalid")
		}
	}
	if strings.TrimSpace(c.PercentageOf) != "" {
		if strings.TrimSpace(c.PercentageOf) == strings.TrimSpace(c.Field) {
			return fmt.Errorf("column.percentageOf cannot reference itself")
		}
	}
	if c.Width < 0 {
		return fmt.Errorf("column.width must be >= 0")
	}
	if c.VerticalAlignment != "" && !c.VerticalAlignment.IsValid() {
		return fmt.Errorf("column.verticalAlignment is invalid")
	}
	if c.Format != "" && !c.Format.IsValid() {
		return fmt.Errorf("column.format is invalid")
	}
	if c.CellType != "" && !c.CellType.IsValid() {
		return fmt.Errorf("column.cellType is invalid")
	}
	if c.CellType == ExcelCellLookup {
		if c.Lookup == nil {
			return fmt.Errorf("column.lookup is required when cellType is Lookup")
		}
		if err := c.Lookup.Validate(); err != nil {
			return fmt.Errorf("column.lookup: %w", err)
		}
	} else if c.Lookup != nil {
		return fmt.Errorf("column.lookup is only allowed when cellType is Lookup")
	}
	if c.CellType == ExcelCellSelect {
		if strings.TrimSpace(c.ValidationRange) != "" && len(c.Options) > 0 {
			return fmt.Errorf("column.validationRange and column.options are mutually exclusive")
		}
		if strings.TrimSpace(c.ValidationRange) == "" && len(c.Options) == 0 {
			return fmt.Errorf("column.options or column.validationRange is required when cellType is Select")
		}
		for i := range c.Options {
			if strings.TrimSpace(c.Options[i]) == "" {
				return fmt.Errorf("column.options[%d] must be non-empty", i)
			}
		}
	}
	if _, err := normalizeHexColor(c.BackgroundColor); err != nil {
		return fmt.Errorf("column.backgroundColor: %w", err)
	}
	if _, err := normalizeHexColor(c.TextColor); err != nil {
		return fmt.Errorf("column.textColor: %w", err)
	}
	if _, err := normalizeHexColor(c.HeaderColor); err != nil {
		return fmt.Errorf("column.headerColor: %w", err)
	}
	for i := range c.ConditionalFormatting {
		if err := c.ConditionalFormatting[i].Validate(); err != nil {
			return fmt.Errorf("column.conditionalFormatting[%d]: %w", i, err)
		}
	}
	return nil
}

type TableColumnConfig struct {
	Field             string                `json:"field" xml:"Field" example:"name"`
	Title             string                `json:"title,omitempty" xml:"Title,omitempty" example:"Name"`
	Width             float64               `json:"width,omitempty" xml:"Width,omitempty" example:"20"` // relative
	Alignment         ColumnAlignmentEnum   `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"left"`
	VerticalAlignment VerticalAlignmentEnum `json:"verticalAlignment,omitempty" xml:"VerticalAlignment,omitempty" example:"Middle"`
}

func (c TableColumnConfig) Validate() error {
	if strings.TrimSpace(c.Field) == "" {
		return fmt.Errorf("column.field is required")
	}
	if c.Width < 0 {
		return fmt.Errorf("column.width must be >= 0")
	}
	if c.VerticalAlignment != "" && !c.VerticalAlignment.IsValid() {
		return fmt.Errorf("column.verticalAlignment is invalid")
	}
	return nil
}

// ExcelLayoutConfig is a subset of LayoutConfig that is relevant to Excel.
type ExcelLayoutConfig struct {
	PageOrientation PageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Portrait"`
	PageMargin      *PageMarginConfig   `json:"pageMargin,omitempty" xml:"PageMargin,omitempty"`
	DefaultFont     *DefaultFontConfig  `json:"defaultFont,omitempty" xml:"DefaultFont,omitempty"`
	AutoSizeColumns bool                `json:"autoSizeColumns,omitempty" xml:"AutoSizeColumns,omitempty"`
	FreezeHeader    bool                `json:"freezeHeader,omitempty" xml:"FreezeHeader,omitempty"`
	FreezeColumns   int                 `json:"freezeColumns,omitempty" xml:"FreezeColumns,omitempty" example:"1"`
	HideEmptyRows   bool                `json:"hideEmptyRows,omitempty" xml:"HideEmptyRows,omitempty"`
	MaxVisibleRows  int                 `json:"maxVisibleRows,omitempty" xml:"MaxVisibleRows,omitempty" example:"100"`
	GroupBy         string              `json:"groupBy,omitempty" xml:"GroupBy,omitempty" example:"department"`
	ShowTotalRow    bool                `json:"showTotalRow,omitempty" xml:"ShowTotalRow,omitempty"`
	Sheets          []SheetConfig       `json:"sheets,omitempty" xml:"Sheets>Sheet,omitempty"`
	Charts          []ExcelChartConfig  `json:"charts,omitempty" xml:"Charts>Chart,omitempty"`
	Header          *StyleConfig        `json:"header,omitempty" xml:"Header,omitempty"`
	Body            *StyleConfig        `json:"body,omitempty" xml:"Body,omitempty"`
	Columns         []ExcelColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
}

func (l *ExcelLayoutConfig) Validate() error {
	if l == nil {
		return nil
	}
	if l.PageOrientation != "" && !l.PageOrientation.IsValid() {
		return fmt.Errorf("pageOrientation is invalid")
	}
	if err := l.PageMargin.Validate(); err != nil {
		return fmt.Errorf("pageMargin: %w", err)
	}
	if err := l.DefaultFont.Validate(); err != nil {
		return err
	}
	if l.FreezeColumns < 0 {
		return fmt.Errorf("freezeColumns must be >= 0")
	}
	if l.MaxVisibleRows < 0 {
		return fmt.Errorf("maxVisibleRows must be >= 0")
	}
	if err := l.Header.Validate(); err != nil {
		return fmt.Errorf("header: %w", err)
	}
	if err := l.Body.Validate(); err != nil {
		return fmt.Errorf("body: %w", err)
	}
	for i := range l.Sheets {
		if err := l.Sheets[i].Validate(); err != nil {
			return fmt.Errorf("sheets[%d]: %w", i, err)
		}
	}
	for i := range l.Charts {
		if err := l.Charts[i].Validate(); err != nil {
			return fmt.Errorf("charts[%d]: %w", i, err)
		}
	}
	seenFields := map[string]bool{}
	for i := range l.Columns {
		if err := l.Columns[i].Validate(); err != nil {
			return fmt.Errorf("columns[%d]: %w", i, err)
		}
		f := strings.ToLower(strings.TrimSpace(l.Columns[i].Field))
		if seenFields[f] {
			return fmt.Errorf("columns[%d].field '%s' is duplicated", i, l.Columns[i].Field)
		}
		seenFields[f] = true
	}
	return nil
}

func (l *ExcelLayoutConfig) ToLayoutConfig() *LayoutConfig {
	if l == nil {
		return nil
	}
	out := &LayoutConfig{
		PageOrientation: l.PageOrientation,
		PageMargin:      l.PageMargin,
		DefaultFont:     l.DefaultFont,
		AutoSizeColumns: l.AutoSizeColumns,
		FreezeHeader:    l.FreezeHeader,
		FreezeColumns:   l.FreezeColumns,
		HideEmptyRows:   l.HideEmptyRows,
		MaxVisibleRows:  l.MaxVisibleRows,
		GroupBy:         l.GroupBy,
		ShowTotalRow:    l.ShowTotalRow,
		Sheets:          l.Sheets,
		Charts:          l.Charts,
		Header:          l.Header,
		Body:            l.Body,
	}
	if len(l.Columns) > 0 {
		out.Columns = make([]ColumnConfig, 0, len(l.Columns))
		for _, c := range l.Columns {
			out.Columns = append(out.Columns, ColumnConfig{
				Field:                 c.Field,
				Title:                 c.Title,
				Width:                 c.Width,
				Alignment:             c.Alignment,
				VerticalAlignment:     c.VerticalAlignment,
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
	}
	return out
}

// PDFLayoutConfig is a subset of LayoutConfig that is relevant to PDF.
type PDFLayoutConfig struct {
	PageOrientation      PageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Portrait"`
	PageMargin           *PageMarginConfig   `json:"pageMargin,omitempty" xml:"PageMargin,omitempty"`
	UsePageContentBounds *bool               `json:"usePageContentBounds,omitempty" xml:"UsePageContentBounds,omitempty"`
	DefaultFont          *DefaultFontConfig  `json:"defaultFont,omitempty" xml:"DefaultFont,omitempty"`
	PageBreak            *PageBreakConfig    `json:"pageBreak,omitempty" xml:"PageBreak,omitempty"`
	Spacing              *SpacingConfig      `json:"spacing,omitempty" xml:"Spacing,omitempty"`
	Header               *StyleConfig        `json:"header,omitempty" xml:"Header,omitempty"`
	Body                 *StyleConfig        `json:"body,omitempty" xml:"Body,omitempty"`
	Footer               *FooterConfig       `json:"footer,omitempty" xml:"Footer,omitempty"`
	HeaderImage          *HeaderImageConfig  `json:"headerImage,omitempty" xml:"HeaderImage,omitempty"`
	Columns              []TableColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
	Blocks               []PDFBlockConfig    `json:"blocks,omitempty" xml:"Blocks>Block,omitempty"`
}

func (l *PDFLayoutConfig) Validate() error {
	if l == nil {
		return nil
	}
	if l.PageOrientation != "" && !l.PageOrientation.IsValid() {
		return fmt.Errorf("pageOrientation is invalid")
	}
	if err := l.PageMargin.Validate(); err != nil {
		return fmt.Errorf("pageMargin: %w", err)
	}
	if err := l.DefaultFont.Validate(); err != nil {
		return err
	}
	if err := l.PageBreak.Validate(); err != nil {
		return err
	}
	if err := l.Spacing.Validate(); err != nil {
		return err
	}
	if err := l.Header.Validate(); err != nil {
		return fmt.Errorf("header: %w", err)
	}
	if err := l.Body.Validate(); err != nil {
		return fmt.Errorf("body: %w", err)
	}
	if err := l.Footer.Validate(); err != nil {
		return fmt.Errorf("footer: %w", err)
	}
	if err := l.HeaderImage.Validate(); err != nil {
		return fmt.Errorf("headerImage: %w", err)
	}
	for i := range l.Blocks {
		if err := l.Blocks[i].Validate(); err != nil {
			return fmt.Errorf("blocks[%d]: %w", i, err)
		}
	}
	seenFields := map[string]bool{}
	for i := range l.Columns {
		if err := l.Columns[i].Validate(); err != nil {
			return fmt.Errorf("columns[%d]: %w", i, err)
		}
		f := strings.ToLower(strings.TrimSpace(l.Columns[i].Field))
		if seenFields[f] {
			return fmt.Errorf("columns[%d].field '%s' is duplicated", i, l.Columns[i].Field)
		}
		seenFields[f] = true
	}
	return nil
}

func (l *PDFLayoutConfig) ToLayoutConfig() *LayoutConfig {
	if l == nil {
		return nil
	}
	out := &LayoutConfig{
		PageOrientation:      l.PageOrientation,
		PageMargin:           l.PageMargin,
		UsePageContentBounds: l.UsePageContentBounds,
		DefaultFont:          l.DefaultFont,
		PageBreak:            l.PageBreak,
		Spacing:              l.Spacing,
		Header:               l.Header,
		Body:                 l.Body,
		Footer:               l.Footer,
		HeaderImage:          l.HeaderImage,
		Blocks:               l.Blocks,
	}
	if len(l.Columns) > 0 {
		out.Columns = make([]ColumnConfig, 0, len(l.Columns))
		for _, c := range l.Columns {
			out.Columns = append(out.Columns, ColumnConfig{
				Field:             c.Field,
				Title:             c.Title,
				Width:             c.Width,
				Alignment:         c.Alignment,
				VerticalAlignment: c.VerticalAlignment,
			})
		}
	}
	return out
}

// WordLayoutConfig is a subset of LayoutConfig that is relevant to Word (DOCX).
type WordLayoutConfig struct {
	Word            *WordConfig         `json:"word,omitempty" xml:"Word,omitempty"`
	PageOrientation PageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Portrait"`
	PageMargin      *PageMarginConfig   `json:"pageMargin,omitempty" xml:"PageMargin,omitempty"`
	DefaultFont     *DefaultFontConfig  `json:"defaultFont,omitempty" xml:"DefaultFont,omitempty"`
	PageBreak       *PageBreakConfig    `json:"pageBreak,omitempty" xml:"PageBreak,omitempty"`
	Header          *StyleConfig        `json:"header,omitempty" xml:"Header,omitempty"`
	Body            *StyleConfig        `json:"body,omitempty" xml:"Body,omitempty"`
	Footer          *FooterConfig       `json:"footer,omitempty" xml:"Footer,omitempty"`
	Columns         []TableColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
	Blocks          []PDFBlockConfig    `json:"blocks,omitempty" xml:"Blocks>Block,omitempty"`
}

// WordConfig groups Word-specific options under layout.word.*
type WordConfig struct {
	IgnorePageMargins bool                    `json:"ignorePageMargins,omitempty" xml:"IgnorePageMargins,omitempty"`
	CenterContent     *bool                   `json:"centerContent,omitempty" xml:"CenterContent,omitempty"`
	PageOrientation   WordPageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Landscape"`
}

func (w *WordConfig) Validate() error {
	if w == nil {
		return nil
	}
	if w.PageOrientation != "" && !w.PageOrientation.IsValid() {
		return fmt.Errorf("layout.word.pageOrientation is invalid")
	}
	return nil
}

func (l *WordLayoutConfig) Validate() error {
	if l == nil {
		return nil
	}
	if err := l.Word.Validate(); err != nil {
		return err
	}
	if l.PageOrientation != "" && !l.PageOrientation.IsValid() {
		return fmt.Errorf("pageOrientation is invalid")
	}
	if err := l.PageMargin.Validate(); err != nil {
		return fmt.Errorf("pageMargin: %w", err)
	}
	if err := l.DefaultFont.Validate(); err != nil {
		return err
	}
	if err := l.PageBreak.Validate(); err != nil {
		return err
	}
	if err := l.Header.Validate(); err != nil {
		return fmt.Errorf("header: %w", err)
	}
	if err := l.Body.Validate(); err != nil {
		return fmt.Errorf("body: %w", err)
	}
	if err := l.Footer.Validate(); err != nil {
		return fmt.Errorf("footer: %w", err)
	}
	for i := range l.Blocks {
		if err := l.Blocks[i].Validate(); err != nil {
			return fmt.Errorf("blocks[%d]: %w", i, err)
		}
	}
	seenFields := map[string]bool{}
	for i := range l.Columns {
		if err := l.Columns[i].Validate(); err != nil {
			return fmt.Errorf("columns[%d]: %w", i, err)
		}
		f := strings.ToLower(strings.TrimSpace(l.Columns[i].Field))
		if seenFields[f] {
			return fmt.Errorf("columns[%d].field '%s' is duplicated", i, l.Columns[i].Field)
		}
		seenFields[f] = true
	}
	return nil
}

func (l *WordLayoutConfig) ToLayoutConfig() *LayoutConfig {
	if l == nil {
		return nil
	}
	pageOri := l.PageOrientation
	useBounds := (*bool)(nil)
	wordCenter := (*bool)(nil)
	wordIgnoreMargins := false
	if l.Word != nil {
		if l.Word.PageOrientation != "" {
			if l.Word.PageOrientation == WordLandscape {
				pageOri = PageLandscape
			} else {
				pageOri = PagePortrait
			}
		}
		if l.Word.CenterContent != nil {
			wordCenter = l.Word.CenterContent
		}
		if l.Word.IgnorePageMargins {
			wordIgnoreMargins = true
			v := false
			useBounds = &v
		}
	}
	out := &LayoutConfig{
		PageOrientation:       pageOri,
		PageMargin:            l.PageMargin,
		DefaultFont:           l.DefaultFont,
		PageBreak:             l.PageBreak,
		Header:                l.Header,
		Body:                  l.Body,
		Footer:                l.Footer,
		UsePageContentBounds:  useBounds,
		WordIgnorePageMargins: wordIgnoreMargins,
		WordCenterContent:     wordCenter,
		Blocks:                l.Blocks,
	}
	if len(l.Columns) > 0 {
		out.Columns = make([]ColumnConfig, 0, len(l.Columns))
		for _, c := range l.Columns {
			out.Columns = append(out.Columns, ColumnConfig{
				Field:             c.Field,
				Title:             c.Title,
				Width:             c.Width,
				Alignment:         c.Alignment,
				VerticalAlignment: c.VerticalAlignment,
			})
		}
	}
	return out
}
