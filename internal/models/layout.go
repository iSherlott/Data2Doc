package models

import (
	"encoding/base64"
	"fmt"
	"regexp"
	"strings"
)

var hexColorRe = regexp.MustCompile(`^#?[0-9a-fA-F]{6}$`)

func normalizeHexColor(s string) (string, error) {
	v := strings.TrimSpace(s)
	if v == "" {
		return "", nil
	}
	if !hexColorRe.MatchString(v) {
		return "", fmt.Errorf("invalid color '%s' (expected hex like #RRGGBB)", s)
	}
	if v[0] != '#' {
		v = "#" + v
	}
	return strings.ToUpper(v), nil
}

type StyleConfig struct {
	FontFamily     FontFamilyEnum        `json:"fontFamily,omitempty" xml:"FontFamily,omitempty" example:"Calibri"`
	FontSize       int                   `json:"fontSize,omitempty" xml:"FontSize,omitempty" example:"11"`
	Bold           bool                  `json:"bold,omitempty" xml:"Bold,omitempty"`
	Italic         bool                  `json:"italic,omitempty" xml:"Italic,omitempty"`
	Underline      bool                  `json:"underline,omitempty" xml:"Underline,omitempty"`
	FontColor      string                `json:"fontColor,omitempty" xml:"FontColor,omitempty" example:"#000000"`
	Background     string                `json:"background,omitempty" xml:"Background,omitempty" example:"#FFFFFF"`
	Alignment      AlignmentEnum         `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"left"`
	VerticalAlign  VerticalAlignmentEnum `json:"verticalAlignment,omitempty" xml:"VerticalAlignment,omitempty" example:"Middle"`
	Border         bool                  `json:"border,omitempty" xml:"Border,omitempty"`
	ZebraStripe    bool                  `json:"zebraStripe,omitempty" xml:"ZebraStripe,omitempty"`
	ZebraColorEven string                `json:"zebraColorEven,omitempty" xml:"ZebraColorEven,omitempty" example:"#FFFFFF"`
	ZebraColorOdd  string                `json:"zebraColorOdd,omitempty" xml:"ZebraColorOdd,omitempty" example:"#F3F3F3"`
}

func (s *StyleConfig) Validate() error {
	if s == nil {
		return nil
	}
	if s.FontFamily != "" && !s.FontFamily.IsValid() {
		return fmt.Errorf("fontFamily is invalid")
	}
	if s.FontSize < 0 {
		return fmt.Errorf("fontSize must be >= 0")
	}
	if _, err := normalizeHexColor(s.FontColor); err != nil {
		return fmt.Errorf("fontColor: %w", err)
	}
	if _, err := normalizeHexColor(s.Background); err != nil {
		return fmt.Errorf("background: %w", err)
	}
	if _, err := normalizeHexColor(s.ZebraColorEven); err != nil {
		return fmt.Errorf("zebraColorEven: %w", err)
	}
	if _, err := normalizeHexColor(s.ZebraColorOdd); err != nil {
		return fmt.Errorf("zebraColorOdd: %w", err)
	}
	if s.VerticalAlign != "" && !s.VerticalAlign.IsValid() {
		return fmt.Errorf("verticalAlignment is invalid")
	}
	return nil
}

type DefaultFontConfig struct {
	FontFamily FontFamilyEnum `json:"fontFamily,omitempty" xml:"FontFamily,omitempty" example:"Calibri"`
	FontSize   int            `json:"fontSize,omitempty" xml:"FontSize,omitempty" example:"11"`
	FontColor  string         `json:"fontColor,omitempty" xml:"FontColor,omitempty" example:"#000000"`
}

func (d *DefaultFontConfig) Validate() error {
	if d == nil {
		return nil
	}
	if d.FontFamily != "" && !d.FontFamily.IsValid() {
		return fmt.Errorf("defaultFont.fontFamily is invalid")
	}
	if d.FontSize < 0 {
		return fmt.Errorf("defaultFont.fontSize must be >= 0")
	}
	if _, err := normalizeHexColor(d.FontColor); err != nil {
		return fmt.Errorf("defaultFont.fontColor: %w", err)
	}
	return nil
}

type PageMarginConfig struct {
	Top    float64 `json:"top,omitempty" xml:"Top,omitempty" example:"20"`
	Bottom float64 `json:"bottom,omitempty" xml:"Bottom,omitempty" example:"20"`
	Left   float64 `json:"left,omitempty" xml:"Left,omitempty" example:"15"`
	Right  float64 `json:"right,omitempty" xml:"Right,omitempty" example:"15"`
}

func (m *PageMarginConfig) Validate() error {
	if m == nil {
		return nil
	}
	if m.Top < 0 || m.Bottom < 0 || m.Left < 0 || m.Right < 0 {
		return fmt.Errorf("pageMargin values must be >= 0")
	}
	return nil
}

type PageBreakConfig struct {
	Enabled     bool `json:"enabled,omitempty" xml:"Enabled,omitempty"`
	RowsPerPage int  `json:"rowsPerPage,omitempty" xml:"RowsPerPage,omitempty" example:"30"`
}

func (p *PageBreakConfig) Validate() error {
	if p == nil {
		return nil
	}
	if p.RowsPerPage < 0 {
		return fmt.Errorf("pageBreak.rowsPerPage must be >= 0")
	}
	// When enabled=true and rowsPerPage is omitted/0, pagination is automatic (by available page height).
	return nil
}

type SpacingConfig struct {
	ParagraphSpacing float64 `json:"paragraphSpacing,omitempty" xml:"ParagraphSpacing,omitempty" example:"10"`
	TableSpacing     float64 `json:"tableSpacing,omitempty" xml:"TableSpacing,omitempty" example:"15"`
}

func (s *SpacingConfig) Validate() error {
	if s == nil {
		return nil
	}
	if s.ParagraphSpacing < 0 {
		return fmt.Errorf("spacing.paragraphSpacing must be >= 0")
	}
	if s.TableSpacing < 0 {
		return fmt.Errorf("spacing.tableSpacing must be >= 0")
	}
	return nil
}

type PageNumberConfig struct {
	Enabled bool                 `json:"enabled,omitempty" xml:"Enabled,omitempty"`
	Format  PageNumberFormatEnum `json:"format,omitempty" xml:"Format,omitempty" example:"Arabic"`
}

func (p *PageNumberConfig) Validate() error {
	if p == nil {
		return nil
	}
	if p.Format != "" && !p.Format.IsValid() {
		return fmt.Errorf("pageNumber.format is invalid")
	}
	return nil
}

// FooterConfig combines footer behavior (pagination) and optional styling.
// Backward compatible: existing payloads that used layout.footer as StyleConfig still bind,
// because StyleConfig fields are embedded at the same level.
type FooterConfig struct {
	Show       *bool             `json:"show,omitempty" xml:"Show,omitempty"`
	Alignment  AlignmentEnum     `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"center"`
	PageNumber *PageNumberConfig `json:"pageNumber,omitempty" xml:"PageNumber,omitempty"`
	StyleConfig
}

func (f *FooterConfig) Validate() error {
	if f == nil {
		return nil
	}
	if err := f.StyleConfig.Validate(); err != nil {
		return err
	}
	if err := f.PageNumber.Validate(); err != nil {
		return err
	}
	return nil
}

type HeaderImageConfig struct {
	Data                string              `json:"data,omitempty" xml:"Data,omitempty"` // base64 or data-uri
	Position            ImagePositionEnum   `json:"position,omitempty" xml:"Position,omitempty" example:"top-center"`
	Height              float64             `json:"height,omitempty" xml:"Height,omitempty" example:"24"` // mm
	FitMode             ImageFitModeEnum    `json:"fitMode,omitempty" xml:"FitMode,omitempty" example:"Contain"`
	Stretch             bool                `json:"stretch,omitempty" xml:"Stretch,omitempty"`
	KeepAspectRatio     *bool               `json:"keepAspectRatio,omitempty" xml:"KeepAspectRatio,omitempty"`
	HorizontalAlignment ColumnAlignmentEnum `json:"horizontalAlignment,omitempty" xml:"HorizontalAlignment,omitempty" example:"center"`
	FillHeaderWidth     bool                `json:"fillHeaderWidth,omitempty" xml:"FillHeaderWidth,omitempty"`
}

func (c *HeaderImageConfig) Validate() error {
	if c == nil {
		return nil
	}
	if c.Height < 0 {
		return fmt.Errorf("headerImage.height must be >= 0")
	}
	if c.FitMode != "" && !c.FitMode.IsValid() {
		return fmt.Errorf("headerImage.fitMode is invalid")
	}
	if c.HorizontalAlignment != "" && !c.HorizontalAlignment.IsValid() {
		return fmt.Errorf("headerImage.horizontalAlignment is invalid")
	}
	if strings.TrimSpace(c.Data) != "" {
		if err := validateBase64OrDataURI(c.Data); err != nil {
			return fmt.Errorf("headerImage.data: %w", err)
		}
	}
	return nil
}

func validateBase64OrDataURI(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		return nil
	}
	if strings.HasPrefix(strings.ToLower(v), "data:") {
		parts := strings.SplitN(v, ",", 2)
		if len(parts) != 2 {
			return fmt.Errorf("invalid data URI")
		}
		payload := strings.TrimSpace(parts[1])
		if payload == "" {
			return fmt.Errorf("invalid data URI: empty base64 payload")
		}
		if _, err := base64.StdEncoding.DecodeString(payload); err != nil {
			return fmt.Errorf("invalid base64")
		}
		return nil
	}
	if _, err := base64.StdEncoding.DecodeString(v); err != nil {
		return fmt.Errorf("invalid base64")
	}
	return nil
}

type ColumnConfig struct {
	Field             string                `json:"field" xml:"Field" example:"name"`
	Title             string                `json:"title,omitempty" xml:"Title,omitempty" example:"Name"`
	Width             float64               `json:"width,omitempty" xml:"Width,omitempty" example:"20"` // Excel: characters; PDF: relative
	Alignment         ColumnAlignmentEnum   `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"left"`
	VerticalAlignment VerticalAlignmentEnum `json:"verticalAlignment,omitempty" xml:"VerticalAlignment,omitempty" example:"Middle"`
	Format            ColumnFormatEnum      `json:"format,omitempty" xml:"Format,omitempty" example:"currency"`

	// Excel calculation features
	// formula is an expression based on fields (e.g. "price * qty") that will be resolved to an Excel formula per-row.
	Formula string `json:"formula,omitempty" xml:"Formula,omitempty" example:"price * qty"`
	// sheetFormula is a raw Excel formula (can reference other sheets). Mutually exclusive with formula.
	SheetFormula string `json:"sheetFormula,omitempty" xml:"SheetFormula,omitempty" example:"SUM(Employees!B2:B100)"`
	// aggregate supports automatic totals/subtotals (currently: "sum").
	Aggregate string `json:"aggregate,omitempty" xml:"Aggregate,omitempty" example:"sum"`
	// percentageOf makes this column a percentage of another field's total (e.g. salary / totalSalary).
	PercentageOf string `json:"percentageOf,omitempty" xml:"PercentageOf,omitempty" example:"salary"`

	// Excel-only advanced options
	CellType ExcelCellTypeEnum `json:"cellType,omitempty" xml:"CellType,omitempty" example:"Select"`
	Options  []string          `json:"options,omitempty" xml:"Options>Option,omitempty"`
	Hidden   bool              `json:"hidden,omitempty" xml:"Hidden,omitempty"`
	Locked   bool              `json:"locked,omitempty" xml:"Locked,omitempty"`
}

func (c ColumnConfig) Validate() error {
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
	if c.CellType == ExcelCellSelect {
		if len(c.Options) == 0 {
			return fmt.Errorf("column.options is required when cellType is Select")
		}
		for i := range c.Options {
			if strings.TrimSpace(c.Options[i]) == "" {
				return fmt.Errorf("column.options[%d] must be non-empty", i)
			}
		}
	}
	return nil
}

type SheetConfig struct {
	Name       string `json:"name" xml:"Name" example:"Funcionarios"`
	DataSource string `json:"dataSource,omitempty" xml:"DataSource,omitempty" example:"employees"`
}

func (s SheetConfig) Validate() error {
	if strings.TrimSpace(s.Name) == "" {
		return fmt.Errorf("sheet.name is required")
	}
	return nil
}

type LayoutConfig struct {
	PageOrientation      PageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Portrait"`
	PageMargin           *PageMarginConfig   `json:"pageMargin,omitempty" xml:"PageMargin,omitempty"`
	UsePageContentBounds *bool               `json:"usePageContentBounds,omitempty" xml:"UsePageContentBounds,omitempty"`
	DefaultFont          *DefaultFontConfig  `json:"defaultFont,omitempty" xml:"DefaultFont,omitempty"`
	PageBreak            *PageBreakConfig    `json:"pageBreak,omitempty" xml:"PageBreak,omitempty"`
	Spacing              *SpacingConfig      `json:"spacing,omitempty" xml:"Spacing,omitempty"`
	AutoSizeColumns      bool                `json:"autoSizeColumns,omitempty" xml:"AutoSizeColumns,omitempty"`
	FreezeHeader         bool                `json:"freezeHeader,omitempty" xml:"FreezeHeader,omitempty"`
	FreezeColumns        int                 `json:"freezeColumns,omitempty" xml:"FreezeColumns,omitempty" example:"1"`
	HideEmptyRows        bool                `json:"hideEmptyRows,omitempty" xml:"HideEmptyRows,omitempty"`
	MaxVisibleRows       int                 `json:"maxVisibleRows,omitempty" xml:"MaxVisibleRows,omitempty" example:"100"`
	GroupBy              string              `json:"groupBy,omitempty" xml:"GroupBy,omitempty" example:"department"`
	ShowTotalRow         bool                `json:"showTotalRow,omitempty" xml:"ShowTotalRow,omitempty"`
	Sheets               []SheetConfig       `json:"sheets,omitempty" xml:"Sheets>Sheet,omitempty"`

	Header      *StyleConfig       `json:"header,omitempty" xml:"Header,omitempty"`
	Body        *StyleConfig       `json:"body,omitempty" xml:"Body,omitempty"`
	Footer      *FooterConfig      `json:"footer,omitempty" xml:"Footer,omitempty"`
	HeaderImage *HeaderImageConfig `json:"headerImage,omitempty" xml:"HeaderImage,omitempty"`

	// Legacy (v2 initial): showPageNum. Kept for backward compatibility, but hidden from Swagger.
	ShowPageNum *bool          `json:"showPageNum,omitempty" xml:"ShowPageNum,omitempty" swaggerignore:"true"`
	Columns     []ColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`

	// PDF-only: declarative ordered content blocks (document builder mode).
	Blocks []PDFBlockConfig `json:"blocks,omitempty" xml:"Blocks>Block,omitempty"`

	// Word-only flags (internal; request uses layout.word.*).
	WordIgnorePageMargins bool
	WordCenterContent     *bool
}

func (l *LayoutConfig) Validate() error {
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
	for i := range l.Sheets {
		if err := l.Sheets[i].Validate(); err != nil {
			return fmt.Errorf("sheets[%d]: %w", i, err)
		}
	}
	for i := range l.Columns {
		if err := l.Columns[i].Validate(); err != nil {
			return fmt.Errorf("columns[%d]: %w", i, err)
		}
	}
	return nil
}

// ---- PDF declarative blocks ----

type PDFTextStyleConfig struct {
	FontSize    int           `json:"fontSize,omitempty" xml:"FontSize,omitempty" example:"12"`
	Bold        *bool         `json:"bold,omitempty" xml:"Bold,omitempty"`
	Italic      *bool         `json:"italic,omitempty" xml:"Italic,omitempty"`
	Alignment   AlignmentEnum `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"left"`
	LineSpacing float64       `json:"lineSpacing,omitempty" xml:"LineSpacing,omitempty" example:"1.2"`
}

func (s *PDFTextStyleConfig) Validate() error {
	if s == nil {
		return nil
	}
	if s.FontSize < 0 {
		return fmt.Errorf("style.fontSize must be >= 0")
	}
	if s.Alignment != "" && !s.Alignment.IsValid() {
		return fmt.Errorf("style.alignment is invalid")
	}
	if s.LineSpacing < 0 {
		return fmt.Errorf("style.lineSpacing must be >= 0")
	}
	return nil
}

type PDFTableColumnConfig struct {
	Field string `json:"field" xml:"Field" example:"name"`
	Title string `json:"title,omitempty" xml:"Title,omitempty" example:"Nome"`
}

func (c PDFTableColumnConfig) Validate() error {
	if strings.TrimSpace(c.Field) == "" {
		return fmt.Errorf("column.field is required")
	}
	return nil
}

type PDFBlockConfig struct {
	Type PDFBlockTypeEnum `json:"type" xml:"Type" example:"Text"`

	// Common fields
	Content string              `json:"content,omitempty" xml:"Content,omitempty"`
	Style   *PDFTextStyleConfig `json:"style,omitempty" xml:"Style,omitempty"`

	// Table
	DataSource string                 `json:"dataSource,omitempty" xml:"DataSource,omitempty" example:"employees"`
	Columns    []PDFTableColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`

	// Chart
	ChartType     ChartTypeEnum `json:"chartType,omitempty" xml:"ChartType,omitempty" example:"Bar"`
	CategoryField string        `json:"categoryField,omitempty" xml:"CategoryField,omitempty" example:"department"`
	ValueField    string        `json:"valueField,omitempty" xml:"ValueField,omitempty" example:"total"`
	Title         string        `json:"title,omitempty" xml:"Title,omitempty" example:"Vendas por Departamento"`
	Width         float64       `json:"width,omitempty" xml:"Width,omitempty" example:"160"`
	Height        float64       `json:"height,omitempty" xml:"Height,omitempty" example:"90"`

	// Image
	Data      string        `json:"data,omitempty" xml:"Data,omitempty"` // base64 or data-uri
	Alignment AlignmentEnum `json:"alignment,omitempty" xml:"Alignment,omitempty" example:"center"`
}

func (b *PDFBlockConfig) Validate() error {
	if b == nil {
		return nil
	}
	if b.Type == "" {
		return fmt.Errorf("type is required")
	}
	if !b.Type.IsValid() {
		return fmt.Errorf("type is invalid")
	}
	if err := b.Style.Validate(); err != nil {
		return err
	}
	if b.Width < 0 {
		return fmt.Errorf("width must be >= 0")
	}
	if b.Height < 0 {
		return fmt.Errorf("height must be >= 0")
	}
	if b.Alignment != "" && !b.Alignment.IsValid() {
		return fmt.Errorf("alignment is invalid")
	}

	switch b.Type {
	case PDFBlockText:
		if strings.TrimSpace(b.Content) == "" {
			return fmt.Errorf("content is required for Text")
		}
	case PDFBlockSectionTitle:
		if strings.TrimSpace(b.Content) == "" {
			return fmt.Errorf("content is required for SectionTitle")
		}
	case PDFBlockSpacer:
		// Height is used.
		return nil
	case PDFBlockPageBreak:
		// No fields.
		return nil
	case PDFBlockImage:
		if strings.TrimSpace(b.Data) == "" {
			return fmt.Errorf("data is required for Image")
		}
		if err := validateBase64OrDataURI(b.Data); err != nil {
			return fmt.Errorf("data: %w", err)
		}
	case PDFBlockTable:
		if len(b.Columns) == 0 {
			return fmt.Errorf("columns is required for Table")
		}
		seen := map[string]bool{}
		for i := range b.Columns {
			if err := b.Columns[i].Validate(); err != nil {
				return fmt.Errorf("columns[%d]: %w", i, err)
			}
			f := strings.ToLower(strings.TrimSpace(b.Columns[i].Field))
			if seen[f] {
				return fmt.Errorf("columns[%d].field '%s' is duplicated", i, b.Columns[i].Field)
			}
			seen[f] = true
		}
	case PDFBlockChart:
		if b.ChartType == "" {
			return fmt.Errorf("chartType is required for Chart")
		}
		if !b.ChartType.IsValid() {
			return fmt.Errorf("chartType is invalid")
		}
		if strings.TrimSpace(b.CategoryField) == "" {
			return fmt.Errorf("categoryField is required for Chart")
		}
		if strings.TrimSpace(b.ValueField) == "" {
			return fmt.Errorf("valueField is required for Chart")
		}
	default:
		return fmt.Errorf("unsupported block type")
	}
	return nil
}
