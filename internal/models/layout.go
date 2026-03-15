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
	if p.Enabled && p.RowsPerPage <= 0 {
		return fmt.Errorf("pageBreak.rowsPerPage must be > 0 when enabled")
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
}

func (c ColumnConfig) Validate() error {
	if strings.TrimSpace(c.Field) == "" {
		return fmt.Errorf("column.field is required")
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
	AutoSizeColumns      bool                `json:"autoSizeColumns,omitempty" xml:"AutoSizeColumns,omitempty"`
	FreezeHeader         bool                `json:"freezeHeader,omitempty" xml:"FreezeHeader,omitempty"`
	Sheets               []SheetConfig       `json:"sheets,omitempty" xml:"Sheets>Sheet,omitempty"`

	Header      *StyleConfig       `json:"header,omitempty" xml:"Header,omitempty"`
	Body        *StyleConfig       `json:"body,omitempty" xml:"Body,omitempty"`
	Footer      *FooterConfig      `json:"footer,omitempty" xml:"Footer,omitempty"`
	HeaderImage *HeaderImageConfig `json:"headerImage,omitempty" xml:"HeaderImage,omitempty"`

	// Legacy (v2 initial): showPageNum. Kept for backward compatibility, but hidden from Swagger.
	ShowPageNum *bool          `json:"showPageNum,omitempty" xml:"ShowPageNum,omitempty" swaggerignore:"true"`
	Columns     []ColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
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
