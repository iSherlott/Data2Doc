package models

import (
	"fmt"
	"strings"
)

// ExcelGenerateRequest is the request body for POST /generate/excel.
// It intentionally only exposes options relevant to Excel generation.
type ExcelGenerateRequest struct {
	TemplateID string             `json:"templateId,omitempty" xml:"TemplateId,omitempty" example:"default"`
	Layout     *ExcelLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data       DataPayload        `json:"data" xml:"Data" swaggertype:"object"`
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
		Format:     DocumentFormatExcel,
		TemplateID: r.TemplateID,
		Layout:     nil,
		Data:       r.Data,
	}
	if r.Layout != nil {
		out.Layout = r.Layout.ToLayoutConfig()
	}
	return out
}

// PDFGenerateRequest is the request body for POST /generate/pdf.
// It intentionally only exposes options relevant to PDF generation.
type PDFGenerateRequest struct {
	TemplateID string           `json:"templateId,omitempty" xml:"TemplateId,omitempty" example:"default"`
	Layout     *PDFLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data       DynamicData      `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *PDFGenerateRequest) Validate() error {
	if err := r.Layout.Validate(); err != nil {
		return err
	}
	if r.Data.IsEmpty() {
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	return nil
}

func (r PDFGenerateRequest) ToDocumentRequest() DocumentRequest {
	payload := DataPayload{Default: r.Data}
	out := DocumentRequest{
		Format:     DocumentFormatPDF,
		TemplateID: r.TemplateID,
		Layout:     nil,
		Data:       payload,
	}
	if r.Layout != nil {
		out.Layout = r.Layout.ToLayoutConfig()
	}
	return out
}

// WordGenerateRequest is the request body for POST /generate/word.
// It intentionally only exposes options relevant to Word (DOCX) generation.
type WordGenerateRequest struct {
	TemplateID string            `json:"templateId,omitempty" xml:"TemplateId,omitempty" example:"default"`
	Layout     *WordLayoutConfig `json:"layout,omitempty" xml:"Layout,omitempty"`
	Data       DynamicData       `json:"data" xml:"Data" swaggertype:"object"`
}

func (r *WordGenerateRequest) Validate() error {
	if err := r.Layout.Validate(); err != nil {
		return err
	}
	if r.Data.IsEmpty() {
		return fmt.Errorf("data is required and must contain at least one item with keys")
	}
	return nil
}

func (r WordGenerateRequest) ToDocumentRequest() DocumentRequest {
	payload := DataPayload{Default: r.Data}
	out := DocumentRequest{
		Format:     DocumentFormatWord,
		TemplateID: r.TemplateID,
		Layout:     nil,
		Data:       payload,
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
}

func (c ExcelColumnConfig) Validate() error {
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
	Sheets          []SheetConfig       `json:"sheets,omitempty" xml:"Sheets>Sheet,omitempty"`
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
		Sheets:          l.Sheets,
		Header:          l.Header,
		Body:            l.Body,
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
				Format:            c.Format,
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
	Header               *StyleConfig        `json:"header,omitempty" xml:"Header,omitempty"`
	Body                 *StyleConfig        `json:"body,omitempty" xml:"Body,omitempty"`
	Footer               *FooterConfig       `json:"footer,omitempty" xml:"Footer,omitempty"`
	HeaderImage          *HeaderImageConfig  `json:"headerImage,omitempty" xml:"HeaderImage,omitempty"`
	Columns              []TableColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
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
		Header:               l.Header,
		Body:                 l.Body,
		Footer:               l.Footer,
		HeaderImage:          l.HeaderImage,
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
	PageOrientation PageOrientationEnum `json:"pageOrientation,omitempty" xml:"PageOrientation,omitempty" example:"Portrait"`
	PageMargin      *PageMarginConfig   `json:"pageMargin,omitempty" xml:"PageMargin,omitempty"`
	DefaultFont     *DefaultFontConfig  `json:"defaultFont,omitempty" xml:"DefaultFont,omitempty"`
	PageBreak       *PageBreakConfig    `json:"pageBreak,omitempty" xml:"PageBreak,omitempty"`
	Header          *StyleConfig        `json:"header,omitempty" xml:"Header,omitempty"`
	Body            *StyleConfig        `json:"body,omitempty" xml:"Body,omitempty"`
	Footer          *FooterConfig       `json:"footer,omitempty" xml:"Footer,omitempty"`
	Columns         []TableColumnConfig `json:"columns,omitempty" xml:"Columns>Column,omitempty"`
}

func (l *WordLayoutConfig) Validate() error {
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
	out := &LayoutConfig{
		PageOrientation: l.PageOrientation,
		PageMargin:      l.PageMargin,
		DefaultFont:     l.DefaultFont,
		PageBreak:       l.PageBreak,
		Header:          l.Header,
		Body:            l.Body,
		Footer:          l.Footer,
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
