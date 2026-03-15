package models

import (
	"encoding/json"
	"fmt"
	"strings"
)

type DocumentFormatEnum string

const (
	DocumentFormatExcel DocumentFormatEnum = "excel"
	DocumentFormatPDF   DocumentFormatEnum = "pdf"
	DocumentFormatWord  DocumentFormatEnum = "word"
	DocumentFormatCSV   DocumentFormatEnum = "csv"
)

func (e DocumentFormatEnum) IsValid() bool {
	switch e {
	case DocumentFormatExcel, DocumentFormatPDF, DocumentFormatWord, DocumentFormatCSV:
		return true
	default:
		return false
	}
}

func (e *DocumentFormatEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *DocumentFormatEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *DocumentFormatEnum) fromString(s string) error {
	v := DocumentFormatEnum(strings.ToLower(strings.TrimSpace(s)))
	if !v.IsValid() {
		return fmt.Errorf("invalid format '%s' (expected: excel|pdf|word|csv)", s)
	}
	*e = v
	return nil
}

type PageOrientationEnum string

const (
	PagePortrait  PageOrientationEnum = "Portrait"
	PageLandscape PageOrientationEnum = "Landscape"
)

func (e PageOrientationEnum) IsValid() bool {
	switch e {
	case PagePortrait, PageLandscape:
		return true
	default:
		return false
	}
}

func (e *PageOrientationEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *PageOrientationEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *PageOrientationEnum) fromString(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		*e = ""
		return nil
	}
	switch strings.ToLower(v) {
	case "portrait":
		*e = PagePortrait
	case "landscape":
		*e = PageLandscape
	default:
		return fmt.Errorf("invalid pageOrientation '%s' (expected: Portrait|Landscape)", s)
	}
	return nil
}

type FontFamilyEnum string

const (
	FontCalibri       FontFamilyEnum = "Calibri"
	FontArial         FontFamilyEnum = "Arial"
	FontTimesNewRoman FontFamilyEnum = "TimesNewRoman"
	FontHelvetica     FontFamilyEnum = "Helvetica"
)

func (e FontFamilyEnum) IsValid() bool {
	switch e {
	case FontCalibri, FontArial, FontTimesNewRoman, FontHelvetica:
		return true
	default:
		return false
	}
}

func (e *FontFamilyEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *FontFamilyEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *FontFamilyEnum) fromString(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		*e = ""
		return nil
	}
	// Accept case-insensitively.
	switch strings.ToLower(v) {
	case "calibri":
		*e = FontCalibri
	case "arial":
		*e = FontArial
	case "timesnewroman", "times new roman":
		*e = FontTimesNewRoman
	case "helvetica":
		*e = FontHelvetica
	default:
		return fmt.Errorf("invalid fontFamily '%s' (expected: Calibri|Arial|TimesNewRoman|Helvetica)", s)
	}
	return nil
}

type VerticalAlignmentEnum string

const (
	VAlignTop    VerticalAlignmentEnum = "Top"
	VAlignMiddle VerticalAlignmentEnum = "Middle"
	VAlignBottom VerticalAlignmentEnum = "Bottom"
)

func (e VerticalAlignmentEnum) IsValid() bool {
	switch e {
	case VAlignTop, VAlignMiddle, VAlignBottom:
		return true
	default:
		return false
	}
}

func (e *VerticalAlignmentEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *VerticalAlignmentEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *VerticalAlignmentEnum) fromString(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		*e = ""
		return nil
	}
	switch strings.ToLower(v) {
	case "top":
		*e = VAlignTop
	case "middle", "center", "centre":
		*e = VAlignMiddle
	case "bottom":
		*e = VAlignBottom
	default:
		return fmt.Errorf("invalid verticalAlignment '%s' (expected: Top|Middle|Bottom)", s)
	}
	return nil
}

type AlignmentEnum string

const (
	AlignLeft    AlignmentEnum = "left"
	AlignCenter  AlignmentEnum = "center"
	AlignRight   AlignmentEnum = "right"
	AlignJustify AlignmentEnum = "justify"
)

func (e AlignmentEnum) IsValid() bool {
	switch e {
	case AlignLeft, AlignCenter, AlignRight, AlignJustify:
		return true
	default:
		return false
	}
}

func (e *AlignmentEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *AlignmentEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *AlignmentEnum) fromString(s string) error {
	v := AlignmentEnum(strings.ToLower(strings.TrimSpace(s)))
	if v == "" {
		*e = ""
		return nil
	}
	if !v.IsValid() {
		return fmt.Errorf("invalid alignment '%s' (expected: left|center|right|justify)", s)
	}
	*e = v
	return nil
}

type ImagePositionEnum string

const (
	ImageTopLeft   ImagePositionEnum = "top-left"
	ImageTopCenter ImagePositionEnum = "top-center"
	ImageTopRight  ImagePositionEnum = "top-right"
)

func (e ImagePositionEnum) IsValid() bool {
	switch e {
	case ImageTopLeft, ImageTopCenter, ImageTopRight:
		return true
	default:
		return false
	}
}

func (e *ImagePositionEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *ImagePositionEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *ImagePositionEnum) fromString(s string) error {
	v := ImagePositionEnum(strings.ToLower(strings.TrimSpace(s)))
	if v == "" {
		*e = ""
		return nil
	}
	if !v.IsValid() {
		return fmt.Errorf("invalid image position '%s' (expected: top-left|top-center|top-right)", s)
	}
	*e = v
	return nil
}

type ImageFitModeEnum string

const (
	ImageFitContain ImageFitModeEnum = "Contain"
	ImageFitCover   ImageFitModeEnum = "Cover"
	ImageFitStretch ImageFitModeEnum = "Stretch"
	ImageFitCenter  ImageFitModeEnum = "Center"
)

func (e ImageFitModeEnum) IsValid() bool {
	switch e {
	case ImageFitContain, ImageFitCover, ImageFitStretch, ImageFitCenter:
		return true
	default:
		return false
	}
}

func (e *ImageFitModeEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *ImageFitModeEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *ImageFitModeEnum) fromString(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		*e = ""
		return nil
	}
	switch strings.ToLower(v) {
	case "contain":
		*e = ImageFitContain
	case "cover":
		*e = ImageFitCover
	case "stretch":
		*e = ImageFitStretch
	case "center", "centre":
		*e = ImageFitCenter
	default:
		return fmt.Errorf("invalid image fitMode '%s' (expected: Contain|Cover|Stretch|Center)", s)
	}
	return nil
}

type ColumnAlignmentEnum string

const (
	ColAlignLeft   ColumnAlignmentEnum = "left"
	ColAlignCenter ColumnAlignmentEnum = "center"
	ColAlignRight  ColumnAlignmentEnum = "right"
)

func (e ColumnAlignmentEnum) IsValid() bool {
	switch e {
	case ColAlignLeft, ColAlignCenter, ColAlignRight:
		return true
	default:
		return false
	}
}

func (e *ColumnAlignmentEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *ColumnAlignmentEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *ColumnAlignmentEnum) fromString(s string) error {
	v := ColumnAlignmentEnum(strings.ToLower(strings.TrimSpace(s)))
	if v == "" {
		*e = ""
		return nil
	}
	if !v.IsValid() {
		return fmt.Errorf("invalid column alignment '%s' (expected: left|center|right)", s)
	}
	*e = v
	return nil
}

type PageNumberFormatEnum string

const (
	PageNumArabic      PageNumberFormatEnum = "Arabic"
	PageNumRoman       PageNumberFormatEnum = "Roman"
	PageNumRomanUpper  PageNumberFormatEnum = "RomanUpper"
	PageNumTextPageNum PageNumberFormatEnum = "TextPageNumber"
)

func (e PageNumberFormatEnum) IsValid() bool {
	switch e {
	case PageNumArabic, PageNumRoman, PageNumRomanUpper, PageNumTextPageNum:
		return true
	default:
		return false
	}
}

func (e *PageNumberFormatEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *PageNumberFormatEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *PageNumberFormatEnum) fromString(s string) error {
	v := strings.TrimSpace(s)
	if v == "" {
		*e = ""
		return nil
	}
	switch strings.ToLower(v) {
	case "arabic":
		*e = PageNumArabic
	case "roman":
		*e = PageNumRoman
	case "romanupper", "roman_upper", "roman-upper":
		*e = PageNumRomanUpper
	case "textpagenumber", "text_page_number", "text-page-number":
		*e = PageNumTextPageNum
	default:
		return fmt.Errorf("invalid pageNumber.format '%s' (expected: Arabic|Roman|RomanUpper|TextPageNumber)", s)
	}
	return nil
}

type ColumnFormatEnum string

const (
	ColFormatCurrency   ColumnFormatEnum = "currency"
	ColFormatDate       ColumnFormatEnum = "date"
	ColFormatDateTime   ColumnFormatEnum = "datetime"
	ColFormatPercentage ColumnFormatEnum = "percentage"
	ColFormatNumber     ColumnFormatEnum = "number"
)

func (e ColumnFormatEnum) IsValid() bool {
	switch e {
	case "":
		return true
	case ColFormatCurrency, ColFormatDate, ColFormatDateTime, ColFormatPercentage, ColFormatNumber:
		return true
	default:
		return false
	}
}

func (e *ColumnFormatEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	return e.fromString(s)
}

func (e *ColumnFormatEnum) UnmarshalText(text []byte) error {
	return e.fromString(string(text))
}

func (e *ColumnFormatEnum) fromString(s string) error {
	v := strings.ToLower(strings.TrimSpace(s))
	if v == "" {
		*e = ""
		return nil
	}
	cf := ColumnFormatEnum(v)
	if !cf.IsValid() {
		return fmt.Errorf("invalid column.format '%s' (expected: currency|date|datetime|percentage|number)", s)
	}
	*e = cf
	return nil
}

type RendererTypeEnum string

const (
	RendererTable     RendererTypeEnum = "table"
	RendererParagraph RendererTypeEnum = "paragraph"
	RendererImage     RendererTypeEnum = "image"
)

func (e RendererTypeEnum) IsValid() bool {
	switch e {
	case RendererTable, RendererParagraph, RendererImage:
		return true
	default:
		return false
	}
}

func (e *RendererTypeEnum) UnmarshalJSON(b []byte) error {
	var s string
	if err := json.Unmarshal(b, &s); err != nil {
		return err
	}
	v := RendererTypeEnum(strings.ToLower(strings.TrimSpace(s)))
	if v == "" {
		*e = ""
		return nil
	}
	if !v.IsValid() {
		return fmt.Errorf("invalid rendererType '%s' (expected: table|paragraph|image)", s)
	}
	*e = v
	return nil
}

func (e *RendererTypeEnum) UnmarshalText(text []byte) error {
	return e.UnmarshalJSON([]byte("\"" + string(text) + "\""))
}
