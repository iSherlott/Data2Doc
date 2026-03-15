package models

type DocumentType string

const (
	DocumentExcel DocumentType = "excel"
	DocumentWord  DocumentType = "word"
	DocumentPDF   DocumentType = "pdf"
)

func (t DocumentType) IsValid() bool {
	switch t {
	case DocumentExcel, DocumentWord, DocumentPDF:
		return true
	default:
		return false
	}
}
