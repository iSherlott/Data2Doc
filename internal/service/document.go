package service

import (
	"errors"

	"Data2Doc/internal/templates"
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
