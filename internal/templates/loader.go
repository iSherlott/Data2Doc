package templates

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"Data2Doc/internal/models"
)

var (
	ErrInvalidTemplateName = errors.New("invalid template name")
	ErrTemplateNotFound    = errors.New("template not found")
)

type Loader struct {
	BaseDir string
}

func NewLoader(baseDir string) *Loader {
	if strings.TrimSpace(baseDir) == "" {
		baseDir = "templates"
	}
	return &Loader{BaseDir: baseDir}
}

func (l *Loader) Load(docType models.DocumentType, templateName string) ([]byte, string, error) {
	if strings.TrimSpace(templateName) == "" {
		return nil, "", ErrTemplateNotFound
	}

	// Prevent path traversal; only allow bare filenames.
	cleanName := filepath.Base(templateName)
	if cleanName != templateName || strings.ContainsAny(templateName, "\\/") {
		return nil, "", ErrInvalidTemplateName
	}

	path := filepath.Join(l.BaseDir, string(docType), cleanName)
	b, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, "", fmt.Errorf("%w: %s", ErrTemplateNotFound, path)
		}
		return nil, "", err
	}
	return b, path, nil
}
