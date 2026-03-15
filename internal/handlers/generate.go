package handlers

import (
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"

	"Data2Doc/internal/models"
	"Data2Doc/internal/service"

	"github.com/gin-gonic/gin"
)

type GenerateHandler struct {
	documentService *service.DocumentService
}

func NewGenerateHandler(documentService *service.DocumentService) *GenerateHandler {
	return &GenerateHandler{documentService: documentService}
}

// GenerateDocument godoc
// @Summary      Generate a document
// @Description  Generates a document in excel, word or pdf.
// @Tags         documents
// @Security     BearerAuth
// @Accept       json
// @Produce      application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
// @Produce      application/vnd.openxmlformats-officedocument.wordprocessingml.document
// @Produce      application/pdf
// @Param        type      path     string          true   "excel|word|pdf"
// @Param        template  query    string          false  "template filename placed under templates/<type>/"
// @Param        id        query    string          false  "base filename (without extension)"
// @Param        payload   body     []map[string]interface{}  true   "JSON payload (array of objects)."
// @Success      200       {file}   file
// @Failure      400       {object} map[string]string
// @Failure      401       {object} map[string]string
// @Failure      404       {object} map[string]string
// @Router       /generate/{type} [post]
func (h *GenerateHandler) Generate(c *gin.Context) {
	docTypeStr := strings.ToLower(strings.TrimSpace(c.Param("type")))
	docType := models.DocumentType(docTypeStr)
	if !docType.IsValid() {
		c.JSON(http.StatusBadRequest, gin.H{"error": "invalid type; expected excel|word|pdf"})
		return
	}

	templateName := strings.TrimSpace(c.Query("template"))
	baseName := strings.TrimSpace(c.Query("id"))
	if baseName == "" {
		baseName = strings.TrimSpace(c.GetHeader("X-Request-Id"))
	}
	if baseName == "" {
		baseName = "document"
	}

	body, err := io.ReadAll(c.Request.Body)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "failed to read body"})
		return
	}
	if strings.TrimSpace(string(body)) == "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "payload is empty"})
		return
	}

	filename, contentType, b, err := h.documentService.GenerateFromPayload(docType, baseName, templateName, body)
	if err != nil {
		status := http.StatusBadRequest
		if service.IsTemplateNotFound(err) {
			status = http.StatusNotFound
		}

		resp := gin.H{"error": err.Error()}
		// Extra help for the most common Swagger mistake: sending [{}] or {}.
		if errors.Is(err, service.ErrEmptyPayload) {
			resp["hint"] = "Swagger often defaults to [{}]. Replace it with real data, e.g. [{\"name\":\"pedro\",\"age\":20}]"
		}
		if os.Getenv("ENV") == "Development" {
			resp["debug.contentType"] = c.GetHeader("Content-Type")
			trim := strings.TrimSpace(string(body))
			if len(trim) > 200 {
				trim = trim[:200] + "..."
			}
			resp["debug.bodyPreview"] = trim
		}

		c.JSON(status, resp)
		return
	}

	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", filename))
	c.Data(http.StatusOK, contentType, b)
}
