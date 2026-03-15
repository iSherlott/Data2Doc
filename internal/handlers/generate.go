package handlers

import (
	"fmt"
	"net/http"
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

func baseNameFromRequest(c *gin.Context) string {
	baseName := strings.TrimSpace(c.Query("id"))
	if baseName == "" {
		baseName = strings.TrimSpace(c.GetHeader("X-Request-Id"))
	}
	if baseName == "" {
		baseName = "document"
	}
	return baseName
}

// GenerateExcel godoc
// @Summary      Generate an Excel document
// @Description  Generates an XLSX document. Supports JSON and XML.
// @Tags         documents
// @Security     BearerAuth
// @Accept       json
// @Accept       xml
// @Produce      application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
// @Param        id       query    string                    false  "base filename (without extension)"
// @Param        request  body     models.ExcelGenerateRequest true   "Excel request"
// @Success      200      {file}   file
// @Failure      400      {object} map[string]string
// @Failure      401      {object} map[string]string
// @Failure      404      {object} map[string]string
// @Router       /generate/excel [post]
func (h *GenerateHandler) GenerateExcel(c *gin.Context) {
	var req models.ExcelGenerateRequest
	if err := c.ShouldBind(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}
	if err := req.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	docReq := req.ToDocumentRequest()
	if err := docReq.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	gen, err := h.documentService.GenerateV2(docReq)
	if err != nil {
		status := http.StatusBadRequest
		if service.IsTemplateNotFound(err) {
			status = http.StatusNotFound
		}
		c.JSON(status, gin.H{"error": err.Error()})
		return
	}

	filename := baseNameFromRequest(c) + ".xlsx"
	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", filename))
	c.Data(http.StatusOK, gen.ContentType, gen.Bytes)
}

// GeneratePDF godoc
// @Summary      Generate a PDF document
// @Description  Generates a PDF document. Supports JSON and XML.
// @Tags         documents
// @Security     BearerAuth
// @Accept       json
// @Accept       xml
// @Produce      application/pdf
// @Param        id       query    string                  false  "base filename (without extension)"
// @Param        request  body     models.PDFGenerateRequest true   "PDF request"
// @Success      200      {file}   file
// @Failure      400      {object} map[string]string
// @Failure      401      {object} map[string]string
// @Failure      404      {object} map[string]string
// @Router       /generate/pdf [post]
func (h *GenerateHandler) GeneratePDF(c *gin.Context) {
	var req models.PDFGenerateRequest
	if err := c.ShouldBind(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}
	if err := req.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	docReq := req.ToDocumentRequest()
	if err := docReq.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	gen, err := h.documentService.GenerateV2(docReq)
	if err != nil {
		status := http.StatusBadRequest
		if service.IsTemplateNotFound(err) {
			status = http.StatusNotFound
		}
		c.JSON(status, gin.H{"error": err.Error()})
		return
	}

	filename := baseNameFromRequest(c) + ".pdf"
	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", filename))
	c.Data(http.StatusOK, gen.ContentType, gen.Bytes)
}

// GenerateWord godoc
// @Summary      Generate a Word (DOCX) document
// @Description  Generates a Word (DOCX) document. Supports JSON and XML.
// @Tags         documents
// @Security     BearerAuth
// @Accept       json
// @Accept       xml
// @Produce      application/vnd.openxmlformats-officedocument.wordprocessingml.document
// @Param        id       query    string                   false  "base filename (without extension)"
// @Param        request  body     models.WordGenerateRequest true   "Word request"
// @Success      200      {file}   file
// @Failure      400      {object} map[string]string
// @Failure      401      {object} map[string]string
// @Failure      404      {object} map[string]string
// @Router       /generate/word [post]
func (h *GenerateHandler) GenerateWord(c *gin.Context) {
	var req models.WordGenerateRequest
	if err := c.ShouldBind(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}
	if err := req.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	docReq := req.ToDocumentRequest()
	if err := docReq.Validate(); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	gen, err := h.documentService.GenerateV2(docReq)
	if err != nil {
		status := http.StatusBadRequest
		if service.IsTemplateNotFound(err) {
			status = http.StatusNotFound
		}
		c.JSON(status, gin.H{"error": err.Error()})
		return
	}

	filename := baseNameFromRequest(c) + ".docx"
	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", filename))
	c.Data(http.StatusOK, gen.ContentType, gen.Bytes)
}
