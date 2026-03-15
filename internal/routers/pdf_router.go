package routers

import (
	"Data2Doc/internal/handlers"

	"github.com/gin-gonic/gin"
)

func RegisterPDFRoutes(rg *gin.RouterGroup, generateHandler *handlers.GenerateHandler) {
	rg.POST("/generate/pdf", generateHandler.GeneratePDF)
}
