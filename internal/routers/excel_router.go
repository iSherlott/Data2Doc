package routers

import (
	"Data2Doc/internal/handlers"

	"github.com/gin-gonic/gin"
)

func RegisterExcelRoutes(rg *gin.RouterGroup, generateHandler *handlers.GenerateHandler) {
	rg.POST("/generate/excel", generateHandler.GenerateExcel)
}
