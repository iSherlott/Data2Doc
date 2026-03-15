package routers

import (
	"Data2Doc/internal/handlers"

	"github.com/gin-gonic/gin"
)

func RegisterWordRoutes(rg *gin.RouterGroup, generateHandler *handlers.GenerateHandler) {
	rg.POST("/generate/word", generateHandler.GenerateWord)
}
