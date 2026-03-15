package routers

import (
	"net/http"
	"strings"
	"time"

	"Data2Doc/internal/auth"
	"Data2Doc/internal/handlers"
	"Data2Doc/internal/service"

	"github.com/gin-contrib/cors"
	"github.com/gin-gonic/gin"

	swaggerFiles "github.com/swaggo/files"
	ginSwagger "github.com/swaggo/gin-swagger"

	docs "Data2Doc/docs"
)

func NewRouter() *gin.Engine {
	r := gin.Default()

	// Disable CORS restrictions (browser-side) by allowing all origins.
	r.Use(cors.New(cors.Config{
		// Avoid "*" so browsers allow credentials if the client sets them.
		AllowOriginFunc:  func(origin string) bool { return true },
		AllowCredentials: true,
		AllowMethods:     []string{"GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"},
		AllowHeaders: []string{
			"Origin",
			"Content-Length",
			"Content-Type",
			"Authorization",
			"Accept",
			"X-Requested-With",
			"X-Request-Id",
		},
		ExposeHeaders: []string{"Content-Disposition"},
		MaxAge:        12 * time.Hour,
	}))

	// Swagger UI
	r.Use(func(c *gin.Context) {
		// When swagger is accessed via IP/hostname, override generated Host (which may be localhost)
		// so Try-it-out uses the same origin and doesn't trigger CORS.
		if strings.HasPrefix(c.Request.URL.Path, "/swagger") {
			docs.SwaggerInfo.Host = c.Request.Host
			if c.Request.TLS != nil {
				docs.SwaggerInfo.Schemes = []string{"https"}
			} else {
				docs.SwaggerInfo.Schemes = []string{"http"}
			}
		}
		c.Next()
	})

	r.GET("/swagger", func(c *gin.Context) {
		c.Redirect(http.StatusFound, "/swagger/index.html")
	})
	r.GET("/swagger/*any", ginSwagger.WrapHandler(swaggerFiles.Handler))

	documentService := service.NewDocumentService()
	generateHandler := handlers.NewGenerateHandler(documentService)

	protected := r.Group("/")
	protected.Use(auth.AuthIdentityMiddleware())
	RegisterExcelRoutes(protected, generateHandler)
	RegisterPDFRoutes(protected, generateHandler)
	RegisterWordRoutes(protected, generateHandler)

	return r
}
