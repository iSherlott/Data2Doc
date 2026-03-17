package routers

import (
	"net/http"
	"os"
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

	isProduction := strings.EqualFold(strings.TrimSpace(os.Getenv("ENV")), "Production")

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

	// Swagger UI (disabled in production)
	if !isProduction {
		r.Use(func(c *gin.Context) {
			// When swagger is accessed via IP/hostname, override generated Host (which may be localhost)
			// so Try-it-out uses the same origin and doesn't trigger CORS.
			if strings.HasPrefix(c.Request.URL.Path, "/swagger") {
				// Avoid stale swagger specs due to browser/proxy caching.
				c.Header("Cache-Control", "no-store")
				c.Header("Pragma", "no-cache")
				c.Header("Expires", "0")

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
	}

	// Public metadata endpoint (no auth).
	metaHandler := handlers.NewMetaHandler()
	r.GET("/meta", metaHandler.GetMeta)

	// Local JWT mode helper: issue tokens using USER/PASS from env.
	// Only enabled when AUTH_JWKS_URL is not set.
	if strings.TrimSpace(os.Getenv("AUTH_JWKS_URL")) == "" {
		authHandler := handlers.NewAuthHandler()
		r.POST("/auth/token", authHandler.IssueToken)
	}

	documentService := service.NewDocumentService()
	generateHandler := handlers.NewGenerateHandler(documentService)

	protected := r.Group("/")
	if strings.TrimSpace(os.Getenv("AUTH_JWKS_URL")) != "" {
		protected.Use(auth.AuthIdentityMiddleware())
	} else {
		protected.Use(auth.AuthMiddleware())
	}
	RegisterExcelRoutes(protected, generateHandler)
	RegisterPDFRoutes(protected, generateHandler)
	RegisterWordRoutes(protected, generateHandler)

	return r
}
