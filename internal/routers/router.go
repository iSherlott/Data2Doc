package routers

import (
	"encoding/json"
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
	"github.com/swaggo/swag"

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

				// Azure Container Apps / reverse proxies typically terminate TLS before the app,
				// so rely on forwarded headers when available.
				host := strings.TrimSpace(c.GetHeader("X-Forwarded-Host"))
				if host == "" {
					host = c.Request.Host
				}
				docs.SwaggerInfo.Host = host

				proto := strings.TrimSpace(c.GetHeader("X-Forwarded-Proto"))
				if proto == "" {
					if c.Request.TLS != nil {
						proto = "https"
					} else {
						proto = "http"
					}
				}
				proto = strings.ToLower(proto)
				if proto != "https" {
					proto = "http"
				}
				docs.SwaggerInfo.Schemes = []string{proto}
			}
			c.Next()
		})

		r.GET("/swagger", func(c *gin.Context) {
			c.Redirect(http.StatusFound, "/swagger/index.html")
		})

		swaggerUIHandler := ginSwagger.WrapHandler(swaggerFiles.Handler)
		r.GET("/swagger/*any", func(c *gin.Context) {
			// gin doesn't allow registering both /swagger/doc.json and /swagger/*any.
			// So we intercept doc.json here to serve a filtered spec.
			if c.Param("any") == "/doc.json" {
				// Avoid stale swagger specs due to browser/proxy caching.
				c.Header("Cache-Control", "no-store")
				c.Header("Pragma", "no-cache")
				c.Header("Expires", "0")

				raw, err := swag.ReadDoc()
				if err != nil {
					c.JSON(http.StatusInternalServerError, gin.H{"error": "failed to load swagger spec"})
					return
				}
				var spec map[string]any
				if err := json.Unmarshal([]byte(raw), &spec); err != nil {
					c.Data(http.StatusOK, "application/json; charset=utf-8", []byte(raw))
					return
				}

				if strings.TrimSpace(os.Getenv("AUTH_JWKS_URL")) != "" {
					if paths, ok := spec["paths"].(map[string]any); ok {
						delete(paths, "/auth/token")
					}
					// Best-effort cleanup of request model definition.
					if defs, ok := spec["definitions"].(map[string]any); ok {
						delete(defs, "handlers.tokenRequest")
					}
				}

				out, err := json.Marshal(spec)
				if err != nil {
					c.Data(http.StatusOK, "application/json; charset=utf-8", []byte(raw))
					return
				}
				c.Data(http.StatusOK, "application/json; charset=utf-8", out)
				return
			}

			swaggerUIHandler(c)
		})
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
