package handlers

import (
	"os"
	"runtime/debug"
	"strings"
	"time"

	"github.com/gin-gonic/gin"
)

type MetaHandler struct{}

type MetaResponse struct {
	Name       string            `json:"name"`
	Version    string            `json:"version"`
	Build      map[string]string `json:"build,omitempty"`
	Env        string            `json:"env"`
	TimeUTC    string            `json:"timeUtc"`
	AuthMode   string            `json:"authMode"`
	Port       string            `json:"port"`
	HasSentry  bool              `json:"hasSentry"`
	HasJWKSURL bool              `json:"hasJwksUrl"`
}

func NewMetaHandler() *MetaHandler { return &MetaHandler{} }

// GetMeta godoc
// @Summary      Service metadata
// @Description  Returns service metadata and current auth mode (does not expose secrets).
// @Tags         meta
// @Produce      json
// @Success      200  {object}  handlers.MetaResponse
// @Router       /meta [get]
func (h *MetaHandler) GetMeta(c *gin.Context) {
	name := strings.TrimSpace(os.Getenv("APP_NAME"))
	if name == "" {
		name = "Data2Doc"
	}

	version := strings.TrimSpace(os.Getenv("APP_VERSION"))
	if version == "" {
		version = "dev"
	}

	build := map[string]string{}
	if bi, ok := debug.ReadBuildInfo(); ok {
		build["goVersion"] = bi.GoVersion
		for _, s := range bi.Settings {
			if s.Key == "vcs.revision" {
				build["revision"] = s.Value
			}
			if s.Key == "vcs.time" {
				build["revisionTime"] = s.Value
			}
			if s.Key == "vcs.modified" {
				build["modified"] = s.Value
			}
		}
	}
	if len(build) == 0 {
		build = nil
	}

	env := strings.TrimSpace(os.Getenv("ENV"))
	jwks := strings.TrimSpace(os.Getenv("AUTH_JWKS_URL"))
	authMode := "local_jwt"
	if jwks != "" {
		authMode = "jwks"
	}

	port := strings.TrimSpace(os.Getenv("PORT"))
	if port == "" {
		port = "80"
	}

	sentryDsn := strings.TrimSpace(os.Getenv("SENTRY_DSN"))

	c.JSON(200, MetaResponse{
		Name:       name,
		Version:    version,
		Build:      build,
		Env:        env,
		TimeUTC:    time.Now().UTC().Format(time.RFC3339),
		AuthMode:   authMode,
		Port:       port,
		HasSentry:  sentryDsn != "",
		HasJWKSURL: jwks != "",
	})
}
