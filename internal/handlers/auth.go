package handlers

import (
	"crypto/subtle"
	"net/http"
	"os"
	"strconv"
	"strings"

	"Data2Doc/internal/auth"

	"github.com/gin-gonic/gin"
)

type AuthHandler struct{}

func NewAuthHandler() *AuthHandler { return &AuthHandler{} }

type tokenRequest struct {
	User string `json:"user" binding:"required"`
	Pass string `json:"pass" binding:"required"`
}

// IssueToken godoc
// @Summary      Issue a JWT token (local mode)
// @Description  Issues a local HMAC JWT when AUTH_JWKS_URL is not configured.
// @Tags         auth
// @Accept       json
// @Produce      json
// @Param        request  body     tokenRequest true "Credentials"
// @Success      200      {object} map[string]any
// @Failure      400      {object} map[string]string
// @Failure      401      {object} map[string]string
// @Failure      500      {object} map[string]string
// @Router       /auth/token [post]
func (h *AuthHandler) IssueToken(c *gin.Context) {
	if strings.TrimSpace(os.Getenv("AUTH_JWKS_URL")) != "" {
		c.JSON(http.StatusBadRequest, gin.H{"error": "auth/token is disabled when AUTH_JWKS_URL is set"})
		return
	}

	expectedUser := strings.TrimSpace(os.Getenv("AUTH_USER"))
	if expectedUser == "" {
		expectedUser = strings.TrimSpace(os.Getenv("USER"))
	}
	expectedPass := strings.TrimSpace(os.Getenv("AUTH_PASS"))
	if expectedPass == "" {
		expectedPass = strings.TrimSpace(os.Getenv("PASS"))
	}

	if expectedUser == "" || expectedPass == "" {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "missing USER/PASS in environment"})
		return
	}

	var req tokenRequest
	if err := c.ShouldBindJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	userOK := subtle.ConstantTimeCompare([]byte(req.User), []byte(expectedUser)) == 1
	passOK := subtle.ConstantTimeCompare([]byte(req.Pass), []byte(expectedPass)) == 1
	if !userOK || !passOK {
		c.JSON(http.StatusUnauthorized, gin.H{"error": "invalid credentials"})
		return
	}

	expSec := int64(0)
	if v := strings.TrimSpace(os.Getenv("JWT_EXP_SEC")); v != "" {
		if n, err := strconv.ParseInt(v, 10, 64); err == nil {
			expSec = n
		}
	}

	tok, err := auth.GenerateJWT(expectedUser, expectedUser, expSec, nil)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	resp := gin.H{
		"access_token": tok,
		"token_type":   "Bearer",
	}
	if expSec > 0 {
		resp["expires_in"] = expSec
	}
	c.JSON(http.StatusOK, resp)
}
