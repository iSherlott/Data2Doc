package auth

import (
	"errors"
	"fmt"
	"net/http"
	"os"
	"strings"
	"sync"
	"time"

	"github.com/MicahParks/keyfunc"
	"github.com/gin-gonic/gin"
	"github.com/golang-jwt/jwt/v4"
)

const (
	defaultJWKSURL = "https://connect-staging.fi-group.com/identity/.well-known/openid-configuration/jwks"
)

var (
	jwksOnce sync.Once
	jwksInst *keyfunc.JWKS
	jwksErr  error
)

func getJWKS() (*keyfunc.JWKS, error) {
	jwksOnce.Do(func() {
		jwksURL := strings.TrimSpace(os.Getenv("AUTH_JWKS_URL"))
		if jwksURL == "" {
			jwksURL = defaultJWKSURL
		}

		jwksInst, jwksErr = keyfunc.Get(jwksURL, keyfunc.Options{
			RefreshInterval: time.Hour,
			RefreshErrorHandler: func(err error) {
				fmt.Printf("Erro ao atualizar JWKS: %v\n", err)
			},
		})
	})

	return jwksInst, jwksErr
}

func AuthIdentityMiddleware() gin.HandlerFunc {
	return func(c *gin.Context) {
		jwks, err := getJWKS()
		if err != nil {
			c.AbortWithStatusJSON(http.StatusInternalServerError, gin.H{"error": "failed to load JWKS"})
			return
		}

		authHeader := c.GetHeader("Authorization")
		if authHeader == "" {
			c.AbortWithStatusJSON(http.StatusUnauthorized, gin.H{"error": "Authorization header missing"})
			return
		}

		parts := strings.SplitN(authHeader, " ", 2)
		if len(parts) != 2 || parts[0] != "Bearer" {
			c.AbortWithStatusJSON(http.StatusUnauthorized, gin.H{"error": "Invalid Authorization header"})
			return
		}

		tokenString := parts[1]

		token, err := jwt.Parse(tokenString, jwks.Keyfunc)
		if err != nil || !token.Valid {
			c.AbortWithStatusJSON(http.StatusUnauthorized, gin.H{"error": "Invalid or expired token"})
			return
		}

		claims, ok := token.Claims.(jwt.MapClaims)
		if !ok {
			c.AbortWithStatusJSON(http.StatusUnauthorized, gin.H{"error": "Invalid token claims"})
			return
		}

		expectedClientID := strings.TrimSpace(os.Getenv("AUTH_EXPECTED_CLIENT_ID"))
		if expectedClientID != "" {
			clientID, ok := claims["client_id"].(string)
			if !ok || clientID != expectedClientID {
				c.AbortWithStatusJSON(http.StatusUnauthorized, gin.H{"error": "Invalid client_id"})
				return
			}
		}

		// Store claims for downstream handlers.
		c.Set("auth.claims", claims)
		if sub, ok := claims["sub"].(string); ok {
			c.Set("auth.sub", sub)
		}
		if azp, ok := claims["azp"].(string); ok {
			c.Set("auth.azp", azp)
		}

		c.Next()
	}
}

// Useful when you want to unit-test validation paths.
var ErrAuthConfig = errors.New("auth config error")
