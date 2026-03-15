package auth

import (
	"errors"
	"os"
	"time"

	"github.com/golang-jwt/jwt/v4"
)

var (
	ErrMissingSecret = errors.New("missing SECRETKEY environment variable")
)

func GenerateJWT(sub, name string, expSec int64, extra map[string]string) (string, error) {
	if expSec <= 0 {
		expSec = int64(100 * 365 * 24 * 60 * 60) // 100 anos
	}

	claims := jwt.MapClaims{
		"sub":  sub,
		"name": name,
		"iat":  time.Now().Unix(),
		"exp":  time.Now().Add(time.Duration(expSec) * time.Second).Unix(),
	}

	for k, v := range extra {
		claims[k] = v
	}

	secret := os.Getenv("SECRETKEY")
	if secret == "" {
		return "", ErrMissingSecret
	}

	token := jwt.NewWithClaims(jwt.SigningMethodHS256, claims)
	return token.SignedString([]byte(secret))
}
