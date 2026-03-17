package config

import (
	"Data2Doc/internal/routers"
	"fmt"
	"os"
	"strconv"
	"strings"
)

func RequestHTTP() {
	r := routers.NewRouter()
	port := strings.TrimSpace(os.Getenv("PORT"))
	if port == "" {
		port = "80"
	}
	if strings.HasPrefix(port, ":") {
		port = strings.TrimPrefix(port, ":")
	}
	if p, err := strconv.Atoi(port); err != nil || p < 1 || p > 65535 {
		port = "80"
	}
	_ = r.Run(fmt.Sprintf(":%s", port))
}
