package config

import (
	"Data2Doc/internal/routers"
)

func RequestHTTP() {
	r := routers.NewRouter()
	_ = r.Run(":8080")
}
