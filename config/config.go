package config

import (
	"github.com/joho/godotenv"
)

func LoadEnv() {
	// Use Overload so .env values take precedence over OS env vars.
	// This matters on Windows where USER is commonly pre-defined.
	_ = godotenv.Overload()
}
