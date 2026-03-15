package main

import (
	"os"

	"Data2Doc/config"
	"Data2Doc/utils"
)

// @title       Data2Doc API
// @version     1.0
// @description Data binding to documents (excel/word/pdf).
// @host        localhost:8080
// @BasePath    /
// @securityDefinitions.apikey  BearerAuth
// @in                          header
// @name                        Authorization
// @description                 Type "Bearer {token}"

func main() {
	config.LoadEnv()

	var appMode = os.Getenv("APP_MODE")
	if appMode == "" {
		utils.PrintIn("APP_MODE not found.")
		return
	}

	config.RequestHTTP()

	utils.PrintIn("App Running!")
}
