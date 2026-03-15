package config

import (
	"Data2Doc/utils"

	"os/exec"
)

func Swagger() {
	err := exec.Command("go", "run", "github.com/swaggo/swag/cmd/swag@latest", "init", "--generalInfo", "cmd/main.go", "--output", "docs").Run()
	if err != nil {
		utils.FatalIf(err, "⚠️ Erro ao gerar documentação Swagger:")
	}

	utils.LogIfDevelopment("✅ Documentação Swagger atualizada com sucesso.")
}
