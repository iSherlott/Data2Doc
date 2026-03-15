package utils

import (
	"fmt"
	"log"
	"os"
)

func LogIfDevelopment(format string, v ...interface{}) {
	if os.Getenv("ENV") == "Development" {
		log.Printf(format, v...)
	}
}

func PrintIn(format string, v ...interface{}) {
	if os.Getenv("ENV") == "Development" {
		fmt.Printf(format+"\n", v...)
	}
}

func FatalIf(err error, format string, v ...interface{}) {
	if err != nil {

		log.Fatalf(format, v...)
	}
}

func LogError(err error, format string, v ...interface{}) {
	if err != nil {

		log.Printf(format+": %v", append(v, err)...)
	}
}
