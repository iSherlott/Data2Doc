package sentry

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/getsentry/sentry-go"
)

func Setup() error {
	dsn := os.Getenv("SENTRY_DSN")
	if dsn == "" {
		log.Println("Aviso: SENTRY_DSN não está configurado. O monitoramento de erros está desativado.")
		return nil
	}

	environment := os.Getenv("ENV")
	if environment == "" {
		environment = "production"
	}

	err := sentry.Init(sentry.ClientOptions{
		Dsn:              dsn,
		Environment:      environment,
		Release:          os.Getenv("APP_VERSION"),
		TracesSampleRate: 1.0,
		Debug:            environment == "Development",
		BeforeSend: func(event *sentry.Event, hint *sentry.EventHint) *sentry.Event {

			if environment == "Development" {
				log.Printf("Enviando evento para o Sentry: %s - %s", event.Level, event.Message)
			}
			return event
		},
	})

	if err != nil {
		return fmt.Errorf("erro ao inicializar Sentry: %w", err)
	}

	sentry.Flush(2 * time.Second)

	log.Println("Sentry inicializado com sucesso")
	return nil
}

func CaptureException(err error) {
	if err == nil {
		return
	}

	sentry.CaptureException(err)
}

func CaptureMessage(message string) {
	sentry.CaptureMessage(message)
}

func WithScope(f func(scope *sentry.Scope)) {
	sentry.WithScope(f)
}

func ConfigureScope(f func(scope *sentry.Scope)) {
	sentry.ConfigureScope(f)
}

func Flush(timeout time.Duration) {
	sentry.Flush(timeout)
}

func CaptureError(err error, tags map[string]string, extras map[string]interface{}) {
	if err == nil {
		return
	}

	sentry.WithScope(func(scope *sentry.Scope) {

		for key, value := range tags {
			scope.SetTag(key, value)
		}

		for key, value := range extras {
			scope.SetExtra(key, value)
		}

		sentry.CaptureException(err)
	})
}

func RecoverPanic() {
	if err := recover(); err != nil {
		sentry.CurrentHub().Recover(err)
		sentry.Flush(2 * time.Second)

		panic(err)
	}
}
