package middleware

import (
	"Data2Doc/pkg/sentry"
	"fmt"
	"net/http"
	"time"

	sentrygo "github.com/getsentry/sentry-go"
	"github.com/gin-gonic/gin"
)

func SentryMiddleware() gin.HandlerFunc {
	return func(c *gin.Context) {

		c.Next()

		if len(c.Errors) > 0 {

			for _, ginErr := range c.Errors {

				sentry.WithScope(func(scope *sentrygo.Scope) {

					scope.SetTag("endpoint", c.Request.URL.Path)
					scope.SetTag("method", c.Request.Method)
					scope.SetTag("status", http.StatusText(c.Writer.Status()))
					scope.SetTag("status_code", string(rune(c.Writer.Status())))

					scope.SetRequest(c.Request)

					if c.Request.ContentLength > 0 && c.Request.ContentLength < 10000 {
						if bodyBytes, err := c.GetRawData(); err == nil {
							scope.SetExtra("request_body", string(bodyBytes))
						}
					}

					if userID, exists := c.Get("userID"); exists {
						scope.SetUser(sentrygo.User{
							ID: userID.(string),
						})
					}

					if c.Writer.Status() >= 400 {
						scope.SetLevel(sentrygo.LevelError)
					}

					sentry.CaptureException(ginErr.Err)
				})
			}

			sentry.Flush(1 * time.Second)
		} else if c.Writer.Status() >= 500 {

			sentry.WithScope(func(scope *sentrygo.Scope) {
				scope.SetTag("endpoint", c.Request.URL.Path)
				scope.SetTag("method", c.Request.Method)
				scope.SetTag("status", http.StatusText(c.Writer.Status()))
				scope.SetTag("status_code", string(rune(c.Writer.Status())))
				scope.SetLevel(sentrygo.LevelError)

				errMsg := fmt.Sprintf("Status 500 em %s %s sem erro explícito",
					c.Request.Method, c.Request.URL.Path)
				sentry.CaptureMessage(errMsg)
			})
			sentry.Flush(1 * time.Second)
		}
	}
}
