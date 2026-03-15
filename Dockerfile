## Data2Doc Dockerfile
## Build: docker build -t data2doc .
## Run:   docker run --rm -p 8080:8080 --env-file .env data2doc

FROM golang:1.25-alpine AS builder

WORKDIR /src

RUN apk add --no-cache ca-certificates tzdata && update-ca-certificates

# Cache deps
COPY go.mod go.sum ./
RUN go mod download

# Build
COPY . .

ENV CGO_ENABLED=0
RUN go build -trimpath -ldflags="-s -w" -o /out/data2doc ./cmd


FROM alpine:3.21

RUN apk add --no-cache ca-certificates tzdata && update-ca-certificates \
	&& addgroup -S app \
	&& adduser -S -G app app

WORKDIR /app
COPY --from=builder /out/data2doc /app/data2doc

USER app

EXPOSE 8080
ENTRYPOINT ["/app/data2doc"]
