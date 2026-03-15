package service

import (
	"errors"

	"Data2Doc/internal/models"
	"Data2Doc/utils"
)

type DataBindService struct{}

func NewDataBindService() *DataBindService {
	return &DataBindService{}
}

func (s *DataBindService) Process(req models.DataBind) error {

	if req.Id == "" {
		return errors.New("id is required")
	}

	if req.Payload == "" {
		return errors.New("payload is empty")
	}

	if req.Format == "" {
		return errors.New("format not specified")
	}

	switch req.Format {
	case "excel", "word", "pdf":
		// ok
	default:
		return errors.New("invalid format; expected excel|word|pdf")
	}

	switch req.Type {
	case models.DataBindJSON, models.DataBindXML:
		// ok
	default:
		return errors.New("invalid type; expected json|xml")
	}

	utils.PrintIn(req.Id)
	utils.PrintIn(req.Format)
	utils.PrintIn(string(req.Type))
	utils.PrintIn(req.Payload)

	// aqui futuramente envia para fila
	// queue.Publish(req)

	return nil
}
