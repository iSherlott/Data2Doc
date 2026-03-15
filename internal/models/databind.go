package models

type DataBindType string

const (
	DataBindJSON DataBindType = "json"
	DataBindXML  DataBindType = "xml"
)

type DataBind struct {
	Id      string       `json:"id" xml:"id"`
	Type    DataBindType `json:"type" xml:"type"`
	Payload string       `json:"payload" xml:"payload"`
	Format  string       `json:"format" xml:"format"`
}
