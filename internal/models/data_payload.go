package models

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

// DataPayload represents the `data` field in v2.
//
// It is backward compatible with the original shape:
//   - JSON: array of objects OR a single object
//   - XML:  <Data><Item>...</Item></Data>
//
// It also supports a multi-dataset shape required by layout.sheets:
//   - JSON: object whose values are arrays/objects, e.g. {"employees": [...], "departments": [...]}
//   - XML:  <Data><employees><Item>...</Item></employees><departments>...</departments></Data>
//
// Internals are hidden from Swagger; the OpenAPI schema documents `data` generically.
type DataPayload struct {
	Default DynamicData            `json:"-" xml:"-" swaggerignore:"true"`
	Sources map[string]DynamicData `json:"-" xml:"-" swaggerignore:"true"`
}

func (p *DataPayload) IsEmpty() bool {
	if p == nil {
		return true
	}
	if !p.Default.IsEmpty() {
		return false
	}
	for _, d := range p.Sources {
		if !d.IsEmpty() {
			return false
		}
	}
	return true
}

func (p *DataPayload) Get(source string) DynamicData {
	if p == nil {
		return DynamicData{}
	}
	src := strings.TrimSpace(source)
	if src == "" || strings.EqualFold(src, "default") {
		return p.Default
	}
	if p.Sources == nil {
		return DynamicData{}
	}
	if d, ok := p.Sources[src]; ok {
		return d
	}
	return DynamicData{}
}

func (p *DataPayload) UnmarshalJSON(b []byte) error {
	bb := bytes.TrimSpace(b)
	if len(bb) == 0 || bytes.Equal(bb, []byte("null")) {
		return nil
	}

	switch bb[0] {
	case '[':
		var d DynamicData
		if err := json.Unmarshal(bb, &d); err != nil {
			return err
		}
		p.Default = d
		return nil
	case '{':
		// Attempt to detect multi-dataset object.
		var raw map[string]json.RawMessage
		if err := json.Unmarshal(bb, &raw); err != nil {
			return err
		}
		// If any value is an array, treat as sources.
		isSources := false
		for _, v := range raw {
			vv := bytes.TrimSpace(v)
			if len(vv) == 0 {
				continue
			}
			if vv[0] == '[' {
				isSources = true
				break
			}
		}
		if isSources {
			p.Sources = make(map[string]DynamicData, len(raw))
			for k, v := range raw {
				vv := bytes.TrimSpace(v)
				if len(vv) == 0 || bytes.Equal(vv, []byte("null")) {
					continue
				}
				var d DynamicData
				if err := json.Unmarshal(vv, &d); err != nil {
					return fmt.Errorf("data.%s: %w", k, err)
				}
				p.Sources[k] = d
			}
			return nil
		}

		// Fallback: single object = one-row dataset.
		var d DynamicData
		if err := json.Unmarshal(bb, &d); err != nil {
			return err
		}
		p.Default = d
		return nil
	default:
		return fmt.Errorf("expected JSON array or object")
	}
}

func (p *DataPayload) UnmarshalXML(dec *xml.Decoder, start xml.StartElement) error {
	if p.Sources == nil {
		p.Sources = map[string]DynamicData{}
	}

	for {
		tok, err := dec.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if strings.EqualFold(t.Name.Local, "Item") {
				item, order, err := readXMLItem(dec, t)
				if err != nil {
					return err
				}
				// Append to default dataset.
				if p.Default.Items == nil {
					p.Default.Items = []map[string]any{}
				}
				p.Default.Items = append(p.Default.Items, item)
				// Maintain order.
				seen := map[string]bool{}
				for _, k := range p.Default.Order {
					seen[k] = true
				}
				for _, k := range order {
					if !seen[k] {
						p.Default.Order = append(p.Default.Order, k)
						seen[k] = true
					}
				}
				continue
			}

			// Dataset container element: <employees>...</employees>
			ds, err := readXMLDataSet(dec, t)
			if err != nil {
				return err
			}
			if !ds.IsEmpty() {
				p.Sources[t.Name.Local] = ds
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
	return nil
}

func readXMLDataSet(dec *xml.Decoder, start xml.StartElement) (DynamicData, error) {
	var d DynamicData
	seen := map[string]bool{}
	appendKey := func(k string) {
		if !seen[k] {
			seen[k] = true
			d.Order = append(d.Order, k)
		}
	}

	for {
		tok, err := dec.Token()
		if err != nil {
			return DynamicData{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if strings.EqualFold(t.Name.Local, "Item") {
				item, order, err := readXMLItem(dec, t)
				if err != nil {
					return DynamicData{}, err
				}
				for _, k := range order {
					appendKey(k)
				}
				d.Items = append(d.Items, item)
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return d, nil
			}
		}
	}
}
