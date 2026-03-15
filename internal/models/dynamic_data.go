package models

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// DynamicData supports JSON (array/object) and a simple XML representation:
// <Data>
//
//	<Item><name>John</name><age>30</age></Item>
//	<Item><name>Mary</name></Item>
//
// </Data>
//
// Items are stored as maps; Order preserves a stable column order based on first
// time a key appears while parsing.
type DynamicData struct {
	Items []map[string]any `json:"-" xml:"-" swaggerignore:"true"`
	Order []string         `json:"-" xml:"-" swaggerignore:"true"`
}

func (d *DynamicData) IsEmpty() bool {
	if d == nil {
		return true
	}
	if len(d.Items) == 0 {
		return true
	}
	// Consider empty if all items are empty objects.
	for _, it := range d.Items {
		if len(it) > 0 {
			return false
		}
	}
	return true
}

func (d *DynamicData) UnmarshalJSON(b []byte) error {
	dec := json.NewDecoder(bytes.NewReader(b))
	dec.UseNumber()

	tok, err := dec.Token()
	if err != nil {
		return err
	}

	seen := map[string]bool{}
	appendKey := func(k string) {
		if !seen[k] {
			seen[k] = true
			d.Order = append(d.Order, k)
		}
	}

	switch t := tok.(type) {
	case json.Delim:
		switch t {
		case '[':
			for dec.More() {
				item, order, err := decodeOrderedObject(dec)
				if err != nil {
					return err
				}
				for _, k := range order {
					appendKey(k)
				}
				d.Items = append(d.Items, item)
			}
			if _, err := dec.Token(); err != nil { // ']'
				return err
			}
			return nil
		case '{':
			// Single object becomes one-item array. We already consumed '{'.
			item, order, err := decodeOrderedObjectOpened(dec)
			if err != nil {
				return err
			}
			for _, k := range order {
				appendKey(k)
			}
			d.Items = []map[string]any{item}
			return nil
		default:
			return fmt.Errorf("expected JSON array or object")
		}
	default:
		return fmt.Errorf("expected JSON array or object")
	}
}

func decodeOrderedObjectOpened(dec *json.Decoder) (map[string]any, []string, error) {
	obj := map[string]any{}
	order := make([]string, 0, 8)
	for dec.More() {
		keyTok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		key, ok := keyTok.(string)
		if !ok {
			return nil, nil, fmt.Errorf("expected string key")
		}
		order = append(order, key)
		val, err := decodeAny(dec)
		if err != nil {
			return nil, nil, err
		}
		obj[key] = val
	}
	if _, err := dec.Token(); err != nil { // '}'
		return nil, nil, err
	}
	return obj, order, nil
}

func decodeOrderedObject(dec *json.Decoder) (map[string]any, []string, error) {
	// Expect '{'
	tok, err := dec.Token()
	if err != nil {
		return nil, nil, err
	}
	delim, ok := tok.(json.Delim)
	if !ok || delim != '{' {
		return nil, nil, fmt.Errorf("expected '{'")
	}

	obj := map[string]any{}
	order := make([]string, 0, 8)
	for dec.More() {
		keyTok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		key, ok := keyTok.(string)
		if !ok {
			return nil, nil, fmt.Errorf("expected string key")
		}
		order = append(order, key)

		val, err := decodeAny(dec)
		if err != nil {
			return nil, nil, err
		}
		obj[key] = val
	}
	if _, err := dec.Token(); err != nil { // '}'
		return nil, nil, err
	}
	return obj, order, nil
}

func decodeAny(dec *json.Decoder) (any, error) {
	tok, err := dec.Token()
	if err != nil {
		return nil, err
	}
	switch v := tok.(type) {
	case json.Delim:
		switch v {
		case '{':
			// Push back by constructing a new decoder isn't possible; instead, parse object manually.
			m := map[string]any{}
			for dec.More() {
				kTok, err := dec.Token()
				if err != nil {
					return nil, err
				}
				k := kTok.(string)
				vv, err := decodeAny(dec)
				if err != nil {
					return nil, err
				}
				m[k] = vv
			}
			if _, err := dec.Token(); err != nil {
				return nil, err
			}
			return m, nil
		case '[':
			var arr []any
			for dec.More() {
				vv, err := decodeAny(dec)
				if err != nil {
					return nil, err
				}
				arr = append(arr, vv)
			}
			if _, err := dec.Token(); err != nil {
				return nil, err
			}
			return arr, nil
		default:
			return nil, fmt.Errorf("unexpected delimiter")
		}
	case json.Number:
		// Prefer int when possible.
		if i, err := v.Int64(); err == nil {
			return i, nil
		}
		if f, err := v.Float64(); err == nil {
			return f, nil
		}
		return v.String(), nil
	default:
		return v, nil
	}
}

func (d *DynamicData) UnmarshalXML(dec *xml.Decoder, start xml.StartElement) error {
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
				for _, k := range order {
					appendKey(k)
				}
				d.Items = append(d.Items, item)
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return nil
			}
		}
	}
	return nil
}

func readXMLItem(dec *xml.Decoder, start xml.StartElement) (map[string]any, []string, error) {
	item := map[string]any{}
	order := make([]string, 0, 8)
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			key := t.Name.Local
			var content string
			if err := dec.DecodeElement(&content, &t); err != nil {
				return nil, nil, err
			}
			order = append(order, key)
			item[key] = parseScalar(strings.TrimSpace(content))
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return item, order, nil
			}
		}
	}
}

func parseScalar(s string) any {
	if s == "" {
		return ""
	}
	if b, err := strconv.ParseBool(s); err == nil {
		return b
	}
	if i, err := strconv.ParseInt(s, 10, 64); err == nil {
		return i
	}
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		return f
	}
	return s
}
