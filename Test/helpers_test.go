package data2doc_test

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"compress/zlib"
	"encoding/hex"
	"io"
	"strconv"
	"strings"
	"testing"

	"golang.org/x/text/encoding/charmap"
)

func readZipEntry(t *testing.T, zipBytes []byte, name string) []byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(zipBytes), int64(len(zipBytes)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	for _, f := range r.File {
		if f.Name != name {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", name, err)
		}
		defer rc.Close()
		b, err := io.ReadAll(rc)
		if err != nil {
			t.Fatalf("read %s: %v", name, err)
		}
		return b
	}
	t.Fatalf("missing entry %s", name)
	return nil
}

func readZipEntryString(t *testing.T, zipBytes []byte, name string) string {
	return string(readZipEntry(t, zipBytes, name))
}

func extractFlateStreams(b []byte) [][]byte {
	out := make([][]byte, 0)
	needle := []byte("stream")
	endNeedle := []byte("endstream")
	i := 0
	for {
		idx := bytes.Index(b[i:], needle)
		if idx < 0 {
			break
		}
		idx += i + len(needle)
		if idx < len(b) && b[idx] == '\r' {
			idx++
		}
		if idx < len(b) && b[idx] == '\n' {
			idx++
		}
		// Prefer endstream markers that start on a new line to avoid false matches
		// within binary/compressed content.
		end := -1
		if e := bytes.Index(b[idx:], []byte("\nendstream")); e >= 0 {
			end = idx + e
		} else if e := bytes.Index(b[idx:], []byte("\rendstream")); e >= 0 {
			end = idx + e
		} else if e := bytes.Index(b[idx:], endNeedle); e >= 0 {
			end = idx + e
		}
		if end < 0 {
			break
		}

		raw := b[idx:end]
		// Strip a single trailing line break if present.
		raw = bytes.TrimSuffix(raw, []byte("\r\n"))
		raw = bytes.TrimSuffix(raw, []byte("\n"))
		raw = bytes.TrimSuffix(raw, []byte("\r"))
		if zr, err := zlib.NewReader(bytes.NewReader(raw)); err == nil {
			inflated, _ := io.ReadAll(zr)
			_ = zr.Close()
			if len(inflated) > 0 {
				out = append(out, inflated)
			} else {
				out = append(out, raw)
			}
		} else if fr := flate.NewReader(bytes.NewReader(raw)); fr != nil {
			inflated, _ := io.ReadAll(fr)
			_ = fr.Close()
			if len(inflated) > 0 {
				out = append(out, inflated)
			} else {
				out = append(out, raw)
			}
		} else {
			out = append(out, raw)
		}
		i = end + len(endNeedle)
	}
	return out
}

func pdfLiteralEscapedISO88591(s string) string {
	enc := charmap.ISO8859_1.NewEncoder()
	iso, _ := enc.Bytes([]byte(s))

	buf := make([]byte, 0, len(iso)*2)
	for _, c := range iso {
		switch c {
		case '\\', '(', ')':
			buf = append(buf, '\\', c)
		default:
			if c < 32 || c > 126 {
				buf = append(buf, '\\')
				buf = append(buf, byte('0'+((c>>6)&7)))
				buf = append(buf, byte('0'+((c>>3)&7)))
				buf = append(buf, byte('0'+(c&7)))
			} else {
				buf = append(buf, c)
			}
		}
	}
	return string(buf)
}

func pdfHexISO88591(s string) (upper, lower string) {
	enc := charmap.ISO8859_1.NewEncoder()
	iso, _ := enc.Bytes([]byte(s))
	h := hex.EncodeToString(iso)
	return "<" + strings.ToUpper(h) + ">", "<" + strings.ToLower(h) + ">"
}

func countPDFPages(b []byte) int {
	// Count occurrences of '/Type' + whitespace + '/Page' while excluding '/Pages'.
	count := 0
	for i := 0; i+5 < len(b); i++ {
		if b[i] != '/' || !bytes.HasPrefix(b[i:], []byte("/Type")) {
			continue
		}
		j := i + len("/Type")
		for j < len(b) {
			c := b[j]
			if c == ' ' || c == '\t' || c == '\r' || c == '\n' {
				j++
				continue
			}
			break
		}
		if j >= len(b) || b[j] != '/' {
			continue
		}
		if bytes.HasPrefix(b[j:], []byte("/Pages")) {
			continue
		}
		if bytes.HasPrefix(b[j:], []byte("/Page")) {
			count++
		}
	}
	return count
}
func xmlAttrInt(s, attr string) (int, bool) {
	// Fast-path for w:w="123" etc.
	idx := strings.Index(s, attr+"=\"")
	if idx < 0 {
		return 0, false
	}
	idx += len(attr) + len("=\"")
	end := strings.IndexByte(s[idx:], '"')
	if end < 0 {
		return 0, false
	}
	v, err := strconv.Atoi(s[idx : idx+end])
	if err != nil {
		return 0, false
	}
	return v, true
}

func ptrBool(v bool) *bool { return &v }
