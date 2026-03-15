package service

import (
	"fmt"
	"regexp"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

type excelSheetMeta struct {
	Name          string
	FieldToCol    map[string]string // lower(field) -> column letter (A, B, ...)
	DataStartRow  int
	DataEndRow    int // last row that contains data/subtotals (excludes total row)
	TotalRow      int // 0 when absent
	SubtotalRows  []int
	HasSubtotals  bool
	GeneratedRows int // last generated row index (>=1)
}

type excelSheetRegistry struct {
	byName map[string]excelSheetMeta // lower(sheetName) -> meta
}

func newExcelSheetRegistry() *excelSheetRegistry {
	return &excelSheetRegistry{byName: map[string]excelSheetMeta{}}
}

func (r *excelSheetRegistry) Register(meta excelSheetMeta) {
	if r == nil {
		return
	}
	if r.byName == nil {
		r.byName = map[string]excelSheetMeta{}
	}
	r.byName[strings.ToLower(strings.TrimSpace(meta.Name))] = meta
}

func (r *excelSheetRegistry) Get(sheet string) (excelSheetMeta, bool) {
	if r == nil || r.byName == nil {
		return excelSheetMeta{}, false
	}
	m, ok := r.byName[strings.ToLower(strings.TrimSpace(sheet))]
	return m, ok
}

type pendingExcelFormula struct {
	Sheet   string
	Cell    string
	Formula string
}

// v2ResolveRowFormula converts a field-based expression into a per-row Excel formula.
// Example: "price * qty" -> "A2*B2".
func v2ResolveRowFormula(expr string, fieldToCol map[string]string, row int) (string, error) {
	if strings.TrimSpace(expr) == "" {
		return "", fmt.Errorf("empty formula")
	}
	if row < 1 {
		return "", fmt.Errorf("invalid row")
	}
	out := strings.Builder{}
	in := expr
	for i := 0; i < len(in); {
		r := rune(in[i])
		if isIdentStart(r) {
			j := i + 1
			for j < len(in) {
				r2 := rune(in[j])
				if !isIdentContinue(r2) {
					break
				}
				j++
			}
			tok := in[i:j]
			col, ok := fieldToCol[strings.ToLower(tok)]
			if ok {
				out.WriteString(fmt.Sprintf("%s%d", col, row))
			} else {
				out.WriteString(tok)
			}
			i = j
			continue
		}
		out.WriteByte(in[i])
		i++
	}
	return out.String(), nil
}

func isIdentStart(r rune) bool {
	return r == '_' || unicode.IsLetter(r)
}

func isIdentContinue(r rune) bool {
	return r == '_' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

var sheetTokenRe = regexp.MustCompile(`sheet:([^\.\s]+)\.([A-Za-z_][A-Za-z0-9_]*)`)

// v2ResolveSheetTokens replaces occurrences like "sheet:Employees.salary" with a range like "Employees!B2:B10".
func v2ResolveSheetTokens(formula string, reg *excelSheetRegistry) (string, error) {
	if strings.TrimSpace(formula) == "" {
		return formula, nil
	}
	matches := sheetTokenRe.FindAllStringSubmatchIndex(formula, -1)
	if len(matches) == 0 {
		return formula, nil
	}
	out := strings.Builder{}
	last := 0
	for _, m := range matches {
		fullStart, fullEnd := m[0], m[1]
		sheetStart, sheetEnd := m[2], m[3]
		fieldStart, fieldEnd := m[4], m[5]
		out.WriteString(formula[last:fullStart])
		sheet := formula[sheetStart:sheetEnd]
		field := formula[fieldStart:fieldEnd]
		meta, ok := reg.Get(sheet)
		if !ok {
			return "", fmt.Errorf("unknown sheet '%s' in formula", sheet)
		}
		col, ok := meta.FieldToCol[strings.ToLower(field)]
		if !ok {
			return "", fmt.Errorf("unknown field '%s' in sheet '%s'", field, meta.Name)
		}
		start := maxInt(2, meta.DataStartRow)
		end := maxInt(start, meta.DataEndRow)
		sheetRef := excelQuoteSheetName(meta.Name)
		out.WriteString(fmt.Sprintf("%s!%s%d:%s%d", sheetRef, col, start, col, end))
		last = fullEnd
	}
	out.WriteString(formula[last:])
	return out.String(), nil
}

func excelQuoteSheetName(name string) string {
	v := strings.TrimSpace(name)
	if v == "" {
		return "Sheet1"
	}
	// Always quote for safety; escape single quotes by doubling.
	es := strings.ReplaceAll(v, "'", "''")
	return "'" + es + "'"
}

func excelColLetter(n int) (string, error) {
	return excelize.ColumnNumberToName(n)
}
