package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	var inFile string
	var outFile string
	var sheet string
	var onlyDiff bool

	flag.StringVar(&inFile, "in", "", "input xlsx file")
	flag.StringVar(&outFile, "out", "updates.sql", "output sql file")
	flag.StringVar(&sheet, "sheet", "", "sheet name (default first sheet)")
	flag.BoolVar(&onlyDiff, "only-diff", true,
		"generate UPDATE with condition to update only when name differs")
	flag.Parse()

	if inFile == "" {
		log.Fatalf("please provide input file: -in file.xlsx")
	}

	f, err := excelize.OpenFile(inFile)
	if err != nil {
		log.Fatalf("open excel: %v", err)
	}

	if sheet == "" {
		sheets := f.GetSheetList()
		if len(sheets) == 0 {
			log.Fatalf("no sheets in workbook")
		}
		sheet = sheets[0]
	}

	updates, err := loadRows(f, sheet)
	if err != nil {
		log.Fatalf("load rows: %v", err)
	}

	vals, count := buildValuesClause(updates)
	log.Printf("Collected %d value(s) from sheet %q", count, sheet)

	if err := writeValuesSQL(outFile, vals, onlyDiff); err != nil {
		log.Fatalf("write sql: %v", err)
	}

	log.Printf("Done: generated %s with %d values", outFile, count)
}

// esc экранирует одинарные кавычки для вставки в SQL литерал.
func esc(s string) string {
	return strings.ReplaceAll(s, "'", "''")
}

func loadRows(f *excelize.File, sheet string) ([][]string, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("get rows: %w", err)
	}
	return rows, nil
}

// buildValuesClause принимает строки excel (в формате GetRows) и возвращает
// массив value-строк для VALUES ('code','name').
// Ожидает: первая строка - заголовок; колонки: A=code, B=nameFull(игнор), C=shortName
func buildValuesClause(rows [][]string) ([]string, int) {
	var vals []string
	count := 0
	for i, r := range rows {
		if i == 0 {
			// предполагаем, что первая строка — заголовок
			continue
		}
		// Колонки: A=0 B=1 C=2
		if len(r) < 3 {
			continue
		}
		code := strings.TrimSpace(r[0])
		shortName := strings.TrimSpace(r[2])
		if code == "" || shortName == "" {
			continue
		}

		escShort := esc(shortName)
		escCode := esc(code)

		val := fmt.Sprintf("('%s','%s')", escCode, escShort)
		vals = append(vals, val)
		count++
	}
	return vals, count
}

func writeValuesSQL(outFile string, vals []string, onlyDiff bool) error {
	of, err := os.Create(outFile)
	if err != nil {
		return fmt.Errorf("create out file: %w", err)
	}
	defer of.Close()

	fmt.Fprintln(of, "BEGIN;")
	fmt.Fprintln(of)

	// Если нет значений — ничего не делаем
	if len(vals) == 0 {
		fmt.Fprintln(of, "-- no values generated from Excel")
		fmt.Fprintln(of)
		fmt.Fprintln(of, "COMMIT;")
		return nil
	}

	// Печатаем WITH vals(code, name) AS ( VALUES
	//	VALUES... )
	fmt.Fprintln(of, "WITH vals(code, name) AS (")
	fmt.Fprintln(of, "VALUES")
	// join с запятой и переносом строк
	for i, v := range vals {
		sep := ","
		if i == len(vals)-1 {
			sep = ""
		}
		fmt.Fprintf(of, "    %s%s\n", v, sep)
	}
	fmt.Fprintln(of, ")")
	fmt.Fprintln(of)

	// Основной UPDATE: один проход
	// Добавляем условие onlyDiff если нужно
	fmt.Fprintln(of, "UPDATE documents d")
	fmt.Fprintln(of, "SET data = jsonb_set(d.data, '{name}', to_jsonb(v.name::text), false)")
	fmt.Fprintln(of, "FROM vals v")
	fmt.Fprint(of, "WHERE d.data->>'code' = v.code")
	if onlyDiff {
		fmt.Fprintln(of, " AND d.data->>'name' IS DISTINCT FROM v.name;")
	} else {
		fmt.Fprintln(of, ";")
	}

	fmt.Fprintln(of)
	fmt.Fprintln(of, "COMMIT;")
	return nil
}
