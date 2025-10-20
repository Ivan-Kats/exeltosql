package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	var inFile string
	var outFile string
	var sheet string
	flag.StringVar(&inFile, "in", "", "input xlsx file")
	flag.StringVar(&outFile, "out", "updates.sql", "output sql file")
	flag.StringVar(&sheet, "sheet", "", "sheet name (default first sheet)")
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

	updates, err := generateUpdates(f, sheet)
	if err != nil {
		log.Fatalf("generate updates: %v", err)
	}

	if err := writeSQL(outFile, updates); err != nil {
		log.Fatalf("write sql: %v", err)
	}

	log.Printf("Done: generated %s (%d updates)", outFile, len(updates))
}

func esc(s string) string {
	// эскейп для SQL строк: одинарная кавычка -> два одинарных
	return strings.ReplaceAll(s, "'", "''")
}

// generateUpdates читает переданный excelize.File и sheet и возвращает slice SQL UPDATE строк.
// Поведение соответствует предыдущей реализации: пропускает заголовок (первая строка),
// пропускает строки с < 3 колонками или пустыми code/shortName.
func generateUpdates(f *excelize.File, sheet string) ([]string, error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("get rows: %w", err)
	}

	var updates []string
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
		shortName := strings.TrimSpace(r[2]) // краткое наименование

		if code == "" || shortName == "" {
			continue
		}

		escShort := esc(shortName)
		escCode := esc(code)

		update := fmt.Sprintf(
			"UPDATE documents\nSET data = jsonb_set(data, '{name}', to_jsonb('%s'::text), false)\nWHERE data->>'code' = '%s';",
			escShort, escCode,
		)
		updates = append(updates, update)
	}

	return updates, nil
}

// writeSQL создает выходной файл и записывает BEGIN; updates... COMMIT;
func writeSQL(outFile string, updates []string) error {
	dir := filepath.Dir(outFile)
	if dir != "." && dir != "" {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return fmt.Errorf("mkdir out dir: %w", err)
		}
	}
	of, err := os.Create(outFile)
	if err != nil {
		return fmt.Errorf("create out file: %w", err)
	}
	defer of.Close()

	fmt.Fprintln(of, "BEGIN;")
	for _, u := range updates {
		fmt.Fprintln(of, u)
	}
	fmt.Fprintln(of, "COMMIT;")
	return nil
}
