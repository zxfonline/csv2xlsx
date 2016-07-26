package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"path"
	"strings"

	"github.com/tealeg/xlsx"
	"github.com/zxfonline/fileutil"
)

var xlsxPath = flag.String("o", "", "Path to the XLSX output file")
var csvPath = flag.String("f", "", "Path to the CSV input file")
var delimiter = flag.String("d", ",", "Delimiter for felds in the CSV input.")

func usage() {
	fmt.Printf(`%s: -f=<CSV Input File> -o=<XLSX Output File> -d=<Delimiter>

`,
		os.Args[0])
}

func generateCSVFromXLSX(csvPath string, XLSXPath string, delimiter string) error {
	csvPath = strings.Replace(csvPath, "\\", "/", -1)
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return err
	}
	defer csvFile.Close()
	reader := csv.NewReader(csvFile)
	if len(delimiter) > 0 {
		reader.Comma = rune(delimiter[0])
	} else {
		reader.Comma = rune(',')
	}
	xlsxFile := xlsx.NewFile()

	_, fn := path.Split(csvPath)
	ext := path.Ext(fn)
	if ext != "" {
		fn = strings.Split(fn, ext)[0]
	}

	sheet, err := xlsxFile.AddSheet(fn)
	if err != nil {
		return err
	}
	fields, err := reader.Read()
	for err == nil {
		row := sheet.AddRow()
		for _, field := range fields {
			cell := row.AddCell()
			cell.Value = field
		}
		fields, err = reader.Read()
	}
	if err != nil && err != io.EOF {
		fmt.Printf(err.Error())
	}
	return xlsxFile.Save(XLSXPath)
}

func main() {
	flag.Parse()
	if *csvPath == "" {
		if len(os.Args) < 3 {
			usage()
			return
		}
		flag.Parse()
	}
	if *xlsxPath == "" {
		*xlsxPath = fileutil.ChangeFileExt(*csvPath, ".xlsx")

	}
	err := generateCSVFromXLSX(*csvPath, *xlsxPath, *delimiter)
	if err != nil {
		panic(err)
	}
}
