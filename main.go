package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var sourceXLSXPath = flag.String("i", "", "Path to the XLSX input file")

var destXSLXPath = flag.String("o", "", "Path to the XLSX output file")

var sheetIndex = flag.Int("idx", 1, "Sheet Index")

var rowsToDelete = flag.String("r", "2", "Rows to remove separated by comma or a range")

func main() {
	flag.Parse()
	if *sourceXLSXPath == "" || *destXSLXPath == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	processFile(*sourceXLSXPath, *destXSLXPath, *rowsToDelete, *sheetIndex)

}

func processFile(sourceXLSXPath string, destXSLXPath string, rowsToDelete string, sheetIndex int) {

	xlsx, err := excelize.OpenFile(sourceXLSXPath)
	if err != nil {
		fmt.Println(err)
		return
	}

	rowRangeStrings := strings.Split(rowsToDelete, ":")

	sheetName := xlsx.GetSheetName(sheetIndex)

	if len(rowRangeStrings) == 1 {

		rowNumT, err := strconv.ParseInt(rowRangeStrings[0], 10, 64)
		if err != nil {
			fmt.Println(err)
			return
		}
		rowNum := int(rowNumT)
		xlsx.RemoveRow(sheetName, rowNum-1)

	} else if len(rowRangeStrings) == 2 {

		rowNumRangeBegin, err := strconv.ParseInt(rowRangeStrings[0], 10, 64)
		if err != nil {
			fmt.Println(err)
			return
		}
		rowNumRangeEnd, err := strconv.ParseInt(rowRangeStrings[1], 10, 64)
		nOfRowsToDelete := int(rowNumRangeEnd) - int(rowNumRangeBegin) + 1
		for nOfRowsToDelete > 0 {
			xlsx.RemoveRow(sheetName, int(rowNumRangeBegin)-1)
			nOfRowsToDelete--
		}
	}

	xlsx.SaveAs(destXSLXPath)
	return
}
