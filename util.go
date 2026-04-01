package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/charmbracelet/log"
	"github.com/xuri/excelize/v2"
)

const exportConfigArgsSeparator = "|"
const excelExportConfigSheet = "export"
const excelSystemConfigSheet = "config"

func getLogLevel() log.Level {
	level := os.Getenv(keyEnvLogLevel)
	switch level {
	case "debug":
		return log.DebugLevel
	case "info":
		return log.InfoLevel
	case "warn":
		return log.WarnLevel
	case "fatal":
		return log.FatalLevel
	default:
		return log.InfoLevel
	}
}

func getOutputArgs(configFile, inputName string) (string, []string, error) {
	var err error

	basename := filepath.Base(inputName)
	ext := filepath.Ext(basename)
	stem := strings.TrimSuffix(basename, ext)

	file, err := os.Open(configFile)
	if err != nil {
		return "", nil, fmt.Errorf("failed to open config file: %s", err)
	}
	defer file.Close()

	outputName := ""
	var args []string

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		if strings.HasPrefix(line, stem) {
			parts := strings.Split(line, exportConfigArgsSeparator)
			if len(parts) < 2 {
				err = fmt.Errorf("export config line lacks necessary argument")
				break
			}

			outputName = parts[1]
			args = parts[2:]
		}
	}

	if err = scanner.Err(); err != nil {
		return "", nil, fmt.Errorf("error occured while reading config file: %s", err)
	}

	if outputName == "" {
		return "", nil, fmt.Errorf("no export name found for %s", inputName)
	}

	return outputName, args, nil
}

// excelInputFileCheck checks if given file path is OK for further processing.
func excelInputFileCheck(inputName string) error {
	if !isExcelFile(inputName) {
		return fmt.Errorf("input is not a Excel file: %s", inputName)
	}

	info, err := os.Stat(inputName)
	if err != nil {
		return fmt.Errorf("can't not access input file: %s", err)
	}

	if info.IsDir() {
		return fmt.Errorf("target %s is a directory", inputName)
	}

	return nil
}

// isExcelFile checks file type by extension in filename
func isExcelFile(filename string) bool {
	ext := strings.ToLower(filepath.Ext(filename))
	return ext == ".xlsx" || ext == ".xls"
}

// getExportSheetList reads out a list of export target name list.
func getExportSheetList(f *excelize.File) ([]string, error) {
	exportRows, err := f.GetRows(excelExportConfigSheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read export config sheet: %s", err)
	}

	if len(exportRows) <= 0 {
		return nil, fmt.Errorf("no target found in export sheet")
	}

	sheetList := []string{}

	for _, row := range exportRows {
		if len(row) > 0 {
			sheetList = append(sheetList, row[0])
		}

	}

	return sheetList, nil
}

// readMetaRow reads one meta row from sheet rows object.
func readMetaRow(rows *excelize.Rows, metaName string) ([]string, error) {
	if !rows.Next() {
		return nil, fmt.Errorf("unexpected EOF while reading %s meta row", metaName)
	}

	row, err := rows.Columns()
	if err != nil {
		return nil, fmt.Errorf("failed to read %s meta row: %s", metaName, err)
	}

	return row, nil
}

type excelField struct {
	index       int
	dataType    int
	elementType string
	name        string
	comment     string
	indexPath   []string
}

// composeExcelFieldList makes a list of fields from meta rows.
func composeExcelFieldList(commentNames, exportNames, dataTypes []string) ([]excelField, int, error) {
	totalCnt := len(exportNames)
	if len(dataTypes) < totalCnt {
		return nil, -1, fmt.Errorf("the number of data types is less then the number of exported columns")
	}

	commentCnt := len(commentNames)
	usedNames := map[string]bool{}
	results := []excelField{}

	uidColIndex := -1
	var err error

	for i := range totalCnt {
		name := exportNames[i]
		comment := ""
		var indexPath []string

		if _, repeated := usedNames[name]; repeated {
			err = fmt.Errorf("column name '%s' is repeated", name)
			break
		}

		if commentCnt > i {
			comment = commentNames[i]
		}

		typeString := dataTypes[i]
		elementType := ""
		dataType := typeStringToDataType(typeString)

		if name != "" {
			if dataType == excelDataTypeUnknown {
				err = fmt.Errorf("no proper data type is assigned to column %d", i+1)
				break
			}

			if dataType == excelDataTypeUID {
				if uidColIndex >= 0 {
					err = fmt.Errorf("multiple UID column is defined at column %d and %d", uidColIndex+1, i+1)
					break
				}

				uidColIndex = i
			}

			usedNames[name] = true

			if dataType == excelDataTypeList {
				elementType = strings.TrimPrefix(typeString, excelDataTypeNameListPrefix)
			}

			if strings.Contains(name, ".") {
				indexPath = strings.Split(name, ".")
			}
		}

		results = append(results, excelField{
			index:       i,
			dataType:    dataType,
			elementType: elementType,
			name:        name,
			comment:     comment,
			indexPath:   indexPath,
		})
	}

	if uidColIndex < 0 {
		err = fmt.Errorf("no UID column is defined")
	}

	return results, uidColIndex, err
}

// readExcelFieldList assumes rows is now at its begining, and has following meta
// rows:
// - 1st: human readable column name comment
// - 2nd: column name used in exported JSON
// - 3rd: data type of this column
func readExcelFieldList(rows *excelize.Rows) ([]excelField, int, error) {
	commentNames, err := readMetaRow(rows, "comment names")
	if err != nil {
		return nil, -1, err
	}

	exportNames, err := readMetaRow(rows, "export names")
	if err != nil {
		return nil, -1, err
	}

	dataTypes, err := readMetaRow(rows, "data types")
	if err != nil {
		return nil, -1, err
	}

	return composeExcelFieldList(commentNames, exportNames, dataTypes)
}

// writeLn writes one line of string to writer.
func writeStringLn(writer *bufio.Writer, content ...string) error {
	for _, s := range content {
		_, err := writer.WriteString(s)
		if err != nil {
			return err
		}
	}

	_, err := writer.WriteString("\n")

	return err
}
