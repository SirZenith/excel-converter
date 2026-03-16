package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/charmbracelet/log"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"
)

const keyEnvLogLevel = "EXCEL_CONVERTER_LOG_LEVEL"

const excelExportConfigSheet = "export"

const excelDataTypeNameListPrefix = "list"

const (
	excelDataTypeNameInt    = "int"
	excelDataTypeNameFloat  = "float"
	excelDataTypeNameString = "string"
	excelDataTypeNameObject = "object"
	excelDataTypeNameUID    = "uid"
)

const (
	excelDataTypeUnknown = iota
	excelDataTypeInt
	excelDataTypeFloat
	excelDataTypeString
	excelDataTypeList
	excelDataTypeObject
	excelDataTypeUID
)

func main() {
	logger := log.NewWithOptions(os.Stderr, log.Options{
		ReportTimestamp: true,
		TimeFormat:      time.TimeOnly,
	})
	log.SetDefault(logger)
	log.SetLevel(getLogLevel())

	var (
		inputName  string
		outputName string
	)

	cmd := &cli.Command{
		Name:  "excel-converter",
		Usage: "a tool that converts Excel file to JSON",
		Flags: []cli.Flag{},
		Arguments: []cli.Argument{
			&cli.StringArg{
				Name:        "inputName",
				UsageText:   "<input>",
				Destination: &inputName,
			},
			&cli.StringArg{
				Name:        "outputName",
				UsageText:   " <output>",
				Destination: &outputName,
			},
		},
		Action: func(_ context.Context, cmd *cli.Command) error {
			if err := convertFile(inputName, outputName); err == nil {
				log.Infof("exported: %s -> %s", inputName, outputName)
			} else {
				log.Warnf("failed to export %s: %s\n", cmd.String("inputName"), err)
			}
			return nil
		},
	}

	err := cmd.Run(context.Background(), os.Args)
	if err != nil {
		log.Errorf("Error occured while converting: %s\n", err)
		os.Exit(1)
	}
}

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

// convertFile converts given input file to output json.
func convertFile(inputName, outputName string) error {
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

	outputDir := filepath.Dir(outputName)
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %s", err)
	}

	f, err := excelize.OpenFile(inputName)
	if err != nil {
		return fmt.Errorf("failed to open input file: %s", err)
	}
	defer f.Close()

	exportRows, err := f.GetRows("export")
	if err != nil {
		return fmt.Errorf("failed to read export config sheet: %s", err)
	}

	if len(exportRows) <= 0 {
		return fmt.Errorf("no target found in export sheet")
	}

	result := map[string]any{}
	for _, row := range exportRows {
		if len(row) <= 0 {
			continue
		}

		sheet := row[0]

		if sheetData, err := convertSheet(f, sheet); err == nil {
			result[sheet] = sheetData
		} else {
			log.Warnf("failed to convert sheet %s: %s\n", sheet, err)
		}
	}

	return writeJSON(outputName, result)
}

// isExcelFile checks file type by extension in filename
func isExcelFile(filename string) bool {
	ext := strings.ToLower(filepath.Ext(filename))
	return ext == ".xlsx" || ext == ".xls"
}

// convertSheet converts a single sheet in Excel file to map object.
// This function assume the first three line of a sheet is meta data.
// - 1st: human readable column name comment
// - 2nd: column name used in exported JSON
// - 3rd: data type of this column
func convertSheet(f *excelize.File, sheet string) (map[string]any, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read sheet '%s': %s", sheet, err)
	}

	commentNames, err := readMetaRow(rows, "comment names")
	if err != nil {
		return nil, err
	}

	exportNames, err := readMetaRow(rows, "export names")
	if err != nil {
		return nil, err
	}

	dataTypes, err := readMetaRow(rows, "data types")
	if err != nil {
		return nil, err
	}

	fields, uidIndex, err := composeFieldList(commentNames, exportNames, dataTypes)
	if err != nil {
		return nil, fmt.Errorf("meta error: %s", err)
	}

	fieldCnt := len(fields)
	if fieldCnt <= 0 {
		return nil, fmt.Errorf("no field is defined in sheet %s", sheet)
	}

	result := map[string]any{}
	var row []string
	for rows.Next() {
		row, err = rows.Columns()
		if err != nil {
			err = fmt.Errorf("failed to read row from sheet '%s': %s", sheet, err)
			break
		}

		if len(row) == 0 {
			// skip empty row
			continue
		}

		uid := ""
		rowData := map[string]any{}
		for index, rawValue := range row {
			if index >= fieldCnt {
				break
			}

			if rawValue == "" {
				continue
			}

			field := fields[index]
			if field.name == "" {
				continue
			}

			if index == uidIndex {
				uid = rawValue
			}

			storeValue(rowData, field, rawValue)
		}

		if uid != "" {
			result[uid] = rowData
		}
	}

	if err = rows.Close(); err != nil {
		fmt.Println(err)
		log.Warnf("failed to close sheet %s: %s\n", sheet, err)
	}

	return result, nil
}

// readMetaRow reads one meta row from sheet rows object.
func readMetaRow(rows *excelize.Rows, metaName string) ([]string, error) {
	if !rows.Next() {
		return nil, fmt.Errorf("unexpected EOF while reading %s meta row", metaName)
	}

	row, err := rows.Columns()
	if err != nil {
		return nil, fmt.Errorf("failed to read %s meta row: ", metaName, err)
	}

	return row, nil
}

type dataField struct {
	index       int
	dataType    int
	elementType string
	name        string
	comment     string
	indexPath   []string
}

// composeFieldList makes a list of fields from meta rows.
func composeFieldList(commentNames, exportNames, dataTypes []string) ([]dataField, int, error) {
	totalCnt := len(exportNames)
	if len(dataTypes) < totalCnt {
		return nil, -1, fmt.Errorf("the number of data types is less then the number of exported columns")
	}

	commentCnt := len(commentNames)
	usedNames := map[string]bool{}
	results := []dataField{}

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

		if commentCnt >= i {
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

		results = append(results, dataField{
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

// typeStringToDataType converts type string to type enum value.
func typeStringToDataType(typeString string) int {
	dataType := excelDataTypeUnknown

	switch typeString {
	case excelDataTypeNameInt:
		dataType = excelDataTypeInt
	case excelDataTypeNameFloat:
		dataType = excelDataTypeFloat
	case excelDataTypeNameString:
		dataType = excelDataTypeString
	case excelDataTypeNameObject:
		dataType = excelDataTypeObject
	case excelDataTypeNameUID:
		dataType = excelDataTypeUID
	default:
		if strings.HasPrefix(typeString, excelDataTypeNameListPrefix) {
			dataType = excelDataTypeList
		}
	}

	return dataType
}

// storeValue writes converted value to result map.
func storeValue(resultMap map[string]any, field dataField, rawValue string) error {
	var (
		value any
		err   error
	)

	indexPath := field.indexPath
	if indexPath == nil {
		value, err = convertRawValue(rawValue, field.dataType, field.elementType)
		if err == nil {
			resultMap[field.name] = value
		}
	} else {
		lastIndex := len(indexPath) - 1
		target := resultMap

		for i := range lastIndex - 1 {
			segment := indexPath[i]
			container, ok := target[segment].(map[string]any)
			if ok {
				target = container
			} else {
				err = fmt.Errorf("indexing into a non-map value", field.name)
			}
		}

		if err == nil {
			lastSegment := indexPath[lastIndex]
			value, err = convertRawValue(rawValue, field.dataType, field.elementType)
			if err == nil {
				target[lastSegment] = value
			}
		}
	}

	return err
}

// convertRawValue converts Excel value to a data.
func convertRawValue(rawValue string, dataType int, elementType string) (any, error) {
	switch dataType {
	case excelDataTypeInt:
		return strconv.Atoi(rawValue)
	case excelDataTypeFloat:
		return strconv.ParseFloat(rawValue, 10)
	case excelDataTypeString:
		return rawValue, nil
	case excelDataTypeList:
		separator := elementType[:1]
		typeName := elementType[1:]

		elements := strings.Split(rawValue, separator)
		innerType := typeStringToDataType(typeName)

		innerElementType := ""
		if innerType == excelDataTypeList {
			innerElementType = strings.TrimPrefix(typeName, excelDataTypeNameListPrefix)
		}

		list := []any{}
		var (
			value any
			err   error
		)
		for i, element := range elements {
			value, err = convertRawValue(element, innerType, innerElementType)
			if err != nil {
				err = fmt.Errorf("failed to parse list element at index %d: %s", i, err)
				break
			}
			list = append(list, value)
		}

		return list, err
	case excelDataTypeObject:
		value := map[string]any{}
		err := json.Unmarshal([]byte(rawValue), &value)
		return value, err
	case excelDataTypeUID:
		return rawValue, nil
	default:
		return nil, fmt.Errorf("unknown data type")
	}
}

// writeJSON generate JSON file from map data.
func writeJSON(outputPath string, data any) error {
	file, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %s", err)
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	encoder.SetEscapeHTML(false) // HTML escape is not needed

	if err := encoder.Encode(data); err != nil {
		return fmt.Errorf("failed to write output JSON: %s", err)
	}

	return nil
}
