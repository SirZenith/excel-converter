package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/charmbracelet/log"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"
)

func cmdToJSON() *cli.Command {
	var (
		inputName string
	)

	return &cli.Command{
		Name:  "json",
		Usage: "converting Excel file to JSON",
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:     "output-dir",
				Aliases:  []string{"o"},
				Required: true,
			},
			&cli.StringFlag{
				Name:     "export-config",
				Aliases:  []string{"c"},
				Required: true,
			},
			&cli.StringFlag{
				Name:    "script-gen-dir",
				Aliases: []string{"s"},
			},
		},
		Arguments: []cli.Argument{
			&cli.StringArg{
				Name:        "inputName",
				UsageText:   "<input>",
				Destination: &inputName,
			},
		},
		Action: func(_ context.Context, cmd *cli.Command) error {
			outputDir := cmd.String("output-dir")
			configPath := cmd.String("export-config")

			if err := os.MkdirAll(outputDir, 0755); err != nil {
				return fmt.Errorf("failed to create JSON output directory: %s", err)
			}

			nameStem, args, err := getOutputArgs(configPath, inputName)
			if err != nil {
				return fmt.Errorf("can't find export name for %s: %s", inputName, err)
			}

			outputName := filepath.Join(outputDir, nameStem+".json")
			if err := convertFile(inputName, outputName, args); err != nil {
				return fmt.Errorf("failed to export %s: %s\n", cmd.String("inputName"), err)
			}

			fmt.Print(outputName)

			return nil
		},
	}
}

// convertFile converts given input file to output json.
func convertFile(inputName, outputName string, _ []string) error {
	err := excelInputFileCheck(inputName)
	if err != nil {
		return err
	}

	f, err := excelize.OpenFile(inputName)
	if err != nil {
		return fmt.Errorf("failed to open input file: %s", err)
	}
	defer f.Close()

	sheetList, err := getExportSheetList(f)
	if err != nil {
		return err
	}

	result := map[string]any{}
	for _, sheet := range sheetList {
		if sheetData, err := convertSheet(f, sheet); err == nil {
			result[sheet] = sheetData
		} else {
			log.Warnf("failed to convert sheet %s: %s\n", sheet, err)
		}
	}

	return writeJSON(outputName, result)
}

// convertSheet converts a single sheet in Excel file to map object.
func convertSheet(f *excelize.File, sheet string) (map[string]any, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read sheet '%s': %s", sheet, err)
	}

	fields, uidIndex, err := readExcelFieldList(rows)
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
		log.Warnf("failed to close sheet %s: %s\n", sheet, err)
	}

	return result, nil
}

// storeValue writes converted value to result map.
func storeValue(resultMap map[string]any, field excelField, rawValue string) error {
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
	/* case excelDataTypeObject:
	value := map[string]any{}
	err := json.Unmarshal([]byte(rawValue), &value)
	return value, err */
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
