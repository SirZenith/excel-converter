package main

import (
	"bufio"
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strconv"

	"github.com/charmbracelet/log"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"

	"github.com/iancoleman/strcase"
)

func cmdToScript() *cli.Command {
	var (
		inputName string
	)

	return &cli.Command{
		Name:  "csharp",
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
				Name:     "json-dir",
				Aliases:  []string{"o"},
				Required: true,
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
			jsonDir := cmd.String("json-dir")

			if !isExcelFile(inputName) {
				return fmt.Errorf("input is not a Excel file: %s", inputName)
			}

			if err := os.MkdirAll(outputDir, 0755); err != nil {
				return fmt.Errorf("failed to create script generation directory: %s", err)
			}

			nameStem, _, err := getOutputArgs(configPath, inputName)
			if err != nil {
				return fmt.Errorf("can't find export name for %s: %s", inputName, err)
			}

			jsonName := filepath.Join(jsonDir, nameStem+".json")
			outputName := filepath.Join(outputDir, strcase.ToCamel(nameStem)+".cs")
			if err := genDataClassDefinition(inputName, outputName, jsonName, nameStem); err != nil {
				return fmt.Errorf("failed to export %s: %s\n", cmd.String("inputName"), err)
			}

			fmt.Print(outputName)

			return nil
		},
	}
}

// genDataClassDefinition generates C# data class definition for given Excel file.
func genDataClassDefinition(inputName, outputName, jsonName, nameStem string) error {
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

	rootClass := dataClass{
		name:         strcase.ToCamel(nameStem) + "Cfg",
		fields:       map[string]dataField{},
		childClasses: map[string]dataClass{},
		jsonFile:     jsonName,
	}
	for _, sheet := range sheetList {
		if child, err := genSheetDataClass(f, sheet); err == nil {
			rootClass.fields[sheet] = dataField{
				t: dataFieldType{
					t: excelDataTypeDictionary,
					keyType: &dataFieldType{
						t: excelDataTypeString,
					},
					elementType: &dataFieldType{
						t:          excelDataTypeObject,
						customName: child.name,
					},
				},
				name: sheet,
			}
			rootClass.childClasses[child.name] = *child
		} else {
			log.Warnf("failed to generate data class for sheet %s: %s\n", sheet, err)
		}
	}

	return writeCSScript(outputName, rootClass)
}

// genSheetDataClass generate data class info for given sheet.
func genSheetDataClass(f *excelize.File, sheet string) (*dataClass, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read sheet '%s': %s", sheet, err)
	}

	fields, _, err := readExcelFieldList(rows)
	if err != nil {
		return nil, fmt.Errorf("meta error: %s", err)
	}

	fieldCnt := len(fields)
	if fieldCnt <= 0 {
		return nil, fmt.Errorf("no field is defined in sheet %s", sheet)
	}

	if err = rows.Close(); err != nil {
		log.Warnf("failed to close sheet %s: %s\n", sheet, err)
	}

	result := genDataClassFromFields(sheet, fields)

	return &result, nil
}

// genDataClassFromFields takes fields list as input, and return a list of data
// class definitions.
func genDataClassFromFields(sheetName string, fields []excelField) dataClass {
	mainClass := dataClass{
		name:   strcase.ToCamel(sheetName) + "Sheet",
		fields: map[string]dataField{},
	}

	for _, field := range fields {
		name := field.name

		if field.indexPath == nil {
			switch field.dataType {
			case excelDataTypeUnknown:
				// no opt
			case excelDataTypeInt, excelDataTypeFloat, excelDataTypeString:
				mainClass.fields[name] = dataField{
					t: dataFieldType{
						t: field.dataType,
					},
					name:    name,
					comment: field.comment,
				}
			case excelDataTypeUID:
				mainClass.fields[name] = dataField{
					t: dataFieldType{
						t: excelDataTypeString,
					},
					name:    name,
					comment: field.comment,
				}
			case excelDataTypeList:
				mainClass.fields[name] = dataField{
					t:       parseListElementType(field.elementType),
					name:    name,
					comment: field.comment,
				}
			}
		} else {
			addChildClassByIndexPath(&mainClass, field.indexPath, field.dataType, field.comment)
		}
	}

	return mainClass
}

func addChildClassByIndexPath(parentClass *dataClass, segments []string, dataType int, comment string) error {
	var err error
	last_index := len(segments) - 1
	classWalker := parentClass

	for index := range last_index - 1 {
		seg := segments[index]
		next_seg := segments[index+1]

		fieldName := seg

		_, err = strconv.Atoi(next_seg)
		if err != nil {
			// current segment is not a list
			childMap := classWalker.childClasses
			if childMap == nil {
				childMap = map[string]dataClass{}
				classWalker.childClasses = childMap
			}

			className := strcase.ToCamel(seg) + "DataClass"

			child, ok := childMap[className]
			if !ok {
				child := dataClass{
					name:   className,
					fields: map[string]dataField{},
				}
				childMap[className] = child
			}

			classWalker.fields[fieldName] = dataField{
				t: dataFieldType{
					t:          excelDataTypeObject,
					customName: className,
				},
				name: fieldName,
			}

			classWalker = &child
		} else if index+1 == last_index {
			// only the second to last segment is allowed to list
			classWalker.fields[fieldName] = dataField{
				t: dataFieldType{
					t: excelDataTypeList,
					elementType: &dataFieldType{
						t: dataType,
					},
				},
				name:    fieldName,
				comment: comment,
			}
			classWalker = nil
		} else {
			err = fmt.Errorf("only the last segment of index path can be number")
			break
		}
	}

	if err == nil && classWalker != nil {
		lastField := segments[last_index]
		classWalker.fields[lastField] = dataField{
			t: dataFieldType{
				t: dataType,
			},
			name:    lastField,
			comment: comment,
		}
	}

	return err
}

// writeCSScript generate script file for data class.
func writeCSScript(outputPath string, rootClass dataClass) error {
	file, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create script output file: %s", err)
	}
	defer file.Close()

	writer := bufio.NewWriter(file)

	err = rootClass.genCSScript(writer, "")
	if err != nil {
		return fmt.Errorf("failed to write class definition string: %s", err)
	}

	err = writer.Flush()
	if err != nil {
		return fmt.Errorf("failed to flush class definition: %s", err)
	}

	return nil
}
