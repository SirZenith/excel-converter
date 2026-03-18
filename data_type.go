package main

import (
	"bufio"
	"fmt"
	"sort"
	"strings"
)

const excelDataTypeNameListPrefix = "list"

const (
	excelDataTypeNameInt    = "int"
	excelDataTypeNameFloat  = "float"
	excelDataTypeNameString = "string"
	// excelDataTypeNameObject = "object"
	excelDataTypeNameUID = "uid"
)

const (
	excelDataTypeUnknown = iota
	excelDataTypeInt
	excelDataTypeFloat
	excelDataTypeString
	excelDataTypeList
	excelDataTypeDictionary
	excelDataTypeObject
	excelDataTypeUID
)

const dataClassGeneratedWarning = "// This code is generated via external tool, don't modify it."

type dataFieldType struct {
	t           int
	keyType     *dataFieldType
	elementType *dataFieldType
	customName  string // class name for object type
}

func (t *dataFieldType) getCSTypeName() string {
	switch t.t {
	case excelDataTypeInt:
		return "long"
	case excelDataTypeFloat:
		return "double"
	case excelDataTypeString:
		return "string"
	case excelDataTypeList:
		if t.elementType == nil {
			return ""
		} else {
			return fmt.Sprintf("List<%s>", t.elementType.getCSTypeName())
		}
	case excelDataTypeDictionary:
		return fmt.Sprintf("Dictionary<%s, %s>", t.keyType.getCSTypeName(), t.elementType.getCSTypeName())
	case excelDataTypeObject:
		return t.customName
	case excelDataTypeUID:
		return "string"
	default:
		return ""
	}
}

type dataField struct {
	t       dataFieldType
	name    string
	comment string
}

// genCSScript writes C# definition of current field to writer.

func (f *dataField) genCSScript(writer *bufio.Writer, indent string) error {
	typeName := f.t.getCSTypeName()
	if typeName == "" {
		return fmt.Errorf("failed to generate proper type name")
	}

	if f.name == "" {
		return fmt.Errorf("field name is empty")
	}

	if f.comment != "" {
		writeStringLn(writer, indent, "/// <summary>")
		writeStringLn(writer, indent, "/// ", f.comment)
		writeStringLn(writer, indent, "/// </summary>")
	}

	writer.WriteString(indent)
	writer.WriteString("public ")
	writer.WriteString(typeName)
	writer.WriteString(" ")
	writer.WriteString(f.name)
	writer.WriteString(";")

	return nil
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
	// case excelDataTypeNameObject:
	// dataType = excelDataTypeObject
	case excelDataTypeNameUID:
		dataType = excelDataTypeUID
	default:
		if strings.HasPrefix(typeString, excelDataTypeNameListPrefix) {
			dataType = excelDataTypeList
		}
	}

	return dataType
}

// parseListElementType translate list type string into type data
func parseListElementType(typeStr string) dataFieldType {
	typeStr = typeStr[1:] // remove separator

	result := dataFieldType{}
	result.t = typeStringToDataType(typeStr)
	if result.t != excelDataTypeList {
		return result
	}

	elementTypeStr := strings.TrimPrefix(typeStr, excelDataTypeNameListPrefix)
	elementType := parseListElementType(elementTypeStr)

	result.elementType = &elementType

	return result
}

type dataClass struct {
	name         string
	fields       map[string]dataField
	childClasses map[string]dataClass
	jsonFile     string
}

const dataClassPrelude = `using System;
using System.Collections.Generic;
using UnityEngine;

namespace GameConfig
{
`

// genCSScript writes C# definition of current class to writer.
func (c *dataClass) genCSScript(writer *bufio.Writer, indent string) error {
	fields := []*dataField{}
	for _, f := range c.fields {
		fields = append(fields, &f)
	}
	if len(fields) <= 0 {
		// this class has no field to write
		return nil
	}

	children := []*dataClass{}
	for _, child := range c.childClasses {
		children = append(children, &child)
	}

	sort.Slice(fields, func(i, j int) bool {
		return fields[i].name < fields[j].name
	})
	sort.Slice(children, func(i, j int) bool {
		return children[i].name < children[j].name
	})

	isRoot := indent == ""
	if isRoot {
		_, err := writer.WriteString(dataClassPrelude)
		if err != nil {
			return err
		}

		indent += "    "
	}

	childIndent := indent + "    "

	writeStringLn(writer, indent, dataClassGeneratedWarning)
	if isRoot {
		writeStringLn(writer, indent, "[Serializable]")
	}
	writeStringLn(writer, indent, "public class ", c.name)
	writeStringLn(writer, indent, "{")

	for _, child := range children {
		err := child.genCSScript(writer, childIndent)
		if err != nil {
			return err
		}
		writer.WriteString("\n")
	}

	if c.jsonFile != "" {
		writeStringLn(writer, childIndent, "public readonly static string JsonName = \"", strings.ReplaceAll(c.jsonFile, "\\", "/"), "\";")
		writer.WriteString("\n")
	}

	writeStringLn(writer, childIndent, dataClassGeneratedWarning)
	for _, field := range fields {
		err := field.genCSScript(writer, childIndent)
		if err != nil {
			return fmt.Errorf("field %s export failed: %s", field.name, err)
		}
		writer.WriteString("\n")
	}

	writeStringLn(writer, indent, "}")

	if isRoot {
		writeStringLn(writer, "}")
	}

	return nil
}
