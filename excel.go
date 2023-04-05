package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

func readExcelFile(filePath string) (*excelize.File, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	return f, nil
}

func writeExcelCell(file *excelize.File,
	sheetName string,
	cellName string,
	cellValue string,
	formula string,
	cellType excelize.CellType,
	styleIdx int) error {

	//err := file.SetCellStyle(sheetName, cellName, cellName, styleIdx)
	//if err != nil {
	//	return err
	//}

	if formula != "" {
		file.SetCellFormula(sheetName, cellName, formula)
	} else {
		// 根据类型执行不同的操作
		switch cellType {
		case excelize.CellTypeBool:
			boolValue, _ := strconv.ParseBool(cellValue)
			file.SetCellValue(sheetName, cellName, boolValue)
		case excelize.CellTypeDate:
			fmt.Println(cellValue)
			//file.SetCellValue(sheetName, cellName, time)
		case excelize.CellTypeError:
			file.SetCellValue(sheetName, cellName, "#ERROR")
		case excelize.CellTypeFormula:
			file.SetCellFormula(sheetName, cellName, cellValue)
		case excelize.CellTypeInlineString:
			file.SetCellRichText(sheetName, cellName, []excelize.RichTextRun{
				excelize.RichTextRun{
					Text: cellValue,
				},
			})
		case excelize.CellTypeNumber:
			numberValue, _ := strconv.ParseFloat(cellValue, 64)
			file.SetCellValue(sheetName, cellName, numberValue)
		case excelize.CellTypeSharedString:
			file.SetCellValue(sheetName, cellName, cellValue)
		case excelize.CellTypeUnset: //绝大多数默认单元格都是unset 有些麻烦
			if intValue, err := strconv.ParseInt(cellValue, 10, 32); err == nil {
				file.SetCellValue(sheetName, cellName, intValue)
			} else if floatValue, err := strconv.ParseFloat(cellValue, 64); err == nil {
				file.SetCellValue(sheetName, cellName, floatValue)
			} else if boolValue, err := strconv.ParseBool(cellValue); err == nil {
				file.SetCellBool(sheetName, cellName, boolValue)
			} else {
				file.SetCellValue(sheetName, cellName, cellValue)
			}
		default:
			file.SetCellValue(sheetName, cellName, cellValue)
		}
	}
	return nil
}
