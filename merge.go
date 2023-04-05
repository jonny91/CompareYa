package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strings"
)

func combineDiff(file1, file2, newFile *excelize.File, diff *DiffResult) error {
	_, err := newFile.NewSheet(diff.SheetName)
	if err != nil {
		return err
	}
	var srcFile *excelize.File
	var rows [][]string
	if diff.IndexInFile1 == -1 {
		srcFile = file2
	}

	if diff.IndexInFile2 == -1 {
		srcFile = file1
	}

	rows, err = srcFile.GetRows(diff.SheetName)
	if err != nil {
		return err
	}
	for rowIndex, row := range rows {
		for colIndex, cellValue := range row {
			cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			if err != nil {
				return err
			}
			formula, err := srcFile.GetCellFormula(diff.SheetName, cellName)
			if err != nil {
				return err
			}
			//深度拷贝对象
			styleIdx, err := srcFile.GetCellStyle(diff.SheetName, cellName)
			//style := srcFile.Styles.CellStyles.CellStyle[styleIdx]

			if err != nil {
				return err
			}
			cellType, err := srcFile.GetCellType(diff.SheetName, cellName)
			if err != nil {
				return err
			}

			err = writeExcelCell(newFile, diff.SheetName, cellName, cellValue, formula, cellType, styleIdx)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func compareAndMergeFiles(file1, file2, newFile *excelize.File, sameNameSheet []*DiffResult) (*excelize.File, map[string]string, error) {
	// 记录冲突
	conflicts := make(map[string]string)
	for _, same := range sameNameSheet {
		// 获取第一个 Excel 文件的所有行
		rows1, err := file1.GetRows(file1.GetSheetName(same.IndexInFile1))
		if err != nil {
			return nil, nil, err
		}

		// 获取第二个 Excel 文件的所有行
		rows2, err := file2.GetRows(file2.GetSheetName(same.IndexInFile2))
		if err != nil {
			return nil, nil, err
		}

		//在新文件中创建新的sheet
		_, err = newFile.NewSheet(same.SheetName)
		if err != nil {
			return nil, nil, err
		}

		// 将第一个 Excel 文件的数据添加到新文件
		for rowIndex, row := range rows1 {
			for colIndex, cellValue := range row {
				cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				if err != nil {
					return nil, nil, err
				}
				formula, err := file1.GetCellFormula(same.SheetName, cellName)
				if err != nil {
					return nil, nil, err
				}
				cellType, err := file1.GetCellType(same.SheetName, cellName)
				if err != nil {
					return nil, nil, err
				}

				//todo hjm cell style
				err = writeExcelCell(newFile, same.SheetName, cellName, cellValue, formula, cellType, -1)
				if err != nil {
					return nil, nil, err
				}
			}
		}

		// 将第二个 Excel 文件的数据添加到新文件，并与第一个文件的数据比较
		for rowIndex, row := range rows2 {
			for colIndex, cellValue := range row {
				cellName, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
				if err != nil {
					return nil, nil, err
				}

				// 如果新文件中已经存在该单元格的数据，则需要比较并解决冲突
				if value, err := newFile.GetCellValue(same.SheetName, cellName); err == nil {
					if value != cellValue {
						fmt.Printf("Conflict found in cell 【%s】 of sheet 【%s】, value1: 【%s】, value2: 【%s】\n",
							cellName, same.SheetName, value, cellValue)

						// 记录冲突
						conflicts[cellName+"@"+same.SheetName] = fmt.Sprintf("%s,%s", value, cellValue)
						//最后选择的版本文件
						var chosenFile *excelize.File
						// 询问用户采用哪个版本的数据
						chosenValue := ""
						for chosenValue == "" {
							fmt.Printf("Which value do you want to keep for cell %s? (1/%s  2/%s): ", cellName, value, cellValue)
							var input string
							fmt.Scanln(&input)

							if strings.TrimSpace(input) == "1" {
								chosenValue = value
								chosenFile = file1
							} else if strings.TrimSpace(input) == "2" {
								chosenValue = cellValue
								chosenFile = file2
							} else {
								fmt.Println("Invalid input, please try again.")
							}
						}

						formula, err := chosenFile.GetCellFormula(same.SheetName, cellName)
						if err != nil {
							return nil, nil, err
						}
						cellType, err := chosenFile.GetCellType(same.SheetName, cellName)
						if err != nil {
							return nil, nil, err
						}

						//todo hjm cell style
						err = writeExcelCell(newFile, same.SheetName, cellName, cellValue, formula, cellType, -1)
						if err != nil {
							return nil, nil, err
						}

					}
				} else {
					// 如果新文件中不存在该单元格的数据，则直接将第二个文件的数据添加到新文件
					formula, err := file2.GetCellFormula(same.SheetName, cellName)
					if err != nil {
						return nil, nil, err
					}
					cellType, err := file2.GetCellType(same.SheetName, cellName)
					if err != nil {
						return nil, nil, err
					}

					//todo hjm cell style
					err = writeExcelCell(newFile, same.SheetName, cellName, cellValue, formula, cellType, -1)
					if err != nil {
						return nil, nil, err
					}
				}
			}
		}
	}
	return newFile, conflicts, nil
}
