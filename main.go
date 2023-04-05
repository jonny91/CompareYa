// 比较两个excel冲突的命令行工具
//
// Author: 洪金敏
// Copyright (c) 2023, 洪金敏
// All rights reserved.
package main

import (
	"flag"
	"fmt"
	"github.com/mbndr/figlet4go"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
	"time"
)

func main() {
	asciiObj := figlet4go.NewAsciiRender()
	result, _ := asciiObj.Render("CompareYa")
	fmt.Println(result)
	var file1Path, file2Path, outputPath string
	conflictsPath := "compareYa_log_" + strconv.FormatInt(time.Now().Unix(), 10) + ".txt"
	flag.StringVar(&file1Path, "s1", "", "path of the first source Excel file")
	flag.StringVar(&file2Path, "s2", "", "path of the second source Excel file")
	flag.StringVar(&outputPath, "o", "", "path of the output Excel file")
	flag.Parse()
	// 读取第一个 Excel 文件
	file1, err := readExcelFile(file1Path)
	if err != nil {
		fmt.Printf("Failed to read Excel file %s: %s\n", file1Path, err)
		os.Exit(1)
	}

	// 读取第二个 Excel 文件
	file2, err := readExcelFile(file2Path)
	if err != nil {
		fmt.Printf("Failed to read Excel file %s: %s\n", file2Path, err)
		os.Exit(1)
	}
	// 声明一个新的 Excel 文件
	newFile := excelize.NewFile()
	sameSheet, diffSheet := getDiffSheetName(file1, file2)
	for _, diff := range diffSheet {
		err = combineDiff(file1, file2, newFile, diff)
		if err != nil {
			fmt.Printf("something wrong when merge file: %v", err)
			os.Exit(1)
		}
	}

	// 比较和合并两个 Excel 文件
	mergedFile, conflicts, err := compareAndMergeFiles(file1, file2, newFile, sameSheet)
	if err != nil {
		fmt.Printf("Failed to compare and merge Excel files: %s\n", err)
		os.Exit(1)
	}

	// 将合并后的结果写入新的 Excel 文件
	if err := mergedFile.SaveAs(outputPath); err != nil {
		fmt.Printf("Failed to save merged Excel file to %s: %s\n", outputPath, err)
		os.Exit(1)
	}

	// 将冲突项目写入单独的 txt 文件
	if err := writeConflictsToFile(conflicts, conflictsPath); err != nil {
		fmt.Printf("Failed to write conflicts to %s: %s\n", conflictsPath, err)
		os.Exit(1)
	}
}

func writeConflictsToFile(conflicts map[string]string, filePath string) error {
	f, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer f.Close()
	for cellName, values := range conflicts {
		valueArray := strings.Split(values, ",")
		value1 := valueArray[0]
		value2 := valueArray[1]
		line := fmt.Sprintf("%s: %s/%s\n", cellName, value1, value2)
		if _, err := f.WriteString(line); err != nil {
			return err
		}
	}

	return nil
}
