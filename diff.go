package main

import "github.com/xuri/excelize/v2"

type DiffResult struct {
	SheetName    string
	IndexInFile1 int
	IndexInFile2 int
}

func getDiffSheetName(file1, file2 *excelize.File) ([]*DiffResult, []*DiffResult) {
	var same []*DiffResult
	var diff []*DiffResult
	//比较file1 和 file2 中 一样和不同的sheetName 放到对应的数组中
	file1SheetCount := file1.SheetCount
	file2SheetCount := file2.SheetCount
	for idx1 := 0; idx1 < file1SheetCount; idx1++ {
		sheetName := file1.GetSheetName(idx1)
		//循环判断file2中每隔sheetName是否是file1中的
		ok := false
		idx2 := 0
		for ; idx2 < file2SheetCount; idx2++ {
			if file2.GetSheetName(idx2) == sheetName {
				ok = true
				break
			}
		}
		if ok {
			same = append(same, &DiffResult{
				sheetName,
				idx1,
				idx2,
			})
		} else {
			diff = append(diff, &DiffResult{
				sheetName,
				idx1,
				-1,
			})
		}
	}

	for idx2 := 0; idx2 < file2SheetCount; idx2++ {
		sheetName := file2.GetSheetName(idx2)
		ok := false
		for idx1 := 0; idx1 < file1SheetCount; idx1++ {
			if file1.GetSheetName(idx1) == sheetName {
				ok = true
				break
			}
		}
		if !ok {
			diff = append(diff, &DiffResult{
				sheetName,
				-1,
				idx2,
			})
		}
	}

	return same, diff
}
