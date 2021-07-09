package excelizeUtil

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"strconv"
)

var rowIndex, colIndex int32

type CellType int32

type RowParser interface {
	Parse(value string) string
}

const (
	Type_Normal  CellType = 0
	Type_Picture CellType = 1
)

type Cell struct {
	CellTitle string
	CellType  CellType
	Colspan   int32
}

type RowValue struct {
	RowValue  string
	RowParser RowParser
	Rowspan   int32
}

type Row struct {
	RowValues []RowValue
}

func Generate(sheetName string, haveHeader bool, cells []Cell, rows []Row) (f *excelize.File, err error) {
	f = excelize.NewFile()

	if sheetName == "" {
		sheetName = "Sheet1"
	}
	f.NewSheet(sheetName)
	f.SetActiveSheet(0)

	if haveHeader {
		cellLength := len(cells)
		for i := 0; i < cellLength; i++ {
			// 获取单元格名
			cellName := getCurrentCellName()
			style,_ := f.NewStyle(`{
				"alignment":{
					"horizontal":"center"
				}
			}`)
			_ = f.SetCellStyle(sheetName, cellName, cellName, style)
			if err = f.SetCellValue(sheetName, cellName, cells[i].CellTitle); err != nil {
				return nil, err
			}
			if cells[i].Colspan > 0 {
				for j := 0; int32(j) < cells[i].Colspan-1; j++ {
					rearCells := append([]Cell{}, cells[colIndex:]...)
					cells = append(cells[0:colIndex], cells[i])
					cells = append(cells, rearCells...)
				}

				colIndex = colIndex + cells[i].Colspan - 1
				endCellName := getCurrentCellName()
				if err = f.MergeCell(sheetName, cellName, endCellName); err != nil {
					return nil, err
				}
			}
			colIndex += 1
		}
		rowIndex += 1
	}

	rowSize := len(rows)
	for i := 0; i < rowSize; i++ {
		resetColIndex()

		rowValues := rows[i].RowValues
		rowLength := len(rowValues)

		for j := 0; j < rowLength; j++ {
			cellName := getCurrentCellName()
			rowValue := rowValues[j].RowValue
			if rowValues[j].RowParser != nil {
				rowValue = rowValues[j].RowParser.Parse(rowValue)
			}
			cellType := cells[j].CellType
			switch cellType {
			case Type_Normal:
				if err = f.SetCellValue(sheetName, cellName, rowValue); err != nil {
					return nil, err
				}
				break
			case Type_Picture:
				if err = AddPicture(f, sheetName, cellName, rowValue, ""); err != nil {
					return nil, err
				}
			}

			// 合并单元格
			if rowValues[j].Rowspan > 0 {
				colIndex = colIndex + rowValues[j].Rowspan
				endCellName := getCurrentCellName()
				if err = f.MergeCell(sheetName, cellName, endCellName); err != nil {
					return nil, err
				}
			}

			colIndex += 1
		}
		rowIndex += 1
	}

	return f, nil
}

func getColName(index int32) string {
	if index/26 > 0 {
		return getColName(index/26-1) + getColName(index%26)
	} else {
		Slice := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
			"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
		return Slice[index]
	}
}

func getCurrentCellName() string {
	return getColName(colIndex) + strconv.Itoa(int(rowIndex+1))
}

func resetColIndex() {
	colIndex = 0
}

func GenerateSimple(sheetName string, cells []string, rows interface{}) {

}

func AppendRow(rows []Row, rowVal []RowValue) []Row {
	row := Row{
		RowValues: rowVal,
	}
	return append(rows, row)
}
