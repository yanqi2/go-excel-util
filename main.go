package main

import (
	"fmt"
	"github.com/extrame/xls"
	//"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
)

type MyParser struct {
}

func (*MyParser) Parse(value string) string {
	return value + " parser"
}

func main() {

}

func runExcelizeUtilDemo(){
	// 列名
	//cells := []excelizeUtil.Cell{
	//	{
	//		CellTitle: "编号",
	//	},
	//	{
	//		CellTitle: "姓名",
	//	},
	//	{
	//		CellTitle: "手机号",
	//	},
	//	{
	//		CellTitle: "粤康码",
	//		CellType: excelizeUtil.Type_Picture,
	//		Colspan: 2,
	//	},
	//}
	//
	//// 数据
	//rows := make([]excelizeUtil.Row, 0)
	//
	//row := make([]excelizeUtil.RowValue, 0)
	//
	//row = append(row, excelizeUtil.RowValue{
	//	RowValue:  "1",
	//})
	//
	//row = append(row, excelizeUtil.RowValue{
	//	RowValue:  "张三",
	//	RowParser: &MyParser{},
	//})
	//
	//row = append(row, excelizeUtil.RowValue{
	//	RowValue:  "138000",
	//})
	//
	//row = append(row, excelizeUtil.RowValue{
	//	//RowValue:  "https://www.zhifure.com/upload/images/2018/7/20193719750.jpg",
	//	RowValue:  "http://biu-cn.dwstatic.com/marki/20210123/4a54af58ab80fe6d6fd2d181ecddee44.jpg",
	//})
	//
	////row = append(row, excelizeUtil.RowValue{
	////	//RowValue:  "https://www.zhifure.com/upload/images/2018/7/20193719750.jpg",
	////	RowValue:  "http://cms-bucket.ws.126.net/2020/0409/e7172ef0j00q8ht09001yc000hi00kmc.jpg",
	////})
	//
	//rows = excelizeUtil.AppendRow(rows, row)
	//
	//fmt.Println("go")
	//f, err := excelizeUtil.Generate("登记表", true, cells, rows)
	//if err != nil {
	//	fmt.Println(err.Error())
	//} else {
	//	if err = f.SaveAs("demo9.xlsx"); err != nil {
	//		fmt.Println(err.Error())
	//	} else {
	//		fmt.Println("success")
	//	}
	//}

	//_, err := excelize.OpenFile("excel/999.xls")
	//if err != nil {
	//	fmt.Printf(err.Error())
	//}
}


func openXls(path string) (err error) {
	xlFile, err := xls.Open(path, "utf-8")
	//fmt.Printf("err:%v\n", err)
	sheet := xlFile.GetSheet(0)
	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		if row.LastCol() > 0 {
			for j := 0; j < row.LastCol(); j++ {
				col := row.Col(j)
				fmt.Printf("%s ", col)
			}
		}
		fmt.Printf("\n")
	}
	return
}