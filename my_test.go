package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"math"
	"math/rand"
	"os"
	"strconv"
	"strings"
	"testing"
	"time"
)

func Test01(t *testing.T) {
	beginTime := time.Now()
	endTime := beginTime.Add(20 * time.Hour)

	if endTime.Unix()-beginTime.Unix() == int64(20*time.Hour) {
		fmt.Println("yes")
	}
}

func Test02(t *testing.T) {
	file := excelize.NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	styleID, err := file.NewStyle(`{"font":{"color":"#777777"}}`)
	if err != nil {
		fmt.Println(err)
	}
	if err := streamWriter.SetRow("A1", []interface{}{
		excelize.Cell{StyleID: styleID, Value: "Data"}}); err != nil {
		fmt.Println(err)
	}
	for rowID := 2; rowID <= 102400; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := streamWriter.SetRow(cell, row); err != nil {
			fmt.Println(err)
		}
	}
	if err := streamWriter.Flush(); err != nil {
		fmt.Println(err)
	}
	if err := file.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func Test03(t *testing.T) {
	file := excelize.NewFile()
	_ = file.SetColWidth("Sheet1", "A1", "A1", 30)
	_ = file.SetRowHeight("Sheet1", 1, 100)

	_ = file.AddPicture("Sheet1", "A1", "image.jpg", `{"autofit": true, "x_offset": 10, "y_offset": 10}`)
	if err := file.SaveAs("Book4.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func Test04(t *testing.T) {
	file := excelize.NewFile()
	streamWriter, _ := file.NewStreamWriter("Sheet1")

	img, _ := ioutil.ReadFile("image.jpg")
	if err := streamWriter.SetRow("A1", []interface{}{
		excelize.Cell{Value: img}}); err != nil {
		fmt.Println(err)
	}
	if err := file.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func Test05(t *testing.T) {
	i := 1
	defer fmt.Println("end")
	go print(i)
}

func print(i int) {
	fmt.Println(i)
	time.Sleep(time.Duration(2) * time.Second)
	fmt.Println(i + 1)
}

func Test06(t *testing.T) {
	data := []int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

	count := int(math.Ceil(float64(10) / float64(3)))
	for i := 0; i < count; i++ {
		tail := i*3 + 3
		if tail > len(data) {
			tail = len(data)
		}

		thisData := data[i*3 : tail]
		fmt.Println(thisData)
	}
}

func Test07(t *testing.T) {
	uids := []int64{
		1, 2, 3,
	}

	uid := ""
	for _, v := range uids {
		uid += "," + strconv.FormatInt(v, 10)
	}
	uid = strings.Trim(uid, ",")
	fmt.Println(uid)
}

func Test08(t *testing.T) {
	s1 := []int64{1, 2, 3, 4, 5}
	s2 := []int64{3, 4, 5, 6, 7}
	s3 := differenceInt64Array(s1, s2)
	fmt.Println(s3)
}

func IntArrayToInt64Array(array []int) (result []int64) {
	result = make([]int64, 0)
	for _, v := range array {
		result = append(result, int64(v))
	}
	return
}

func Int64ArrayToIntArray(array []int64) (result []int) {
	result = make([]int, 0)
	for _, v := range array {
		result = append(result, int(v))
	}
	return
}

// 取差集，找出slice1中  slice2中没有的
func differenceInt64Array(slice1, slice2 []int64) []int64 {
	result := make([]int64, 0)
	slice2Length := len(slice2)
	for _, v1 := range slice1 {
		for j, v2 := range slice2 {
			if v1 == v2 {
				break
			}
			if j == slice2Length-1 {
				result = append(result, v1)
			}
		}
	}
	return result
}

func TestStreamWriteExcel(t *testing.T) {
	f := excelize.NewFile()
	sw, _ := f.NewStreamWriter("Sheet1")
	//file, _ := ioutil.ReadFile("image.jpg")
	sw.SetRow("A1", []interface{}{
		"1","2",
	})

	//fmt.Fprintf(sw.GetRawData(), `<row r="%d">`, 1)
	//if err := xml.EscapeText(sw.GetRawData(), file); err != nil {
	//	fmt.Println(err.Error())
	//}
	//sw.GetRawData().WriteString(`</row>`)
	//if err := sw.Flush(); err != nil {
	//	fmt.Println(err.Error())
	//}
	_ = f.SaveAs("5.xlsx")
}


func Test09(t *testing.T){
	t1 := time.Now()
	t2 := t1.Add(time.Duration(16) * time.Second)
	fmt.Println(t2.Unix() - t1.Unix())
}


func Test10(t *testing.T){
	if r := recover(); r != nil {
		//fmt.Println(r)
	}

	panic("error")
}

type MFile struct {
	*os.File
}

func Test11(t *testing.T) {
	_ = MFile{
		File: nil,
	}
}

func Test12(t *testing.T) {
	s := "{\"remarks\":[\"1\\n2\\n3\"],\"msgState\":\"1\"}"
	f,_ := fmt.Printf("%s", s)
	fmt.Println(f)
}