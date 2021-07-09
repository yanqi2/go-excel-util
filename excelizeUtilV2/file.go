package excelizeUtilV2

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"strings"
)

type MergeType string
type ColumnType string

const (
	vertical   = MergeType("vertical")   // 垂直方向合并
	horizontal = MergeType("horizontal") // 水平方向合并

	intCell       = ColumnType("int")
	boolCell      = ColumnType("bool")
	floatCell     = ColumnType("float")
	formulaCell   = ColumnType("formula")
	richTextCell  = ColumnType("richText")
	stringCell    = ColumnType("string")
	hyperLinkCell = ColumnType("hyperLink")
	pictureCell   = ColumnType("picture")
)

func celTypeUnknown(c ColumnType) bool {
	switch c {
	case intCell, boolCell, floatCell, formulaCell, richTextCell, stringCell, hyperLinkCell, pictureCell:
		break
	default:
		// errors.New("unknown cell type: " + string(c))
		return true
	}
	return false
}

func mergeTypeUnknown(m MergeType) bool {
	switch m {
	case vertical, horizontal:
		break
	default:
		// errors.New("unknown merge type: " + string(m))
		return true
	}
	return false
}

/****************
	  类型
*****************
*************/

type File struct {
	*excelize.File
	sheetMeta      map[string]map[string]*ColMeta
	sheetCelName   map[string][]string
	sheetDType     map[string]reflect.Type
	sheetCellValue map[string][]*CellValue
	//sheetCellMerge map[string]map[string]
	//sheetIgnoreField map[string][]string
}

type ColMeta struct {
	Name    string
	ColType ColumnType
	Merge   MergeType
}

type Sheet struct {
	Name  string
	dType reflect.Type
	data  []interface{}
}

type MergeCell struct {
	hcell string
	vcell string
}

type CellValue struct {
	colPoint string
	pre      string
	celMeta  *ColMeta
	value    interface{}
}

/****************
	 公开方法
*****************
*************/

// 创建一个file文件
func NewFile() *File {
	f := &File{
		File:           excelize.NewFile(),
		sheetMeta:      make(map[string]map[string]*ColMeta),
		sheetCelName:   make(map[string][]string),
		sheetDType:     make(map[string]reflect.Type),
		sheetCellValue: make(map[string][]*CellValue),
	}
	// 删除默认sheet
	f.DeleteSheet(f.GetSheetName(0))
	return f
}

func (f *File) fNewSheet(sheetName string, dType reflect.Type) {
	f.sheetMeta[sheetName] = make(map[string]*ColMeta)
	f.sheetCelName[sheetName] = make([]string, 0)
	f.sheetDType[sheetName] = dType
	f.sheetCellValue[sheetName] = make([]*CellValue, 0)
	f.NewSheet(sheetName)
}

// 单个sheet调用此方法
func (f *File) SetData(sheetName string, dType reflect.Type, data []interface{}) error {
	// 创建sheet
	f.fNewSheet(sheetName, dType)

	// 解析 dType 配置信息
	if err := f.parseMeta(sheetName, dType, ""); err != nil {
		return err
	}

	// 赋值sheet所有的单元格
	return f.setData(sheetName, data)
}

func (f *File) setData(sheetName string, data []interface{}) error {
	dType := f.sheetDType[sheetName]
	for row, item := range data {
		if f.dataTypeDifferent(dType, item) {
			return errors.New("struct type is different")
		}

		dValue := reflect.ValueOf(data)
		// 通过反射设置单元格的值
		for col, celName := range f.sheetCelName[sheetName] {
			colPoint, _ := excelize.CoordinatesToCellName(col, row)
			if err := f.setCelData(sheetName, colPoint, dValue, celName); err != nil {
				return err
			}
		}
	}
	return nil
}

func (f *File) dataTypeDifferent(sheetCelType reflect.Type, item interface{}) bool {
	itemType := reflect.TypeOf(item)
	if itemType.Kind() == reflect.Ptr {
		itemType = itemType.Elem()
	}
	return false == (itemType == sheetCelType)
}

func (f *File) setCelData(sheetName, colPoint string, dValue reflect.Value, celName string) error {
	value := f.getValue(dValue, celName)
	colMeta := f.sheetMeta[sheetName][celName]

	// 判断是否需要合并
	if f.cellNeedMerge(dValue) {
		mergeCnt := dValue.Len()
		f.setCellMerge(mergeCnt, colMeta.Merge)
		return nil
	}

	f.sheetCellValue[sheetName] = append(f.sheetCellValue[sheetName], &CellValue{
		colPoint: colPoint,
		celMeta:  colMeta,
		value:    value,
	})

	return nil
}

func (f *File) setCellMerge(mergeCnt int, mergeType MergeType) {
	switch mergeType {
	case vertical:

		break
	case horizontal:

		break
	default:
		// 默认通过 horizontal 合并
		break
	}
}

// 多个sheet调用此方法
//func (f *File) SetSheet(sheets []*Sheet) error {
//	return nil
//}

/****************
	 私有方法
*****************
*************/

func (f *File) parseMeta(sheetName string, dataType reflect.Type, pre string) error {
	if dataType.Kind() == reflect.Ptr {
		dataType = dataType.Elem()
	}
	for i := 0; i < dataType.NumField(); i++ {
		celMeta := &ColMeta{}

		celName := getMetaCelName(dataType.Field(i), pre)
		fieldType := dataType.Field(i).Type

		// TODO ignore field

		switch fieldType.Kind() {
		case reflect.Ptr:
			// 指针类型，递归获取类信息
			if err := f.parseMeta(sheetName, fieldType.Elem(), celName); err != nil {
				return err
			}
			break
		case reflect.Struct:
			// 结构体类型，递归获取类信息
			if err := f.parseMeta(sheetName, fieldType, celName); err != nil {
				return err
			}
			break
		case reflect.Array, reflect.Slice:
			// 数组类型
			switch fieldType.Elem().Kind() {
			case reflect.Ptr, reflect.Struct:
				if err := f.parseMeta(sheetName, fieldType.Elem(), celName); err != nil {
					return err
				}
				break
			default:
				if err := celMeta.setByTag(dataType.Field(i).Tag, true); err != nil {
					return err
				}
				break
			}
			break
		default:
			if err := celMeta.setByTag(dataType.Field(i).Tag, false); err != nil {
				return err
			}
			break
		}

		f.sheetMeta[sheetName][celName] = celMeta
		f.sheetCelName[sheetName] = append(f.sheetCelName[sheetName], celName)
	}

	return nil
}

func getMetaCelName(field reflect.StructField, pre string) string {
	if pre == "" {
		return field.Name
	}
	return pre + "#" + field.Name
}

func (c *ColMeta) setByTag(tag reflect.StructTag, arr bool) error {
	// 获取名称
	if tagName := tag.Get("name"); tagName == "" {
		return errors.New("struct field not specified name")
	} else {
		c.Name = tagName
	}

	// 获取类型
	if tagType := tag.Get("type"); tagType != "" {
		c.ColType = ColumnType(tagType)
		if celTypeUnknown(c.ColType) {
			return errors.New("unknown cell type: " + string(c.ColType))
		}
	}

	if arr {
		// 获取合并类型
		if tagMerge := tag.Get("merge"); tagMerge != "" {
			c.Merge = MergeType(tagMerge)
		}
		if mergeTypeUnknown(c.Merge) {
			return errors.New("unknown merge type: " + string(c.Merge))
		}
	}

	return nil
}

func (f *File) getValue(v reflect.Value, fieldName string) interface{} {
	if strings.Contains(fieldName, "#") {
		// 解析#
		fieldNameGroup := strings.Split(fieldName, "#")
		for _, fieldNameItem := range fieldNameGroup {
			v = v.FieldByName(fieldNameItem)
			if v.Kind() == reflect.Ptr {
				v = v.Elem()
			}
		}
		return v.Interface()
	}
	return v.FieldByName(fieldName).Interface()
}

func (f *File) cellNeedMerge(v reflect.Value) bool {
	if v.Kind() == reflect.Array || v.Kind() == reflect.Slice {
		// 集合类型
		return true
	}
	return false
}
