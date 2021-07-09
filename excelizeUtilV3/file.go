package excelizeUtilV3

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"strings"
)

/****************
	  常量
*****************
*************/

const (
	vertical   = MergeType("vertical")   // 垂直方向合并
	horizontal = MergeType("horizontal") // 水平方向合并

	intCell       = ColType("int")
	boolCell      = ColType("bool")
	floatCell     = ColType("float")
	formulaCell   = ColType("formula")
	richTextCell  = ColType("richText")
	stringCell    = ColType("string")
	hyperLinkCell = ColType("hyperLink")
	pictureCell   = ColType("picture")

	defaultPre      = ""
	nestStructSplit = "#"
)

// 检测 ColType 是否属于预定义项
func colTypeUnknown(c ColType) bool {
	switch c {
	case intCell, boolCell, floatCell, formulaCell, richTextCell, stringCell, hyperLinkCell, pictureCell:
		break
	default:
		// errors.New("unknown cell type: " + string(c))
		return true
	}
	return false
}

// 检测 MergeType 是否属于预定义项
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

type MergeType string
type ColType string

type File struct {
	*excelize.File
	curSheet *Sheet
	sheetMap map[string]*Sheet
	sheets   []*Sheet
}

type Sheet struct {
	name string
	meta *SheetMeta
	rows []*RowValue
}

type SheetMeta struct {
	dataType reflect.Type
	colMap   map[string]*ColMeta
	cols     []*ColMeta
}

type ColMeta struct {
	fieldName string  // 字段名
	header    string  // 标题
	colType   ColType // 字段类型
	isMerge   bool    // 标题是否需要合并
	mergeCnt  int     // 合并的个数
	mergeType MergeType
}

type RowValue struct {
	sheet *Sheet
	row   map[string]*ColValue
}

type ColValue struct {
	sheet     *Sheet
	value     interface{} // 可能是集合类型
	isMerge   bool        // 是否需要合并（只能纵向合并）
	mergeType MergeType
	mergeCnt  int // 合并的个数
	pre       string
	isMulti   bool
	valueCnt  int
}

type MultiValue struct {
	value       interface{}
	parentIndex int
	isMerge     bool
	mergeCnt    int // 合并的个数
}

type MultiReflectValue struct {
	value       reflect.Value
	index       int
	parentIndex int
}

/****************
	  公开方法：File
*****************
*************/

// 创建一个file文件
func NewFile() *File {
	// 初始化属性
	f := &File{
		File:     excelize.NewFile(),
		sheetMap: make(map[string]*Sheet),
		sheets:   make([]*Sheet, 0),
	}
	// 删除默认sheet
	f.DeleteSheet(f.GetSheetName(0))
	return f
}

func (f *File) validData(dataType reflect.Type, data []interface{}) bool {
	return true
}

// 添加一个sheet页
func (f *File) AddSheet(sheetName string, dataType reflect.Type, data []interface{}) error {
	// 验证data
	if !f.validData(dataType, data) {
		return errors.New("data invalid")
	}

	// 创建新的 sheet
	sheet := f.newSheet(sheetName)

	// 设置meta
	if err := sheet.setMeta(dataType, defaultPre); err != nil {
		return err
	}

	// 设置rows
	if err := sheet.setRows(data); err != nil {
		return err
	}

	// 设置合并
	if err := sheet.setMerge(data); err != nil {
		return err
	}

	return nil
}

func (f *File) SaveAs(name string, opt ...excelize.Options) error {
	style, _ := f.NewStyle(`{"font":{"bold":true}, "alignment":{"horizontal":"center"}}`)
	styleLine, _ := f.NewStyle(`{"font":{"color":"#1265BE","underline":"single"}}`)

	for i, sheet := range f.sheets {
		f.NewSheet(sheet.name)
		f.SetActiveSheet(i)

		cols := sheet.meta.cols
		for j, row := range sheet.rows {
			for _, col := range cols {
				value := row.row[col.fieldName]
				tpe := col.colType

				// 判断value是否是集合类型
				if value.isMulti {

				}

				switch tpe {
				case intCell:
					break
				case boolCell:
					break
				case floatCell:
					break
				case formulaCell:
					break
				case richTextCell:
					break
				case stringCell:
					break
				case hyperLinkCell:
					break
				case pictureCell:
					break
				default:
					break
				}
			}
		}
	}
}

/****************
	  私有方法：File
*****************
*************/

// 创建一个新的 sheet，添加到 file 对象中
func (f *File) newSheet(sheetName string) *Sheet {
	sheet := &Sheet{
		name: sheetName,
		meta: &SheetMeta{
			colMap: make(map[string]*ColMeta),
			cols:   make([]*ColMeta, 0),
		},
		rows: make([]*RowValue, 0),
	}
	// 添加此sheet到file对象
	f.sheetMap[sheetName] = sheet
	f.sheets = append(f.sheets, sheet)
	f.curSheet = sheet

	return sheet
}

/****************
	  私有方法：Sheet
*****************
*************/

func (s *Sheet) setMeta(dataType reflect.Type, pre string) error {
	if dataType.Kind() == reflect.Ptr {
		dataType = dataType.Elem()
	}

	s.meta.dataType = dataType

	for i := 0; i < dataType.NumField(); i++ {
		fieldName := getMetaFieldName(dataType.Field(i), pre)
		fieldType := dataType.Field(i).Type

		celMeta := &ColMeta{
			fieldName: fieldName,
		}

		// TODO ignore field

		switch fieldType.Kind() {
		case reflect.Ptr:
			// 指针类型，递归获取类信息
			if err := s.setMeta(fieldType.Elem(), fieldName); err != nil {
				return err
			}
			break
		case reflect.Struct:
			// 结构体类型，递归获取类信息
			if err := s.setMeta(fieldType, fieldName); err != nil {
				return err
			}
			break
		case reflect.Array, reflect.Slice:
			// 数组类型
			switch fieldType.Elem().Kind() {
			case reflect.Ptr, reflect.Struct:
				if err := s.setMeta(fieldType.Elem(), fieldName); err != nil {
					return err
				}
				break
			default:
				if err := celMeta.setByTag(dataType.Field(i).Tag, true); err != nil {
					return err
				}
				s.meta.colMap[fieldName] = celMeta
				s.meta.cols = append(s.meta.cols, celMeta)
				break
			}
			break
		default:
			if err := celMeta.setByTag(dataType.Field(i).Tag, false); err != nil {
				return err
			}
			s.meta.colMap[fieldName] = celMeta
			s.meta.cols = append(s.meta.cols, celMeta)
			break
		}
	}

	return nil
}

func (s *Sheet) setRows(data []interface{}) error {
	for _, item := range data {
		s.newRow(item)
	}

	return nil
}

func (s *Sheet) setMerge(data []interface{}) error {
	for i, item := range data {
		_ = s.testMerge(i, 0, item, "")
	}

	return nil
}

func (s *Sheet) testMerge(rowIdx int, multiIdx int, item interface{}, pre string) int {
	rows := 1

	itemValue := reflect.ValueOf(item)
	if itemValue.Kind() == reflect.Ptr {
		itemValue = itemValue.Elem()
	}

	// 先遍历 field
	arrayFieldName := make([]string, 0)
	for i := 0; i < itemValue.NumField(); i++ {

		itemField := itemValue.Field(i)
		itemType := reflect.TypeOf(item)
		if itemType.Kind() == reflect.Ptr {
			itemType = itemType.Elem()
		}
		fieldName := getMetaFieldName(itemValue.Type().Field(i), pre)
		// 判断字段的类型
		if valueIsNil(itemField) {
			// nil 无需解析
			continue
		}
		if itemField.Kind() == reflect.Ptr {
			itemField = itemField.Elem()
		}

		if itemField.Kind() == reflect.Array || itemField.Kind() == reflect.Slice {
			if valueIsNil(itemField) {
				continue
			}
			arrayFieldName = append(arrayFieldName, itemType.Field(i).Name)

			itemFieldType := itemField.Type().Elem()
			if itemFieldType.Kind() == reflect.Ptr {
				itemFieldType = itemFieldType.Elem()
			}
			if itemFieldType.Kind() == reflect.Struct {
				fieldRows := 0
				// 是对象，还需要再次遍历
				for j := 0; j < itemField.Len(); j++ {
					fieldRows += s.testMerge(rowIdx, j, itemField.Index(j).Interface(), fieldName)
				}
				if fieldRows > rows {
					rows = fieldRows
				}
			} else {
				fieldRows := itemField.Len()
				if fieldRows > rows {
					rows = fieldRows
				}
			}
		} else if itemField.Kind() == reflect.Struct {
			// 是类型对象
			fieldRows := s.testMerge(rowIdx, 0, itemField.Interface(), fieldName)
			if fieldRows > rows {
				rows = fieldRows
			}
		}
	}
	if rows > 1 {
		for fieldName, col := range s.rows[rowIdx].row {
			nestFieldName := getNestFieldName(fieldName)
			if col.pre == pre {
				col.isMerge = true
				if multiIdx == 0 {
					col.mergeCnt = rows
				} else {
					if col.isMulti &&
						(reflect.TypeOf(col.value).Kind() == reflect.Array || reflect.TypeOf(col.value).Kind() == reflect.Slice) {
						cv := reflect.ValueOf(col.value)
						thisValue := cv.Index(multiIdx)
						if v, ok := thisValue.Interface().(*MultiValue); ok {
							skip := false
							for _, af := range arrayFieldName {
								if af == nestFieldName[len(nestFieldName)-1] {
									// 集合类型要略过，不合并
									skip = true
								}
							}
							if !skip {
								v.isMerge = true
								v.mergeCnt = rows
							}
						}
					}
				}
			}
		}
	}

	return rows
}

// 创建一个新的 row，添加到 sheet 对象中
func (s *Sheet) newRow(item interface{}) {
	colsMeta := s.meta.cols

	row := &RowValue{
		sheet: s,
		row:   make(map[string]*ColValue),
	}
	itemValue := reflect.ValueOf(item)
	if valueIsNil(itemValue) {
		// nil 值
		return
	}
	if itemValue.Kind() == reflect.Ptr {
		itemValue = itemValue.Elem()
	}
	//	通过反射赋值
	for _, colMeta := range colsMeta {
		row.newCol(itemValue, colMeta)
	}

	s.rows = append(s.rows, row)
}

func (r *RowValue) newCol(itemValue reflect.Value, colMeta *ColMeta) {
	colValue := &ColValue{
		sheet:    r.sheet,
		valueCnt: 1,
		mergeCnt: 1,
	}
	nestFieldName := getNestFieldName(colMeta.fieldName)
	if len(nestFieldName) >= 2 {
		colValue.pre = strings.Join(nestFieldName[0:len(nestFieldName)-1], nestStructSplit)
	}

	colValue.setValue(itemValue, colMeta.fieldName)

	r.row[colMeta.fieldName] = colValue
}

func (c *ColValue) setValue(itemValue reflect.Value, colFieldName string) {
	if valueIsNil(itemValue) {
		return
	}
	// 判断是否需要合并
	isMerge := isSplitValue(itemValue, colFieldName)
	if false == isMerge {
		// 不需要合并(单值)
		c.value = getSingleValue(itemValue, colFieldName)
		return
	}
	// 需要合并（多值）
	multiFieldValue := make([]MultiReflectValue, 0)
	// 初始 multiFieldValue
	if itemValue.Kind() == reflect.Array || itemValue.Kind() == reflect.Slice {
		// 如果此属性本身就是多值类型
		for i := 0; i < itemValue.Len(); i++ {
			mrv := MultiReflectValue{}
			mrv.index = i
			mrv.value = itemValue.Index(i)
			if itemValue.Index(i).Kind() == reflect.Ptr {
				mrv.value = itemValue.Index(i).Elem()
			}
			multiFieldValue = append(multiFieldValue, mrv)
		}
	} else {
		mrv := MultiReflectValue{}
		mrv.index = 0
		mrv.value = itemValue
		multiFieldValue = append(multiFieldValue, mrv)
	}

	// 分解层级关系
	for _, fieldNameItem := range getNestFieldName(colFieldName) {
		thisMultiFieldValue := make([]MultiReflectValue, 0)
		for _, fieldItem := range multiFieldValue {
			fieldValue := fieldItem.value.FieldByName(fieldNameItem)
			if fieldValue.Kind() == reflect.Ptr {
				fieldValue = fieldValue.Elem()
			}

			if fieldValue.Kind() == reflect.Array || fieldValue.Kind() == reflect.Slice {
				// 是集合对象
				for i := 0; i < fieldValue.Len(); i++ {
					mrv := MultiReflectValue{}
					mrv.index = i
					mrv.value = fieldValue.Index(i)
					mrv.parentIndex = fieldItem.index
					if fieldValue.Index(i).Kind() == reflect.Ptr {
						mrv.value = fieldValue.Index(i).Elem()
					}

					thisMultiFieldValue = append(thisMultiFieldValue, mrv)
				}
			} else {
				mrv := MultiReflectValue{}
				mrv.index = 0
				mrv.value = fieldValue
				mrv.parentIndex = fieldItem.index

				thisMultiFieldValue = append(thisMultiFieldValue, mrv)
			}
		}

		multiFieldValue = thisMultiFieldValue
	}

	cv := make([]interface{}, 0)
	for _, fieldItem := range multiFieldValue {
		mv := &MultiValue{
			value:       fieldItem.value.Interface(),
			parentIndex: fieldItem.parentIndex,
		}
		cv = append(cv, mv)
	}
	c.value = cv
	c.valueCnt = len(cv)
	c.isMulti = true
}

/****************
	  私有方法：ColMeta
*****************
*************/

// 从 struct tag 读取配置，设置 colMeta 对象
func (c *ColMeta) setByTag(tag reflect.StructTag, arr bool) error {
	// 获取名称
	if tagName := tag.Get("name"); tagName == "" {
		return errors.New("struct field not specified name")
	} else {
		c.header = tagName
	}

	// 获取类型
	if tagType := tag.Get("type"); tagType != "" {
		c.colType = ColType(tagType)
		if colTypeUnknown(c.colType) {
			return errors.New("unknown cell type: " + string(c.colType))
		}
	}

	if arr {
		// 获取合并类型
		if tagMerge := tag.Get("merge"); tagMerge != "" {
			c.mergeType = MergeType(tagMerge)
		}
		if mergeTypeUnknown(c.mergeType) {
			return errors.New("unknown merge type: " + string(c.mergeType))
		}
	}

	return nil
}

/****************
	  工具方法
*****************
*************/

func getMetaFieldName(field reflect.StructField, pre string) string {
	if pre == "" {
		return field.Name
	}
	return pre + nestStructSplit + field.Name
}

// 获取单个Field（链路中没有集合对象）
func getSingleField(v reflect.Value, colFieldName string) reflect.Value {
	if v.Kind() == reflect.Array || v.Kind() == reflect.Slice {
		panic("getSingleField have multi value")
	}

	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}

	if isNestStruct(colFieldName) {
		// 解析#
		for _, fieldNameItem := range getNestFieldName(colFieldName) {
			// 获取字段
			v = v.FieldByName(fieldNameItem)
			if valueIsNil(v) {
				return v
			}

			if v.Kind() == reflect.Ptr {
				v = v.Elem()
			}
			if v.Kind() == reflect.Array || v.Kind() == reflect.Slice {
				panic("getSingleField have multi value")
			}
		}
		return v
	}
	// 获取字段
	return v.FieldByName(colFieldName)
}

// 获取单个值（链路中没有集合对象）
func getSingleValue(v reflect.Value, fieldName string) interface{} {
	field := getSingleField(v, fieldName)

	if valueIsNil(v) {
		return ""
	}
	return field.Interface()
}

// 返回 value 是空指针
func valueIsNil(v reflect.Value) bool {
	if v.Kind() == reflect.Map || v.Kind() == reflect.Ptr || v.Kind() == reflect.Slice {
		if v.IsNil() {
			return true
		}
	}
	return false
}

// 判断 colFieldName 是否为嵌套类型
func isNestStruct(colFieldName string) bool {
	return strings.Contains(colFieldName, nestStructSplit)
}

// 获取嵌套类型
func getNestFieldName(colFieldName string) []string {
	return strings.Split(colFieldName, nestStructSplit)
}

// 判断是否为多值类型
func haveMultiValue(v reflect.Value, colFieldName string) bool {
	if v.Kind() == reflect.Array || v.Kind() == reflect.Slice {
		return true
	}

	fieldNameGroup := getNestFieldName(colFieldName)
	for _, fieldNameItem := range fieldNameGroup {
		// 获取字段值
		v = v.FieldByName(fieldNameItem)

		if valueIsNil(v) {
			// nil 空指针，当做非数组处理
			return false
		}

		if v.Kind() == reflect.Ptr {
			v = v.Elem()
		}

		if v.Kind() == reflect.Array || v.Kind() == reflect.Slice {
			// 是数组
			return true
		}
	}

	return false
}

// 判断是否需要分开的列
func isSplitValue(v reflect.Value, colFieldName string) bool {
	// 集合对象需要合并
	return haveMultiValue(v, colFieldName)
}

// 判断是否需要合并
func needCountMerge(pre string, subPre string, mergedPre []string) bool {
	// 两者是同一级，不需要合并
	if pre == subPre {
		return false
	}
	// 判断是否为上级
	if strings.Contains(pre, subPre) {
		// 判断此pre是否被计算过
		for _, mergedPreItem := range mergedPre {
			if mergedPreItem == pre {
				// 被计算过
				return false
			}
		}
		return true
	}
	return false
}
