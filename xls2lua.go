package main

import (
	"bufio"
	"bytes"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
	"sync"

	"github.com/tealeg/xlsx"
)

const (
	fieldDelimiter = ", "
	luaExtension   = ".lua"
	xlsxExtension  = ".xlsx"
	fieldInfoSplit = ":"
)

var (
	wg       *sync.WaitGroup
	confDir  string // 配置文件目录
	excelDir string // Excel目录
	luaOut   string // 输出目录
	luaFiles []string
)

func fixedPathSuffix(s *string) {
	if !strings.HasSuffix(*s, "/") {
		*s = *s + "/"
	}
}

func writeLuaFile(fileName string, text []byte) error {
	return ioutil.WriteFile(luaOut+fileName+luaExtension, []byte(text), 0666)
}

func initConfigs() {
	file, err := os.Open(confDir)
	if err != nil {
		panic(err)
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		luaFiles = append(luaFiles, parseFileName(strings.TrimSpace(scanner.Text())))
	}
	if err := scanner.Err(); err != nil {
		panic(err)
	}
	fixedPathSuffix(&luaOut)
	fixedPathSuffix(&excelDir)
	fmt.Printf("Excel表路径\t[%#v]\n", excelDir)
	fmt.Printf("Lua路径\t[%#v]\n", luaOut)
}

type xlsField struct {
	name  string // 字段名称
	fType string // 字段类型, number, bool, string, any
}

func parseField(s string) (*xlsField, error) {
	args := strings.Split(s, fieldInfoSplit)
	if len(args) < 2 {
		return nil, fmt.Errorf("解析字段错误,没有用:分割.例如fieldName:number [%v]", s)
	}
	name := strings.TrimSpace(args[0])
	if name == "" {
		return nil, fmt.Errorf("解析字段为空")
	}
	t := strings.ToLower(args[1])
	switch t {
	case "number", "bool", "string", "array":
	default:
		return nil, fmt.Errorf("解析字段类型出错,不是number,bool,string,array")
	}
	return &xlsField{name, t}, nil
}

func parseHeader(src string, headRow []*xlsx.Cell) (fields []*xlsField, headerSize int) {
	fields = make([]*xlsField, len(headRow))
	for i, cell := range headRow {
		fieldString := cell.Value
		fieldString = strings.TrimSpace(fieldString)
		if fieldString != "" && fieldString != "-" {
			field, err := parseField(fieldString)
			if field == nil {
				panic(fmt.Errorf("[%s]标题头格式问题[%s], [number, bool, string, array] [%v]", src, fieldString, err))
			}
			fields[i] = field
			headerSize++
		}
	}
	return
}

func fixedFloatType(cell string) string {
	dotIndex := strings.Index(cell, ".")
	if dotIndex != -1 && (len(cell)-dotIndex) > 3 {
		if f, err := strconv.ParseFloat(cell, 64); err == nil {
			newValue := fmt.Sprintf("%.3f", f)
			for {
				if strings.HasSuffix(newValue, "0") {
					newValue = newValue[:len(newValue)-1]
				} else {
					if strings.HasSuffix(newValue, ".") {
						newValue = newValue[:len(newValue)-1]
					}
					break
				}
			}
			return newValue
		}
	}
	if cell == "" {
		return "0"
	} else {
		return cell
	}
}

func fixedBoolType(cell string) string {
	x := strings.ToLower(cell)
	if x == "1" || x == "true" {
		return "true"
	} else {
		return "false"
	}
}

func parseRow(fields []*xlsField, cells []*xlsx.Cell) (error, string) {
	values := make([]string, 0, 128)
	idValue := ""
	for i, f := range fields {
		if f != nil {
			cellStr := cells[i].Value
			var writeValue string
			switch f.fType {
			case "number":
				writeValue = fixedFloatType(cellStr)
			case "string":
				writeValue = `"` + cellStr + `"`
			case "bool":
				writeValue = fixedBoolType(cellStr)
			case "array":
				writeValue = cellStr
			}
			if idValue == "" {
				idValue = writeValue
			}
			values = append(values, fmt.Sprintf(`["%v"]=%v`, f.name, writeValue))
		}
	}
	if idValue == "" {
		return fmt.Errorf("没有主键"), ""
	}
	return nil, fmt.Sprintf("    [%v] = {%s}", idValue, strings.Join(values, fieldDelimiter))
}

func xls2lua(fileName string) bool {
	sb := &strings.Builder{}
	sb.WriteString(fmt.Sprintf("开始处理[%s]\n", fileName))
	if wg != nil {
		defer wg.Done()
	}
	defer func(s *strings.Builder) {
		fmt.Printf(s.String())
	}(sb)
	xlFile, err := xlsx.OpenFile(excelDir + fileName + xlsxExtension)
	if err != nil {
		if wg != nil {
			sb.WriteString(fmt.Sprintf("错误！没有找到[%v][%v]\n", fileName, err))
		}
		return false
	}
	buf := make([]byte, 0, 4096)
	buffer := bytes.NewBuffer(buf)
	buffer.WriteString("return {\n")
	sheet := xlFile.Sheets[0] // 只读第一个Sheet
	fields, fieldSize := parseHeader(fileName, sheet.Rows[1].Cells)
	rows := sheet.Rows[2:] // 忽略前两列（第一列，策划描述，第二列定义程序用的格式）
	for rindex, row := range rows {
		cellLen := len(row.Cells)
		if cellLen == 0 { // 有空行就结束了
			break
		}
		if cellLen < fieldSize {
			sb.WriteString(fmt.Sprintf("[%s]字段数量不匹配,标题数量[%d], 当前[%d]行的字段数量[%d]!\n", fileName, fieldSize, rindex+2, len(row.Cells)))
			return false
		}
		err, line := parseRow(fields, row.Cells)
		if err != nil {
			sb.WriteString(fmt.Sprintf("解析失败[%s][%v], 错误在第[%d]行!\n", fileName, err, rindex+2))
			return false
		}
		buffer.WriteString(line + ",\n")
	}
	buffer.WriteString("}")
	err = writeLuaFile(fileName, buffer.Bytes())
	if err != nil {
		sb.WriteString(fmt.Sprintf("写入文件失败[%s][%v]!\n", fileName, err))
		return false
	}
	sb.WriteString(fmt.Sprintf("写入[%s]成功!\n", fileName))
	return true
}

func parseFileName(x string) string {
	fmt.Println("File ", x)
	i := strings.LastIndex(x, ".")
	chars := x[:i]
	return chars
}

func findInArray(s string, strs []string) bool {
	for _, v := range strs {
		if s == v {
			return true
		}
	}
	return false
}

func outTables() {
	wg = &sync.WaitGroup{}
	for _, v := range luaFiles {
		wg.Add(1)
		go xls2lua(v)
	}
	wg.Wait()
}

func outOneTable() {
	fmt.Println("请输入表格名称:")
	for _, v := range luaFiles {
		fmt.Println(v)
	}
tabSelect:
	var tab string
	fmt.Scan(&tab)
	if !findInArray(tab, luaFiles) {
		fmt.Printf("错误！没有找到[%v], 请注意，大小写敏感!\n", tab)
		fmt.Println("请输入表格名称:")
		goto tabSelect
	} else {
		if !xls2lua(tab) {
			fmt.Println("表格输出失败")
		}
	}
}

func readInputs() {
	var i int
start:
	fmt.Println("1.全部表格")
	fmt.Println("2.单个表格")
	fmt.Printf("请输入:")
	fmt.Scan(&i)
	switch i {
	case 1:
		outTables()
	case 2:
		outOneTable()
	default:
		goto start
	}
}

func init() {
	flag.StringVar(&excelDir, "e", "./excel", "excel directory")
	flag.StringVar(&luaOut, "o", "./gameconf", "lua output directory")
	flag.StringVar(&confDir, "c", "./config", "config file path")
	flag.Parse()
}

func main() {
	initConfigs()
	fmt.Println()
	readInputs()
	fmt.Println("输入Q键退出")
	scanner := bufio.NewScanner(os.Stdin)
	for scanner.Scan() {
		line := scanner.Text()
		if line == "q" || line == "Q" {
			break
		}
	}
}
