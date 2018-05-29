package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"sort"
	"strconv"
	"strings"
	"sync"

	"github.com/tealeg/xlsx"
)

const (
	fieldDelimiter = ", "
	luaExtension   = ".lua"
	fieldInfoSplit = ":"
)

type writeFunc func(string, []byte) error

var (
	wg              *sync.WaitGroup
	writeAllFunc    writeFunc
	writeClientFunc writeFunc
	writeServerFunc writeFunc
	config          = &Config{}
	excelDir        string // Excel目录
	serverOut       string // 服务器输出目录
	clientOut       string // 客户端输出目录
)

type (
	Config struct {
		Servers []string
		Clients []string
	}
)

func fixedPathSuffix(s *string) {
	if !strings.HasSuffix(*s, "/") {
		*s = *s + "/"
	}
}

func doChain(args ...func() error) {
	for _, v := range args {
		if err := v(); err != nil {
			panic(err)
		}
	}
}

func initConfigs() {
	var confFileData []byte
	doChain(
		func() error {
			var err error
			confFileData, err = ioutil.ReadFile("./config.json")
			return err
		},
		func() error {
			return json.Unmarshal(confFileData, config)
		},
	)

	fixedPathSuffix(&serverOut)
	fixedPathSuffix(&clientOut)
	fixedPathSuffix(&excelDir)
	writeServerFunc = func(fileName string, text []byte) error {
		return ioutil.WriteFile(serverOut+fileName+luaExtension, []byte(text), 0666)
	}
	writeClientFunc = func(fileName string, text []byte) error {
		return ioutil.WriteFile(clientOut+fileName+luaExtension, []byte(text), 0666)
	}
	writeAllFunc = func(fileName string, text []byte) error {
		if err := writeServerFunc(fileName, text); err != nil {
			return err
		}
		return writeClientFunc(fileName, text)
	}
	fmt.Printf("Excel表路径\t[%#v]\n", excelDir)
	fmt.Printf("服务器Lua路径\t[%#v]\n", serverOut)
	fmt.Printf("客户端Lua路径\t[%#v]\n", clientOut)
}

type xlsField struct {
	name  string // 字段名称
	fType string // 字段类型, number, bool, string, array
	cORs  byte   // 输出类型，0(默认全部输出), 1 客户端, 2服务器
}

func parseField(s string) *xlsField {
	args := strings.Split(s, fieldInfoSplit)
	if len(args) < 2 {
		return nil
	}
	name := strings.TrimSpace(args[0])
	if name == "" {
		return nil
	}
	t := strings.ToLower(args[1])
	switch t {
	case "number", "bool", "string", "array":
	default:
		return nil
	}
	cORs := byte(0)
	if len(args) > 2 {
		cORsStr := strings.ToLower(args[2])
		switch cORsStr {
		case "c":
			cORs = 1
		case "s":
			cORs = 2
		}
	}
	return &xlsField{name, t, cORs}
}

func parseHeader(src string, headRow []*xlsx.Cell) (fields []*xlsField, headerSize int) {
	fields = make([]*xlsField, len(headRow))
	for i, cell := range headRow {
		fieldString := cell.Value
		fieldString = strings.TrimSpace(fieldString)
		if fieldString != "-" {
			field := parseField(fieldString)
			if field == nil {
				panic(fmt.Errorf("[%s]标题头格式问题, [number, bool, string, array]", src))
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

func parseRow(fields []*xlsField, cORs byte, cells []*xlsx.Cell) (error, string) {
	values := make([]string, 0, 128)
	idValue := ""
	for i, f := range fields {
		if f != nil {
			if cORs != 0 && f.cORs != 0 && f.cORs != cORs {
				continue
			}
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
			values = append(values, fmt.Sprintf("%v=%v", f.name, writeValue))
		}
	}
	if idValue == "" {
		return fmt.Errorf("没有主键"), ""
	}
	return nil, fmt.Sprintf("[%v]={ %s }", idValue, strings.Join(values, fieldDelimiter))
}

// typ 0: all
// typ 1: client
// typ 2: server
func xls2lua(fileName string, cORs byte, w writeFunc) bool {
	if wg != nil {
		defer wg.Done()
	}
	xlFile, err := xlsx.OpenFile(excelDir + fileName + ".xlsx")
	if err != nil {
		if wg != nil {
			fmt.Printf("错误！没有找到[%v][%v]\n", fileName, err)
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
		if len(row.Cells) < fieldSize {
			fmt.Printf("输出失败[%s][字段数量少于标题数量], 错误在第[%d]行!\n", fileName, rindex+2+1)
			return false
		}
		err, line := parseRow(fields, cORs, row.Cells)
		if err != nil {
			fmt.Printf("输出失败[%s][%v], 错误在第[%d]行!\n", fileName, err, rindex+2+1)
			return false
		}
		buffer.WriteString(line + ",\n")
	}
	buffer.WriteString("}")
	err = w(fileName, buffer.Bytes())
	if err != nil {
		fmt.Printf("写入文件失败[%s][%v]!\n", fileName, err)
		return false
	}
	fmt.Printf("写入[%s]成功!\n", fileName)
	return true
}

func parseFileName(x string) string {
	fmt.Println("File ", x)
	i := strings.LastIndex(x, ".")
	chars := x[:i]
	return chars
}

func mergeArrays(src ...[]string) []string {
	set := make(map[string]struct{}, 16)
	ret := make([]string, 0, 16)
	for _, tab := range src {
		for _, v := range tab {
			if _, ok := set[v]; !ok {
				set[v] = struct{}{}
				ret = append(ret, parseFileName(v))
			}
		}
	}
	sort.Sort(sort.StringSlice(ret))
	return ret
}

func findInArray(s string, strs []string) bool {
	for _, v := range strs {
		if s == v {
			return true
		}
	}
	return false
}

func outTables(tables []string, cORs byte, w writeFunc) {
	wg = &sync.WaitGroup{}
	for _, v := range tables {
		wg.Add(1)
		go xls2lua(v, cORs, w)
	}
	wg.Wait()
}

func outOneTable() {
	var i int
start:
	fmt.Println("1.打全部表格")
	fmt.Println("2.打服务器表格")
	fmt.Println("3.打客户端表格")
	fmt.Printf("请输入:")
	fmt.Scan(&i)
	if !(i >= 1 && i <= 3) {
		goto start
	}
	fmt.Println("请输入表格名称:")

	var (
		printTabs []string
		wFunc     writeFunc
	)
	cORs := byte(0)
	switch i {
	case 1:
		printTabs = mergeArrays(config.Clients, config.Servers)
	case 2:
		cORs = 2
		printTabs = mergeArrays(config.Servers)
		wFunc = writeServerFunc
	case 3:
		cORs = 1
		printTabs = mergeArrays(config.Clients)
		wFunc = writeClientFunc
	}
	for _, v := range printTabs {
		fmt.Println(v)
	}
tabSelect:
	var tab string
	fmt.Scan(&tab)
	if !findInArray(tab, printTabs) {
		fmt.Printf("错误！没有找到[%v], 请注意，大小写敏感!\n", tab)
		fmt.Println("请输入表格名称:")
		goto tabSelect
	} else {
		switch cORs {
		case 0:
			if !xls2lua(tab, 1, writeClientFunc) {
				fmt.Println("客户端表格输出失败")
			}
			if !xls2lua(tab, 2, writeServerFunc) {
				fmt.Println("服务端表格输出失败")
			}
		default:
			if !xls2lua(tab, cORs, wFunc) {
				fmt.Println("表格输出失败")
			}
		}
	}
}

func readInputs() {
	var i int
start:
	fmt.Println("1.全部表格")
	fmt.Println("2.服务器表格")
	fmt.Println("3.客户端表格")
	fmt.Println("4.单个表格")
	fmt.Printf("请输入:")
	fmt.Scan(&i)
	switch i {
	case 1:
		outTables(mergeArrays(config.Clients), 1, writeClientFunc)
		outTables(mergeArrays(config.Servers), 2, writeServerFunc)
	case 2:
		outTables(mergeArrays(config.Servers), 2, writeServerFunc)
	case 3:
		outTables(mergeArrays(config.Clients), 1, writeClientFunc)
	case 4:
		outOneTable()
	default:
		goto start
	}
}

func init() {
	flag.StringVar(&excelDir, "e", "./excel", "excel directory")
	flag.StringVar(&clientOut, "c", "./client", "client output directory")
	flag.StringVar(&serverOut, "s", "./server", "server output directory")
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
