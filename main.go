package excel_to_csv

import (
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"path"
	"runtime"
	"strconv"
	"strings"
	"sync"
)

var source string
var wg = sync.WaitGroup{}
var columnIndex = make(map[string]string, 0)

func init() {
	runtime.GOMAXPROCS(2)
	flag.StringVar(&source, "source", "", "需转换文件")
	cols := flag.String("column", "", "保留的列：0,1,2")
	flag.Parse()
	if *cols != ""{
		indexs := strings.Split(*cols, ",")
		for _, index := range indexs {
			columnIndex[index] = index
		}
	}
}

func main() {
	sourceFile, err := xlsx.OpenFile(source)
	if err != nil {
		panic(err)
	}

	for _, sheet := range sourceFile.Sheets {
		wg.Add(1)
		go func(sheet *xlsx.Sheet) {
			trans(sheet)
		}(sheet)
	}

	wg.Wait()
	fmt.Println("Done")
}

func trans(sheet *xlsx.Sheet) {
	filename := path.Dir(source) + "/" + sheet.Name + ".csv"
	fmt.Println(filename)
	csvFile, _ := os.Create(filename)

	defer func() {
		csvFile.Close()
		wg.Done()
	}()

	for i := 0; i < sheet.MaxRow; i++ {
		cols := len(sheet.Row(i).Cells)
		tmpData := []string{}
		for j := 0; j < cols; j ++ {
			if _, ok := columnIndex[strconv.Itoa(j)]; len(columnIndex) >= 1 && !ok {
				continue
			}
			tmpData = append(tmpData, handleValue(sheet.Cell(i, j).Value))
		}
		_, err := csvFile.Write([]byte(strings.Join(tmpData, ",") + "\n"))
		if err != nil {
			panic("csv file write error " + err.Error())
		}
	}
}

func handleValue(value string) string {
	value = strings.Trim(value, " \t\n")
	value = strings.ReplaceAll(value, "\n", "")

	if strings.Contains(value, "\"") {
		value = strings.ReplaceAll(value, "\"", "\"\"")
	}
	if strings.Contains(value, ",") {
		value = "\"" + value + "\""
	}

	return value
}
