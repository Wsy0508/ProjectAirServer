package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

type Config struct {
	TagetPath string `json:"TagetPath"`
	OutPath   string `json:"OutPath"`
}

type DataInfo struct {
	names     []string
	types     []string
	onlytypes []int
}

func (s *DataInfo) GetName(Index int) (*string, error) {
	if Index >= len(s.names) || Index < 0 {
		return nil, errors.New("GetName 索引越界")
	}

	return &s.names[Index], nil
}

func (s *DataInfo) GetType(Index int) (*string, error) {
	if Index >= len(s.types) || Index < 0 {
		return nil, errors.New("GetType 索引越界")
	}

	return &s.types[Index], nil
}

func (s *DataInfo) GetOnlyType(Index int) int {
	if Index >= len(s.types) || Index < 0 {
		return 0
	}

	return s.onlytypes[Index]
}

func main() {

	defer fmt.Scanln()

	if len(os.Args) <= 0 {
		fmt.Println("参数传递错误")
		return
	}

	param := os.Args[0]
	onlytype := 1
	if param == "server" {
		onlytype = 2
	} else if param == "client" {
		onlytype = 1
	}
	// 记录开始时间
	startTime := time.Now()

	configFile, err := os.Open("config.json")
	if err != nil {
		fmt.Println("Error opening config file:", err)
		return
	}
	defer configFile.Close()

	var conf Config
	decoder := json.NewDecoder(configFile)
	err = decoder.Decode(&conf)
	if err != nil {
		fmt.Println("Error decoding config file:", err)
		return
	}

	//removeAllContents("Out", false)
	Convert_Dir(conf.TagetPath, conf.OutPath, onlytype)

	endTime := time.Now()
	elapsedTime := endTime.Sub(startTime)
	fmt.Printf("导出成功 耗时：%f", elapsedTime.Seconds())

}

// 转换文件夹内所有文件
func Convert_Dir(path string, outpath string, onlytype int) error {
	// 打开目录
	dir, err := os.Open(path)
	if err != nil {
		return err
	}
	defer dir.Close()

	// 读取目录内容
	fileListInfos, err := dir.Readdir(-1)
	if err != nil {
		return err
	}

	// 遍历目录内容
	for _, fileInfo := range fileListInfos {
		if fileInfo.Name() == ".svn" {
			continue
		}

		SonPath := filepath.Join(path, fileInfo.Name())
		if fileInfo.IsDir() {
			SonOutPath := filepath.Join(outpath, fileInfo.Name())
			os.Mkdir(SonOutPath, 0755)

			// 如果是子目录，递归删除子目录及其内容
			err := Convert_Dir(SonPath, SonOutPath, onlytype)
			if err != nil {
				return err
			}
		} else {
			// 如果是文件，转换文件
			Convert_File(SonPath, fileInfo.Name(), outpath, onlytype)
		}
	}
	return nil
}

// 转换文件
func Convert_File(excelFilePath string, excelFileName string, outPath string, onlytype int) {
	// 打开Excel文件
	if strings.Contains(excelFileName, "~$") {
		return
	}

	parts := strings.Split(excelFileName, ".")
	if len(parts) != 2 {
		fmt.Printf("转换文件Failed: %s 文件名不匹配\n", excelFileName)
		return
	}

	if parts[1] != "xlsx" {
		fmt.Printf("转换文件Failed: %s 格式不匹配\n", excelFileName)
		return
	}

	xlFile, err := xlsx.OpenFile(excelFilePath)
	if err != nil {
		log.Printf("无法打开Excel文件:%s- %s\n", excelFilePath, err)
		return
	}
	sheet := xlFile.Sheets[0]
	if sheet.MaxRow < 2 {
		log.Printf("数据表格式错误，行数 小于 2: %s\n", excelFileName)
		return
	}

	nameRow := sheet.Rows[0]
	typeRow := sheet.Rows[1]
	onlytypeRow := sheet.Rows[2]
	tempinfo := DataInfo{}
	for i := 0; i < len(nameRow.Cells); i++ {
		tempinfo.names = append(tempinfo.names, nameRow.Cells[i].String())
	}
	for i := 0; i < len(typeRow.Cells); i++ {
		tempinfo.types = append(tempinfo.types, typeRow.Cells[i].String())
	}
	for i := 0; i < len(onlytypeRow.Cells); i++ {
		intValue, err := strconv.Atoi(onlytypeRow.Cells[i].String())
		if err != nil {
			tempinfo.onlytypes = append(tempinfo.onlytypes, 0)
		} else {
			tempinfo.onlytypes = append(tempinfo.onlytypes, intValue)
		}

	}

	// 尝试打开文件
	// 选择要读取的工作表
	luaName := parts[0] + ".lua"
	luaPath := filepath.Join(outPath, luaName)
	file, err := os.OpenFile(luaPath, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0644)
	if err != nil {
		log.Fatalf("OpenAndCreateFile Failed: %s  %s\n", luaPath, err)
		return
	}
	defer file.Close()

	file.WriteString("return\n")
	file.WriteString("{\n")

	for i := 4; i < len(sheet.Rows); i++ {
		dataRow := sheet.Rows[i]
		if len(dataRow.Cells) <= 0 {
			continue
		}

		if dataRow.Cells[0].String() == "" {
			continue
		}

		KeyOnlyType := tempinfo.GetOnlyType(0)
		if KeyOnlyType > 0 && KeyOnlyType != onlytype {
			continue
		}

		file.WriteString("[")
		file.WriteString(dataRow.Cells[0].String())
		file.WriteString("] ")
		file.WriteString("= {")
		for index := 0; index < len(dataRow.Cells); index++ {
			if index >= len(tempinfo.names) {
				continue
			}

			tempname, nameerr := tempinfo.GetName(index)
			if *tempname == "" {
				continue
			}

			ColOnlyType := tempinfo.GetOnlyType(index)
			if ColOnlyType == 3 { //策划备注
				continue
			}
			if ColOnlyType > 0 && ColOnlyType != onlytype {
				continue
			}

			temptype, typeerr := tempinfo.GetType(index)
			if nameerr != nil {
				log.Fatalf("GetName Failed: %s  %s 行 %d  列%d\n", luaPath, nameerr, i, index)
				return
			}

			if typeerr != nil {
				log.Fatalf("GetType Failed: %s  %s--行 %d 列%d\n", luaPath, typeerr, i, index)
				return
			}
			file.WriteString(*tempname)
			file.WriteString("=")
			if *temptype == "string" || *temptype == "String" {
				file.WriteString("\"")
			} else if *temptype == "array" || *temptype == "Array" {
				file.WriteString("{")
			}
			file.WriteString(dataRow.Cells[index].String())
			if *temptype == "string" || *temptype == "String" {
				file.WriteString("\"")
			} else if *temptype == "array" || *temptype == "Array" {
				file.WriteString("}")
			}
			file.WriteString(",")
		}
		file.WriteString("},\n")
	}
	file.WriteString("}\n")
}

// 递归删除目录及其所有内容
func removeAllContents(path string, bRemovePath bool) error {
	// 打开目录
	dir, err := os.Open(path)
	if err != nil {
		return err
	}
	defer dir.Close()

	// 读取目录内容
	fileInfos, err := dir.Readdir(-1)
	if err != nil {
		return err
	}

	// 遍历目录内容
	for _, fileInfo := range fileInfos {
		if fileInfo.Name() == ".svn" {
			continue
		}

		filePath := filepath.Join(path, fileInfo.Name())

		if fileInfo.IsDir() {
			// 如果是子目录，递归删除子目录及其内容
			removeAllContents(filePath, true)
		} else {
			if !strings.Contains(fileInfo.Name(), ".lua") {
				return nil
			}
			// 如果是文件，删除文件
			os.Remove(filePath)
		}
	}

	// 删除目录
	if bRemovePath {
		dir.Close()
		err = os.RemoveAll(path)
		if err != nil {
			fmt.Printf("RemoveAll Failed: %s  ,error: %s\n", bRemovePath, err)
			return err
		}
	}
	return nil
}
