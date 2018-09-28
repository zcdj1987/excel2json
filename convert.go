package main

import (
	// "bufio"
	"encoding/json"
	// "flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	// "os/exec"
	"errors"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

var (
	configListPath string
	configXlsxPath string
	configJsonPath string
)

type setting struct {
	AllIn       bool   `json:"allIn"`       //(true,false)all sheets packaged into ONE jsonfile
	ConfigName  string `json:"configName"`  //if allIn is true,it will be the jsonfile Name
	IsFileSplit bool   `json:"isFileSplit"` // if true the system will read the cloned files(like play_1.xls,play_2.xls... will all read like play.xls)
	SplitKey    string `json:"splitKey"`    // Defining cloned files headname and subname
	List        []struct {
		XlBook  string            `json:"xlBook"`  //excel file name
		XlSheet string            `json:"xlSheet"` //sheet name,should be in xlbook
		Json    string            `json:"json"`    //json file name
		MaxRows int               `json:"maxRows"` //max sheet rows
		Keys    map[string]string `json:"keys"`    // it defined only convert keywords and it's type(if in sheet there have different type,It will cover the original)
	} `json:"list"`
}

var (
	configSetting *setting
	configMap     map[string]([]map[string]interface{})
	startTime     time.Time
	finishTime    time.Time
)

func main() {
	Start("", "", "")
}

// it start with the setting config path, want read excel files path and output json file path
// if there have someting wrong with the path,it will found the files in the runtime app path
func Start(listPath string, excelPath string, jsonPath string) {
	appPath := getAppPath()

	if pathExists(listPath) {
		configListPath = listPath
	} else {
		configListPath = appPath
	}
	if pathExists(excelPath) {
		configXlsxPath = excelPath
	} else {
		configXlsxPath = appPath
	}
	if pathExists(jsonPath) {
		configJsonPath = jsonPath
	} else {
		configJsonPath = appPath + "/ConfigJson"
	}

	configSetting = &setting{}
	err := configSetting.ini(configListPath)
	if err != nil {
		fmt.Println("setting file is wrong:", err)
		return
	}
	configMap = make(map[string]([]map[string]interface{}))

	startTime = time.Now()
	convert(configXlsxPath)
	CreatJsonFiles(configJsonPath)
	finishTime = time.Now()
	fmt.Println("cost time:", finishTime.Sub(startTime))
}

// main func to convert the excels to map
func convert(excelPath string) {
	var listFiles []string
	var listNames []string
	var firstName, lastName string

	listFiles, listNames = findFiles(excelPath)

	for _, sl := range configSetting.List {
		for i, ln := range listNames {
			if configSetting.IsFileSplit {
				firstName = strings.Split(strings.Split(ln, ".")[0], configSetting.SplitKey)[0]
			} else {
				firstName = strings.Split(ln, ".")[0]
			}
			lastName = strings.Split(ln, ".")[len(strings.Split(ln, "."))-1]
			if sl.XlBook == firstName && (lastName == "xlsx" || lastName == "xls" || lastName == "xlsm") {
				xlsxData, err := readXlsx(listFiles[i], sl.XlSheet, sl.MaxRows, sl.Keys)
				if err != nil {
					fmt.Println("[WARNING]excel data is wrong:", err)
					continue
				}
				if _, ok := configMap[sl.Json]; ok {
					configMap[sl.Json] = append(configMap[sl.Json], xlsxData...)
				} else {
					configMap[sl.Json] = xlsxData
				}
			}
		}
	}
}

// output the json files
func CreatJsonFiles(path string) {
	if !pathExists(path) {
		err := os.Mkdir(path, os.ModePerm)
		if err != nil {
			fmt.Println("[ERROR]json path is error.", err)
			return
		}
	}

	if configSetting.AllIn {
		writeJsonFile(path+"/"+configSetting.ConfigName+".json", configMap)
	} else {
		for k, v := range configMap {
			writeJsonFile(path+"/"+k+".json", v)
		}
	}
}

func writeJsonFile(path string, data interface{}) {
	jsonFile, err := os.OpenFile(path, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		fmt.Println("[ERROR] json file create is error:", err)
		return
	}
	jsonByte, err := json.Marshal(data)
	if err != nil {
		fmt.Println("[ERROR] json data is error:", err)
		return
	}
	_, err = jsonFile.Write(jsonByte)
	if err != nil {
		fmt.Println("[ERROR] write json is error:", err)
		return
	}
	err = jsonFile.Close()
	if err != nil {
		fmt.Println("[ERROR] write json close is error:", err)
		return
	}
	fmt.Printf("%s is complete! \n", path)
}

func readXlsx(path string, shName string, maxRows int, keys map[string]string) (data []map[string]interface{}, err error) {
	var key, typ, st string
	var mr, mc int
	var val interface{}
	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		return
	}
	sheet := xlFile.Sheet[shName]
	if sheet == nil {
		err = errors.New(path + "[" + shName + "]" + "can't find")
		return
	}

	if maxRows > 0 {
		mr = 3 + maxRows
	} else {
		mr = sheet.MaxRow
	}
	mc = sheet.MaxCol
	for i := 3; i < mr; i++ {
		rd := make(map[string]interface{})
		for j := 0; j < mc; j++ {
			key = sheet.Cell(1, j).String()
			typ = sheet.Cell(2, j).String()
			if keys != nil {
				if _, ok := keys[key]; !ok {
					continue
				}
				typ = keys[key]
			}
			st = sheet.Cell(i, j).Value
			switch typ {
			case "string":
				val = st
			case "int", "int64":
				if st == "" {
					val = 0
					break
				}
				n, err := sheet.Cell(i, j).Int64()
				if err != nil {
					val = 0
				}
				val = n
			case "double":
				if st == "" {
					val = 0
					break
				}
				fn, err := sheet.Cell(i, j).Float()
				if err != nil {
					val = 0
				}
				val = fn
			case "json":
				if sheet.Cell(i, j).String() != "" {
					jv := make(map[string]interface{})
					err = json.Unmarshal([]byte(st), &jv)
					val = jv
				} else {
					val = nil
				}
			case "arr_string":
				val = strings.Split(st, "#")
			case "arr_int":
				at := strings.Split(st, "#")
				ai := make([]int, 0)
				for _, v := range at {
					n, err := strconv.Atoi(v)
					if err != nil {
						break
					}
					ai = append(ai, n)
				}
				val = ai
			case "bool":
				if st == "" {
					val = false
					break
				}
				bo, err := sheet.Cell(i, j).Int()
				if err != nil || bo > 1 || bo < 0 {
					val = false
				} else {
					val = true
				}
			default:
			}
			if err != nil {
				fmt.Println("err:", err)
				fmt.Printf("[ERROR] xlsx val error,row:%d,col:%d,path:%s,key:%s", i, j, path, key)
				return
			}
			if typ != "" {
				rd[key] = val
			}
		}
		data = append(data, rd)
	}
	return
}

func (self *setting) ini(path string) (err error) {
	file, err := os.Open(path + "/" + "setting.json")
	if err != nil {
		fmt.Println("[ERROR] setting is error.", err)
		return
	}
	data, err := ioutil.ReadAll(file)
	if err != nil {
		fmt.Println("[ERROR] setting bytes is error.", err)
		return
	}
	err = json.Unmarshal(data, self)
	if err != nil {
		fmt.Println("[ERROR] setting init json is error.", err)
	}

	checkStr := self.checkSelf()

	if checkStr != "" {
		err = errors.New("[ERROR] setting check false,have the same json file name:" + checkStr)
	}

	return
}

func (self *setting) checkSelf() string {
	// if the setting config did not defining the configname , it will give a Initialization name
	if self.ConfigName == "" {
		self.ConfigName = "JsonConfig"
	}

	// if you want to read clone file and did not defining the splitKey,it will give a Initialization key
	if self.IsFileSplit && self.SplitKey == "" {
		self.ConfigName = "_"
	}

	cm := make(map[string]string)
	for _, v := range self.List {
		if _, ok := cm[v.Json]; ok {
			return v.Json
		}
	}
	return ""
}

func (self *setting) getJsonName(excelName string) string {
	for _, v := range self.List {
		if v.XlBook == excelName {
			return v.Json
		}
	}
	return ""
}

func pathExists(path string) bool {
	_, err := os.Stat(path)
	if err != nil {
		return false
	}
	return true
}

func findFiles(path string) (listFiles []string, listNames []string) {
	files, _ := ioutil.ReadDir(path)
	for _, file := range files {
		if file.IsDir() {
			listFiles, listNames = findFiles(path + "/" + file.Name())
		} else {
			listFiles = append(listFiles, path+"/"+file.Name())
			listNames = append(listNames, file.Name())
		}
	}
	return
}

func getAppPath() string {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println("[ERROR] appPath is error.", err)
	}
	return strings.Replace(dir, "\\", "/", -1)
}
