package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"
	"sync"
	"time"

	"github.com/go-rod/rod"
	"github.com/go-rod/rod/lib/launcher"
	"github.com/xuri/excelize/v2"
	"github.com/xu3cl40122/AlphaTrack/util" // Import the util package
)

type Config struct {
	URL             string              `json:"url"`
	CategoryMap     map[string][]string `json:"categoryMap"`
	WeekCategoryMap map[string][]string `json:"weekCategoryMap"`
	Target          string              `json:"target"`
	Headless        bool                `json:"headless"`
	OutputPath      string              `json:"outputPath"`
	CountOptionList []string            `json:"countOptionList"`
	ConcurrentCount int                 `json:"concurrentCount"`
	Timeout         int                 `json:"timeout"` // Timeout in seconds
}

func main() {
	// 讀取 JSON 設定檔
	log.Println("=== Start ===")
	startTime := time.Now()
	config := GetConfig("./config.json")
	outputFilePath := composeFileName(config)
	var wg sync.WaitGroup

	excelFile := excelize.NewFile()
	existingTabs, excelFile := checkExistedTabs(outputFilePath, excelFile)

	// 創建一個帶緩衝的 channel 作為信號量，限制同時運行的 goroutine 數量
	sem := make(chan struct{}, config.ConcurrentCount)

	categoryMap := config.CategoryMap
	if config.Target == "week" {
		categoryMap = config.WeekCategoryMap
	}
	log.Println("=== target:", config.Target, "===")

	processCategoryMap(config, categoryMap, excelFile, sem, &wg, existingTabs)

	wg.Wait()

	if err := excelFile.SaveAs(outputFilePath); err != nil {
		log.Fatal(err)
		return
	}
	log.Println("Save output file", outputFilePath)
	log.Println("=== Finished ===", "Time:", time.Since(startTime).Round(time.Second))
}

func checkExistedTabs(outputFilePath string, f *excelize.File) (map[string]bool, *excelize.File) {
	existingTabs := make(map[string]bool)

	// 檢查是否已經存在的 Excel 文件
	if _, err := os.Stat(outputFilePath); err == nil {
		f, err = excelize.OpenFile(outputFilePath)
		if err != nil {
			log.Fatal(err)
		}
		for _, sheetName := range f.GetSheetList() {
			existingTabs[sheetName] = true
		}
	}
	return existingTabs, f
}

func composeFileName(config Config) string {
	formattedTime := time.Now().Format("2006_01_02")
	outputFilePath := fmt.Sprintf("%s/%s_%s.xlsx", config.OutputPath, config.Target, formattedTime)
	return outputFilePath
}

func GetConfig(path string) Config {
	configFile, err := os.Open(path)
	if err != nil {
		log.Fatal(err)
	}
	defer configFile.Close()

	byteValue, err := io.ReadAll(configFile)
	if err != nil {
		log.Fatal(err)
	}

	var config Config
	if err := json.Unmarshal(byteValue, &config); err != nil {
		log.Fatal(err)
	}
	return config
}

func processCategoryMap(config Config, categoryMap map[string][]string, f *excelize.File, sem chan struct{}, wg *sync.WaitGroup, existingTabs map[string]bool) {
	for fileName, categoryList := range categoryMap {
		if existingTabs[fileName] {
			log.Printf("Skipping %s as it already exists", fileName)
			continue
		}
		wg.Add(1)
		go func(tabName, url string) {
			defer wg.Done()
			sem <- struct{}{}        // 獲取信號量
			defer func() { <-sem }() // 釋放信號量

			startTime := time.Now()
			log.Printf("Processing %s", tabName)
			err := util.DoWithTimeout(time.Duration(300)*time.Second, func() error {
				return processSingleTab(tabName, config, categoryList, f)
			})
			if err != nil {
				log.Printf("Error processing %s: %v", tabName, err)
				return
			}
			elapsedTime := time.Since(startTime).Round(time.Second)
			log.Printf("Finished %s took %s", tabName, elapsedTime)
		}(fileName, config.URL)
	}
}

func processSingleTab(tabName string, config Config, categoryList []string, f *excelize.File) error {
	browser := CreateBrowser(config.Headless)
	defer browser.MustClose()

	fileContentList, err := DownloadFileByOptions(browser, config.URL, categoryList, config.CountOptionList, config.Timeout)
	if err != nil {
		return fmt.Errorf("error downloading files for %s: %w", tabName, err)
	}

	err = aggregateCsvDataToExcel(fileContentList, f, tabName)
	if err != nil {
		return fmt.Errorf("error aggregating CSV data for %s: %w", tabName, err)
	}

	return nil
}

func CreateBrowser(headless bool) *rod.Browser {
	uBlockPath := "./uBlock"
	launcher := launcher.New().
		Set("load-extension", uBlockPath).
		Set("extensions-on-chrome-urls")

	if headless {
		launcher = launcher.Headless(true)
	} else {
		launcher = launcher.Headless(false)
	}
	launcher.Leakless(true)

	browser := rod.New().ControlURL(launcher.MustLaunch()).MustConnect()
	return browser
}

func getSelectorIdByIndex(i int) string {
	switch i {
	case 0:
		return "#selSHEET"
	case 1:
		return "#selSHEET2"
	}
	return ""
}
// todo: error handling
func DownloadFileByOptions(browser *rod.Browser, url string, categoryList []string, countOptionList []string, timeout int) ([][]byte, error) {
	page := browser.MustPage(url).Timeout(time.Duration(timeout) * time.Second)
	page.MustWaitLoad()
	output := make([][]byte, 0)

	for _, option := range countOptionList {
		for i, category := range categoryList {
			selectNodeOfCategory := page.MustElement(getSelectorIdByIndex(i))
			selectNodeOfCategory.MustSelect(category)
			page.MustWaitStable()
		}
		selectNodeOfCount := page.MustElement(`#selRANK`)
		selectNodeOfCount.MustSelect(option)
		page.MustWaitStable()

		wait := browser.MustWaitDownload()
		button := page.MustElement(`input[value="匯出CSV"]`)
		button.MustClick()

		fileContent := wait()
		output = append(output, fileContent)
	}

	return output, nil
}

func aggregateCsvDataToExcel(fileContentList [][]byte, f *excelize.File, sheetName string) error {
	var allRecords [][]string

	for i, fileContent := range fileContentList {
		r := bytes.NewReader(fileContent)
		csvReader := csv.NewReader(r)
		csvReader.LazyQuotes = true

		records, err := csvReader.ReadAll()
		if err != nil {
			return err
		}

		if i == 0 {
			allRecords = append(allRecords, records...)
		} else {
			allRecords = append(allRecords, records[1:]...)
		}
	}

	// 創建新的工作表
	index, _ := f.NewSheet(sheetName)

	// 將合併後的資料寫入工作表
	for i, record := range allRecords {
		for j, cell := range record {
			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			f.SetCellValue(sheetName, cellName, cell)
		}
	}

	// 設置活動工作表
	f.SetActiveSheet(index)

	return nil
}
