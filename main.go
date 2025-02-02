package main

import (
	"bytes"
	"context"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log"
	"os"
	"sync"
	"time"
	"github.com/go-rod/rod"
	"github.com/go-rod/rod/lib/launcher"
	"github.com/xuri/excelize/v2"
)

type Config struct {
	URL             string              `json:"url"`
	CategoryMap     map[string][]string `json:"categoryMap"`
	Headless        bool                `json:"headless"`
	OutputPath      string              `json:"outputPath"`
	CountOptionList []string            `json:"countOptionList"`
	ConcurrentCount int            `json:"concurrentCount"`
}

func main() {
	// 讀取 JSON 設定檔
	log.Println("=== Start ===")
	startTime := time.Now()
	config := GetConfig("./config.json")
	var wg sync.WaitGroup
	var mu sync.Mutex

	f := excelize.NewFile()

	// 創建一個帶緩衝的 channel 作為信號量，限制同時運行的 goroutine 數量
	sem := make(chan struct{}, config.ConcurrentCount)

	for tabName, categoryList := range config.CategoryMap {
		wg.Add(1)
		go func(fileName, url string) {
			defer wg.Done()
			sem <- struct{}{}        // 獲取信號量
			defer func() { <-sem }() // 釋放信號量

			startTime := time.Now()
			log.Printf("Processing %s", fileName)
			err := processURL(fileName, config, categoryList, f)
			if err != nil {
				log.Printf("Error processing %s: %v", fileName, err)
				return
			}
			elapsedTime := time.Since(startTime).Round(time.Second)
			log.Printf("Finished %s took %s", fileName, elapsedTime)
		}(tabName, config.URL)
	}

	wg.Wait()

	// 獲取當前時間並格式化
	formattedTime := time.Now().Format("2006_01_02")

	// 保存 Excel 文件
	outputFilePath := fmt.Sprintf("%s/output_%s.xlsx", config.OutputPath, formattedTime)
	mu.Lock()
	if err := f.SaveAs(outputFilePath); err != nil {
		log.Fatal(err)
	}
	mu.Unlock()
	log.Println("=== Finished ===", "Time:", time.Since(startTime).Round(time.Second))
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

func processURL(fileName string, config Config, categoryList []string, f *excelize.File) error {
	browser := CreateBrowser(config.Headless)
	defer browser.MustClose()
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Minute)
	defer cancel()

	fileContentList, err := DownloadFileByOptions(ctx, browser, config.URL, categoryList, config.CountOptionList)
	if err != nil {
		return fmt.Errorf("error downloading files for %s: %w", fileName, err)
	}

	err = aggregateCsvDataToExcel(fileContentList, f, fileName)
	if err != nil {
		return fmt.Errorf("error aggregating CSV data for %s: %w", fileName, err)
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

func DownloadFileByOptions(ctx context.Context, browser *rod.Browser, url string, categoryList []string, countOptionList []string) ([][]byte, error) {
	page := browser.MustPage(url)
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

		select {
		case <-ctx.Done():
			return nil, errors.New("Timeout")
		default:

		}
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
