package main

import (
	"bytes"
	"encoding/csv"
	"log"
	"os"
	"strings"

	"github.com/go-rod/rod"
	"github.com/go-rod/rod/lib/launcher"
)

func main() {
	url := "https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D+%28%E9%AB%98%E2%86%92%E4%BD%8E%29%40%40%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D%40%40%E7%94%B1%E9%AB%98%E2%86%92%E4%BD%8E" // 目標網址
	browser := CreateBrowser()
	defer browser.MustClose()

	optionList := []string{"1", "301", "601"}
	fileContentList := DownloadFileByOptions(browser, url, optionList)
	log.Println("Downloaded", len(fileContentList), "files")
	aggregateCsvData(fileContentList, "./output/output.csv")

}

func CreateBrowser() *rod.Browser {
	uBlockPath := "./uBlock"
	// 設定 Chrome 下載行為，設置下載路徑
	launcher := launcher.New().
		// 加載擴充功能
		Set("load-extension", uBlockPath).
		// 允許安裝擴充功能
		Set("extensions-on-chrome-urls").
		Headless(false)

	browser := rod.New().ControlURL(launcher.MustLaunch()).MustConnect()
	return browser
}

func DownloadFileByOptions(browser *rod.Browser, url string, optionList []string) [][]byte {
	page := browser.MustPage(url)
	page.MustWaitLoad()
	output := make([][]byte, 0)

	// 對每個選項進行下載
	for _, option := range optionList {
		selectElement := page.MustElement(`#selRANK`)
		selectElement.MustSelect(option)
		page.MustWaitStable()

		wait := browser.MustWaitDownload()
		button := page.MustElement(`input[value="匯出CSV"]`)
		button.MustClick()

		fileContent := wait()
		output = append(output, fileContent)
	}

	return output
}

func aggregateCsvData(fileContentList [][]byte, outputPath string) ([]byte, error) {
	var allRecords [][]string

	for i, fileContent := range fileContentList {
		r := bytes.NewReader(fileContent)
		csvReader := csv.NewReader(r)
		csvReader.LazyQuotes = true

		records, err := csvReader.ReadAll()
		if err != nil {
			return nil, err
		}

		// 保留第一份的表頭
		if i == 0 {
			allRecords = append(allRecords, records...)
		} else {
			allRecords = append(allRecords, records[1:]...) // 跳過表頭
		}
	}

	// 將合併後的資料寫回 CSV 格式
	var output strings.Builder
	csvWriter := csv.NewWriter(&output)
	err := csvWriter.WriteAll(allRecords)
	if err != nil {
		return nil, err
	}
	csvWriter.Flush()

	// 將合併後的內容寫入指定路徑的檔案
	err = os.WriteFile(outputPath, []byte(output.String()), 0644)
	if err != nil {
		return nil, err
	}

	return []byte(output.String()), nil
}
