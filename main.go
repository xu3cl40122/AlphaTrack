package main

import (
	"bytes"
	"encoding/csv"
	"log"

	"github.com/go-rod/rod"
	"github.com/go-rod/rod/lib/launcher"
)

func main() {
	url := "https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D+%28%E9%AB%98%E2%86%92%E4%BD%8E%29%40%40%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D%40%40%E7%94%B1%E9%AB%98%E2%86%92%E4%BD%8E" // 目標網址
	// 設定 Chrome 下載行為，設置下載路徑
	launcher := launcher.New().
		Headless(false)

	browser := rod.New().ControlURL(launcher.MustLaunch()).MustConnect()
	defer browser.MustClose()

	// 開啟目標網頁
	page := browser.MustPage(url) // 這裡換成你的目標網址
	page.MustWaitLoad()
	page.MustWaitDOMStable()
	page.MustHandleDialog()

	// page.MustWaitStable()

	// 找到 select 元素並選擇第二個 option
	selectElement := page.MustElement(`#selRANK`)
	log.Println("selectElement:", selectElement)
	// 等待 select 元素的 child 元素 (所有 option 元素)
	options := selectElement.MustElements("option")

	selectElement.MustClick()
	for i, option := range options {
		log.Println("option:", i, option.MustText())
		if i == 1 {
			option.MustClick()
		}
	}
	selectElement.MustSelect(`option[value="1"]`)
	log.Println("選擇第二個 option")
	page.MustWaitStable()

	// 找到按鈕並點擊
	wait := browser.MustWaitDownload()
	button := page.MustElement(`input[value="匯出CSV"]`)
	button.MustClick()
	// 等待下載完成並獲取檔案路徑
	fileContent := wait()
	println("下載完成:", fileContent)
	r := bytes.NewReader(fileContent)
	csvReader := csv.NewReader(r)
	csvReader.LazyQuotes = true

	records, err := csvReader.ReadAll()
	if err != nil {
		log.Fatal(err)
	}

	// 逐行處理
	for i, record := range records {
		if i <= 3 {
			log.Println(record)
		}
	}

}
