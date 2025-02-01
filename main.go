package main

import (
    "bytes"
    "encoding/csv"
    "log"
    "sync"
    "github.com/go-rod/rod"
    "github.com/go-rod/rod/lib/launcher"
    "github.com/xuri/excelize/v2"
)

func main() {
    url := "https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D+%28%E9%AB%98%E2%86%92%E4%BD%8E%29%40%40%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D%40%40%E7%94%B1%E9%AB%98%E2%86%92%E4%BD%8E"
    categoryMap := map[string][]string{
        "成交額": {"交易狀況–成交資料"},
        "短漲跌": {"交易狀況–漲跌及成交統計", "短期累計漲跌統計"},
        // 其他檔名和 URL
    }

    countOptionList := []string{"1", "301", "601"}
    outputDir := "./output"
    var wg sync.WaitGroup

    f := excelize.NewFile()

    for fileName, categoryList := range categoryMap {
        wg.Add(1)
        go func(fileName, url string) {
            defer wg.Done()
            GetDataFlow(fileName, url, categoryList, countOptionList, outputDir, f)
        }(fileName, url)
    }

    wg.Wait()

    // 保存 Excel 文件
    if err := f.SaveAs(outputDir + "/output.xlsx"); err != nil {
        log.Fatal(err)
    }
}

func GetDataFlow(fileName, url string, categoryList, countOptionList []string, outputDir string, f *excelize.File) {
    browser := CreateBrowser()
    defer browser.MustClose()

    fileContentList := DownloadFileByOptions(browser, url, categoryList, countOptionList)

    err := aggregateCsvDataToExcel(fileContentList, f, fileName)
    if err != nil {
        log.Fatal(err)
    }
}

func CreateBrowser() *rod.Browser {
    uBlockPath := "./uBlock"
    launcher := launcher.New().
        Set("load-extension", uBlockPath).
        Set("extensions-on-chrome-urls").
        Headless(false)

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

func DownloadFileByOptions(browser *rod.Browser, url string, categoryList []string, countOptionList []string) [][]byte {
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
    }

    return output
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