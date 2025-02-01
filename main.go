package main

import (
    "bytes"
    "encoding/csv"
    "log"
    "os"
    "strings"
    "sync"

    "github.com/go-rod/rod"
    "github.com/go-rod/rod/lib/launcher"
)

func main() {
    categoryMap := map[string]string{
        "成交額": "https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D+%28%E9%AB%98%E2%86%92%E4%BD%8E%29%40%40%E6%88%90%E4%BA%A4%E9%87%91%E9%A1%8D%40%40%E7%94%B1%E9%AB%98%E2%86%92%E4%BD%8E",
        "當沖": "https://goodinfo.tw/tw2/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E7%8F%BE%E8%82%A1%E7%95%B6%E6%B2%96%E5%BC%B5%E6%95%B8+%28%E7%95%B6%E6%97%A5%29%40%40%E7%8F%BE%E8%82%A1%E7%95%B6%E6%B2%96%E5%BC%B5%E6%95%B8%40%40%E7%95%B6%E6%97%A5",
        // 其他檔名和 URL
    }

    optionList := []string{"1", "301", "601"}
    outputDir := "./output"
    var wg sync.WaitGroup

    for fileName, url := range categoryMap {
        wg.Add(1)
        go func(fileName, url string) {
            defer wg.Done()
            browser := CreateBrowser()
            defer browser.MustClose()

            fileContentList := DownloadFileByOptions(browser, url, optionList)
            log.Println("Downloaded", len(fileContentList), "files for URL:", url)

            outputPath := outputDir + "/" + fileName + ".csv"
            mergedContent, err := aggregateCsvData(fileContentList, outputPath)
            if err != nil {
                log.Fatal(err)
            }

            log.Println("合併後的 CSV 內容 for URL:", url)
            log.Println(string(mergedContent))
        }(fileName, url)
    }

    wg.Wait()
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

func DownloadFileByOptions(browser *rod.Browser, url string, optionList []string) [][]byte {
    page := browser.MustPage(url)
    page.MustWaitLoad()
    output := make([][]byte, 0)

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

        if i == 0 {
            allRecords = append(allRecords, records...)
        } else {
            allRecords = append(allRecords, records[1:]...)
        }
    }

    var output strings.Builder
    csvWriter := csv.NewWriter(&output)
    err := csvWriter.WriteAll(allRecords)
    if err != nil {
        return nil, err
    }
    csvWriter.Flush()

    err = os.WriteFile(outputPath, []byte(output.String()), 0644)
    if err != nil {
        return nil, err
    }

    return []byte(output.String()), nil
}