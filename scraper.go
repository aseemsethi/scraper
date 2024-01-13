package main
    
import (
    "fmt"
    "log"
    "io/ioutil"
    "strings"
    "net/http"
    "github.com/xuri/excelize/v2"
)

func main() {

    client := http.Client{}

    f, error := excelize.OpenFile("file.xlsx")
    if error != nil {
        log.Fatal(error)
    }
    columnName := "B"
    sheetName := "patent"
    totalNumberOfRows := 5

    for i := 2; i < totalNumberOfRows; i++ {
        cellName := fmt.Sprintf("%s%d", columnName, i)
	fmt.Printf(cellName + ":")
        cellValue, _ := f.GetCellValue(sheetName, cellName)
        fmt.Println(cellValue)

    var url = "https://patentcenter.uspto.gov/retrieval/public/v2/application/data?patentNumber=" + cellValue
    fmt.Println("URL: ", url)
    req , err := http.NewRequest("GET", url, nil)
    if err != nil {
        //Handle Error
    }

    req.Header = http.Header{
        "Host": {"patentcenter.uspto.gov"},
        "Content-Type": {"application/json"},
	"Sec-Ch-Ua-Platform": { "macOS" },
    }

    req.Header.Set("pragma", "no-cache")
    req.Header.Set("cache-control", "no-cache")
    req.Header.Set("sec-ch-ua", `"Google Chrome";v="89", "Chromium";v="89", ";Not A Brand";v="99"`)
    req.Header.Set("sec-ch-ua-mobile", "?0")
    req.Header.Set("upgrade-insecure-requests", "1")
    req.Header.Set("user-agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36")
    req.Header.Set("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9")
    req.Header.Set("dnt", "1")
    req.Header.Set("sec-fetch-site", "none")
    req.Header.Set("sec-fetch-mode", "navigate")
    req.Header.Set("sec-fetch-user", "?1")
    req.Header.Set("sec-fetch-dest", "document")
    req.Header.Set("accept-language", "en-GB,en;q=0.9")

    resp , err := client.Do(req)
    if err != nil {
	fmt.Println("Err:", err)
    }

    defer resp.Body.Close()
    fmt.Println("Response status:", resp.Status)
    b, err := ioutil.ReadAll(resp.Body)
    processOutput(string(b))
    } // end for loop
}



func processOutput (resp string) {

    s := strings.Index(resp, "applicationConfirmationNumber")
    if s == -1 {
        return
    }
    skip := s+len("applicationConfirmationNumber") + 3
    //fmt.Println("String found at index: ", s)
    //fmt.Println("skipping: ", skip)
    newStringPtr := resp[skip:]
    //fmt.Println("New String: ", newStringPtr)
    s1 := strings.Index(newStringPtr, "\"")
    final := newStringPtr[:s1]
    fmt.Println("Value: ", final)
}
