package main
    
import (
    "fmt"
    "time"
    "log"
    "io/ioutil"
    "strings"
    "net/http"
    "github.com/xuri/excelize/v2"
)

func main() {


    f, error := excelize.OpenFile("file.xlsx")
    if error != nil {
        log.Fatal(error)
    }
    columnName := "B"
    sheetName := "patent"
    totalNumberOfRows := 100

    for i := 2; i < totalNumberOfRows; i++ {
        if i % 3 == 0 {
	    fmt.Println("Sleeping extra 10 seconds")
	    time.Sleep(10 * time.Second)
        }
        client := http.Client{}
        cellName := fmt.Sprintf("%s%d", columnName, i)
	fmt.Printf(cellName + ":")
        cellValue, _ := f.GetCellValue(sheetName, cellName)
        fmt.Println(cellValue)
	if len(cellValue) == 0 {
	    fmt.Println("Patent number is null..skip")
	    continue
        }

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
	fmt.Println("Do Err:", err)
        fmt.Println("Response status: sleep extra 5 sec", resp.Status)
        time.Sleep(5 * time.Second)
        continue
    }

    defer resp.Body.Close()
    b, err := ioutil.ReadAll(resp.Body)
    final := processOutput(string(b))
    if final == "error" {
	fmt.Println("Process Err:", err)
	continue
    }

    writeCell := fmt.Sprintf("%s%d", "C", i)
    err = f.SetCellValue(sheetName, writeCell, final)

    err = f.Save()
    if err != nil {
	fmt.Println("Save Err:", err)
        fmt.Println(err)
    }

    } // end for loop
}



func processOutput (resp string) string {

    s := strings.Index(resp, "applicationConfirmationNumber")
    if s == -1 {
        return "Search error" 
    }
    skip := s+len("applicationConfirmationNumber") + 3
    //fmt.Println("String found at index: ", s)
    //fmt.Println("skipping: ", skip)
    newStringPtr := resp[skip:]
    //fmt.Println("New String: ", newStringPtr)
    s1 := strings.Index(newStringPtr, "\"")
    final := newStringPtr[:s1]
    fmt.Println("Value: ", final)
    time.Sleep(4 * time.Second)
    return final
}
