package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
)

func main() {
	excelFileName := "./demarks.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Println("no file")
	}

	var i = 0
	var itemName string
	var itemType string
	checkinItem := make(map[string]int)
	testItem := make(map[string]int)
	checkoutItem := make(map[string]int)
	for _, sheet := range xlFile.Sheets {
		switch sheet.Name {
		case "CheckinItem":
			for _, row := range sheet.Rows {
				for _, cell := range row.Cells {
					text := cell.String()
					if i == 3 {
						itemName = text
					}
					if i==6 {

						i2, err := strconv.Atoi(text)
						if err!= nil {

						}
						//添加之前首先查看是否存在
						num, ok := checkinItem[itemName]
						if ok {
							checkinItem[itemName] = i2 + num
						}else {
							checkinItem[itemName] = i2
						}

					}
					i++
				}
				i = 0
			}
			break
		case "TestItem":
			//fmt.Println("testItem building ...")
			for _, row := range sheet.Rows {
				for _, cell := range row.Cells {
					text := cell.String()
					if i == 0 {
						itemName = text
					}
					if i == 2 {
						itemType = text
					}
					if i == 3 {
						i2, err := strconv.Atoi(text)
						if err!= nil {

						}
						switch itemType {
						case "良品":
							itemName = itemName+"_良品"
							num, ok := testItem[itemName]
							if ok {
								testItem[itemName] = i2 + num
							}else {
								testItem[itemName] = i2
							}
							break
						case "故障" :
							itemName = itemName+"_废品"
							num, ok := testItem[itemName]
							if ok {
								testItem[itemName] = i2 + num
							}else {
								testItem[itemName] = i2
							}
							break
						case "划痕" :
							itemName = itemName+"_废品"
							num, ok := testItem[itemName]
							if ok {
								testItem[itemName] = i2 + num
							}else {
								testItem[itemName] = i2
							}
							break
						}
					}
					i++
				}
				i = 0
			}
			break
		case "CheckoutItem":
			for _, row := range sheet.Rows {
				for _, cell := range row.Cells {
					text := cell.String()
					if i == 1 {
						itemName = text
					}
					if i == 2 {
						i2, err := strconv.Atoi(text)
						if err!= nil {

						}
						num, ok := checkoutItem[itemName]
						if ok {
							checkoutItem[itemName] = i2 + num
						}else {
							checkoutItem[itemName] = i2
						}
					}
					i++

				}
				i = 0
			}
			break
		}
	}

	for itemName := range checkinItem{
		fmt.Println(itemName , "||", checkinItem[itemName])
	}
	fmt.Println("-------------------------------------------------------------------")
	for itemName := range testItem {
		fmt.Println(itemName, "||", testItem[itemName])
	}
	fmt.Println("-------------------------------------------------------------------")
	for itemName := range checkoutItem {
		fmt.Println(itemName, "||", checkoutItem[itemName])
	}

	//开始反写数据进入excel表格
	sheet, err := xlFile.AddSheet("Statistic")
	if err != nil {
		fmt.Printf(err.Error())
	}


	for itemName := range checkinItem {

		row := sheet.AddRow()
		row.AddCell().Value = itemName
		row.AddCell().Value = Sum(itemName, checkinItem)
		row.AddCell().Value = "良品"
		row.AddCell().Value = Passed(itemName,testItem)
		row.AddCell().Value = Checkedout(itemName,checkoutItem)
		row = sheet.AddRow()
		row.AddCell().Value = ""
		row.AddCell().Value = ""
		row.AddCell().Value = "废品"
		row.AddCell().Value = NoPassed(itemName,testItem)
	}

	err = xlFile.Save("changed.xlsx")

}
func Sum(itemName string, checkinItem map[string]int)string{
	num, ok := checkinItem [ itemName ]
	if ok {
		return strconv.Itoa(num)
	}
	return ""

}

func Passed(itemName string, testItem map[string]int)string{
	itemName = itemName+"_良品"
	num, ok := testItem [ itemName ]
	if ok {
		return strconv.Itoa(num)
	}
	return ""
}

func NoPassed(itemName string, testItem map[string]int) string{
	itemName = itemName+"_废品"
	num, ok := testItem [ itemName ]
	if ok {
		return strconv.Itoa(num)
	}
	return ""
}

func Checkedout(itemName string, checkoutItem map[string]int)string{
	num, ok := checkoutItem [ itemName ]
	if ok {
		return strconv.Itoa(num)
	}
	return ""

}