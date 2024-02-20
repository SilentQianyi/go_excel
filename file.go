package excel

import (
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
)

type Excel struct {
	file      *excelize.File
	sheetInfo *interSheetInfo
}

type Header struct {
	Key  string
	Name string
}

type Line struct {
	Values map[string]interface{}
}

func NewExcelFile(sheetName string) (*Excel, error) {
	file := excelize.NewFile()

	sheetInfo := newSheetInfo(sheetName)
	index, err := file.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}
	sheetInfo.SetSheet(sheetName, index)

	file.SetActiveSheet(index)

	return &Excel{
		file:      file,
		sheetInfo: sheetInfo,
	}, nil
}

func (excel Excel) SetActiveSheet(index int, sheetName string) {
	excel.file.SetActiveSheet(index)
	excel.sheetInfo.SetActivitySheetName(sheetName, index)
}

func (excel Excel) WriteTop(headers []*Header) error {
	column := 0
	for _, header := range headers {
		line := fmt.Sprintf("%s1", getWord(column))
		err := excel.file.SetCellValue(excel.sheetInfo.activeSheetName, line, header.Name)
		if err != nil {
			return err
		}
		column++
	}

	excel.sheetInfo.AddActivitySheetLine()

	return nil
}

func (excel Excel) WriteData(headers []*Header, data []*Line) error {
	for _, val := range data {
		curLine := excel.sheetInfo.GetActivitySheetLine()

		column := 0
		for _, header := range headers {
			line := fmt.Sprintf("%s%v", getWord(column), curLine)

			err := excel.file.SetCellValue(excel.sheetInfo.activeSheetName, line, val.Values[header.Key])
			if err != nil {
				return err
			}
			column++
		}
		excel.sheetInfo.AddActivitySheetLine()
	}

	return nil
}

func (excel Excel) Save(path string, name string) error {
	filePath := path + "/" + name
	return excel.file.SaveAs(filePath)
}

func (excel Excel) ToBuffer() (*bytes.Buffer, error) {
	return excel.file.WriteToBuffer()
}

func getWord(column int) string {
	if column == 0 {
		return "A"
	}

	word := ""
	for column > 0 {
		remainder := column % 26
		word += fmt.Sprintf("%c", 'A'+remainder)
		column = (column - remainder) / 26
	}

	return word
}
