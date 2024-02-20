package excel

import (
	"errors"
	"fmt"
)

type interSheetInfo struct {
	name2Index      map[string]int
	index2Name      map[int]string
	index2Line      map[int]int
	activeSheetName string
	activeSheetLine int
}

func newSheetInfo(sheetName string) *interSheetInfo {
	return &interSheetInfo{
		name2Index:      make(map[string]int),
		index2Name:      make(map[int]string),
		index2Line:      make(map[int]int),
		activeSheetName: sheetName,
		activeSheetLine: 0,
	}
}

func (info interSheetInfo) ExistSheet(sheetName string) (bool, int) {
	index, ok := info.name2Index[sheetName]
	if ok {
		return true, index
	}

	return false, 0
}

func (info interSheetInfo) SetSheet(sheetName string, index int) {
	info.name2Index[sheetName] = index
	info.index2Name[index] = sheetName
	info.index2Line[index] = 0
}

func (info interSheetInfo) DelSheetByName(sheetName string) {
	index, ok := info.name2Index[sheetName]
	if !ok {
		return
	}

	delete(info.name2Index, sheetName)
	delete(info.index2Name, index)

	return
}

func (info interSheetInfo) AddActivitySheetLine() {
	info.activeSheetLine++
	index, ok := info.name2Index[info.activeSheetName]
	if !ok {
		return
	}
	info.index2Line[index] = info.activeSheetLine
}

func (info interSheetInfo) GetActivitySheetLine() int {
	return info.activeSheetLine
}

func (info interSheetInfo) GetActivitySheetName() string {
	return info.activeSheetName
}

func (info interSheetInfo) SetActivitySheetName(sheetName string, index int) error {
	index, ok := info.name2Index[sheetName]
	if !ok {
		return errors.New(fmt.Sprintf("interSheetInfo.SetActivitySheetName(sheetName: %s) no exist!", sheetName))
	}

	if info.activeSheetName == sheetName {
		return nil
	}

	info.activeSheetName = sheetName
	info.activeSheetLine = info.index2Line[index]

	return nil
}
