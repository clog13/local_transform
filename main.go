package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/qichengzx/coordtransform"
	"github.com/xuri/excelize/v2"
)

const Sheet = "驻地网围栏(眉山1)"
const File = "驻地网围栏(眉山1).xlsx"

func main() {
	Init()
}

func Init() {
	f, _ := excelize.OpenFile(File)
	defer func() {
		if err := f.SaveAs(File); err != nil {
			fmt.Println(err)
		}
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	Start(f)
}

func Start(f *excelize.File) {
	cols, err := f.GetCols(Sheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	grids := cols[5][1:]
	for idx, grid := range grids {
		coords := strings.Split(grid, "；")
		res := ""
		for _, coord := range coords {
			lon, err := strconv.ParseFloat(strings.TrimSpace(strings.Split(coord, ",")[0]), 64)
			if err != nil {
				fmt.Println(err)
				return
			}
			lat, err := strconv.ParseFloat(strings.TrimSpace(strings.Split(coord, ",")[1]), 64)
			if err != nil {
				fmt.Println(err)
				return
			}
			lon_wgs84, lat_wgs84 := coordtransform.BD09toWGS84(lon, lat)
			res += strconv.FormatFloat(lon_wgs84, 'f', 6, 64) + "," + strconv.FormatFloat(lat_wgs84, 'f', 6, 64) + "; "
		}
		cell_idx, _ := excelize.CoordinatesToCellName(8, idx+2)
		f.SetCellValue(Sheet, cell_idx, res)
	}
	f.SetCellValue(Sheet, "H1", "WGS84边框坐标")
}
