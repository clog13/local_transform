package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/qichengzx/coordtransform"
	"github.com/xuri/excelize/v2"
)

const Sheet = "驻地网围栏(眉山1)"

func main() {
	fmt.Println(coordtransform.BD09toWGS84(116.404, 39.915))
	f, _ := excelize.OpenFile("source.xlsx")
	defer func() {
		if err := f.SaveAs("source.xlsx"); err != nil {
			fmt.Println(err)
		}
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
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
		println(res)
		f.SetCellValue(Sheet, cell_idx, res)
	}
	f.SetCellValue(Sheet, "H1", "WGS84边框坐标")
}
