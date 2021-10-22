package main

import (
	"github.com/mazen160/go-random"
	"github.com/xuri/excelize/v2"
	"log"
	"strconv"
)

const (
	ZeinFile = "Zein.xlsx"

	PreSheet                  = "Pre"
	PreCellTanggalIndex       = 0
	PreCellJamIndex           = 1
	PreCellCabangIndex        = 2
	PreCellPromoIndex         = 3
	PreCellJumlahPesananIndex = 4
	PreCellStart              = 2

	DataTanggal = "tanggal"
	DataJam     = "jam"
	DataCabang  = "cabang"
	DataPromo   = "promo"
	DataJumlah  = "jumlah"
)

func getData() (data []map[string]string, err error) {

	// open file
	file, err := excelize.OpenFile(ZeinFile)
	if err != nil {
		return data, err
	}

	// Get all the rows
	rows, err := file.GetRows(PreSheet)
	if err != nil {
		return data, err
	}
	for rowIndex, row := range rows {
		if rowIndex < PreCellStart {
			continue
		}
		data = append(data, map[string]string{
			DataTanggal: row[PreCellTanggalIndex][:10],
			DataJam:     row[PreCellJamIndex][:2],
			DataCabang:  row[PreCellCabangIndex],
			DataPromo:   row[PreCellPromoIndex],
			DataJumlah:  row[PreCellJumlahPesananIndex],
		})
	}

	return data, err
}

func distinct(data []map[string]string, prop string) (results []string) {
	for _, val := range data {
		isExist := false
		for _, res := range results {
			if val[prop] == res {
				isExist = true
				break
			}
		}
		if !isExist {
			results = append(results, val[prop])
		}
	}

	return results
}

func transform(data []map[string]string) (newData []map[string]string) {

	// get tanggal distinct
	distinctTanggal := distinct(data, DataTanggal)
	distinctCabang := distinct(data, DataCabang)
	distinctJam := distinct(data, DataJam)

	// transform
	for _, tanggal := range distinctTanggal {
		for _, cabang := range distinctCabang {
			for _, jam := range distinctJam {
				newData = append(newData, map[string]string{
					DataTanggal: tanggal,
					DataCabang:  cabang,
					DataJam:     jam,
				})
			}
		}
	}

	return newData
}

func countJumlah(data []map[string]string, newData []map[string]string) []map[string]string {
	for idx, newVal := range newData {
		for _, val := range data {
			if (newVal[DataTanggal] == val[DataTanggal]) && (newVal[DataCabang] == val[DataCabang]) && (newVal[DataJam] == val[DataJam]) {
				jumlahBefore, _ := strconv.Atoi(newVal[DataJumlah])
				jumlahAfter, _ := strconv.Atoi(val[DataJumlah])
				jumlahTotal := jumlahBefore + jumlahAfter
				newData[idx][DataJumlah] = strconv.Itoa(jumlahTotal)
			}
		}
	}

	return newData
}

func NewFile(newData []map[string]string) (err error) {

	// new
	f := excelize.NewFile()
	sheet := "Post"

	// Create a new sheet.
	index := f.NewSheet(sheet)

	// Set value of a cell.
	for idx, newVal := range newData {
		no := idx + 2
		f.SetCellValue(sheet, "A"+strconv.Itoa(no), newVal[DataTanggal])
		f.SetCellValue(sheet, "B"+strconv.Itoa(no), newVal[DataJam])
		f.SetCellValue(sheet, "C"+strconv.Itoa(no), newVal[DataCabang])

		// 0
		if newVal[DataJumlah] == "" {
			valRandom, _ := random.IntRange(1, 10)
			f.SetCellValue(sheet, "D"+strconv.Itoa(no), strconv.Itoa(valRandom))
		} else {
			f.SetCellValue(sheet, "D"+strconv.Itoa(no), newVal[DataJumlah])
		}

		// data
		jumlah, _ := strconv.Atoi(newVal[DataJumlah])

		// promo
		promoRandom, _ := random.IntRange(1, 10)
		f.SetCellValue(sheet, "E"+strconv.Itoa(no), "Tidak")
		if promoRandom > 7 {
			promoWeightRandom, _ := random.IntRange(1, 10)
			if promoWeightRandom > 7 {
				if jumlah < 30 && jumlah > 20 {
					f.SetCellValue(sheet, "E"+strconv.Itoa(no), "Promo")
				}
			} else {
				if jumlah > 30 {
					f.SetCellValue(sheet, "E"+strconv.Itoa(no), "Promo")
				}
			}
		}


		// class
		if jumlah < 10 {
			f.SetCellValue(sheet, "F"+strconv.Itoa(no), "Sepi")
		} else if jumlah < 30 {
			f.SetCellValue(sheet, "F"+strconv.Itoa(no), "Normal")
		} else {
			f.SetCellValue(sheet, "F"+strconv.Itoa(no), "Ramai")
		}

	}

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)

	// Save spreadsheet by the given path.
	if err = f.SaveAs("Result.xlsx"); err != nil {
		return err
	}

	return err
}

func main() {

	// get data
	data, err := getData()
	if err != nil {
		log.Fatal(err.Error())
	}

	// pre processing
	newData := transform(data)

	// counting
	newData = countJumlah(data, newData)

	// save
	if err = NewFile(newData); err != nil {
		log.Fatal(err.Error())
	}
}
