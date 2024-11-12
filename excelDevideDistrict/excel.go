package exceldevidedistrict

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	F_unit              *excelize.File
	F_sum               *excelize.File
	OriginSheetNameList []string = make([]string, 0)
	SheetNameList       []string = make([]string, 0)
	CellNum             []string = []string{
		"E21",
		"E22",
		"E23",
		"E24",
		"E25",
		"E26",
		"E27",
		"E28",
		"E29",
		"E30",
		"E31",
		"E32",
		"E33",
		"E34",
		"E35",
		"E36",
		"E37",
		"E38",
		"E39",
		"E40",
		"E41",
		"E42",
		"E43",
		"E44",
		"E46",
	}
	District = map[string]string{
		"广州分行新塘支行":        "zc",
		"广州分行光大花园社区支行":    "hz",
		"广州分行佳信花园社区支行":    "hz",
		"广州分行世纪云顶社区支行":    "hz",
		"广州分行岭南新世界社区支行":   "by",
		"广州分行加和饰品城社区支行":   "lw",
		"广州分行云景花园社区支行":    "by",
		"广州分行花都龙珠社区支行":    "hd",
		"广州分行开发区支行":       "hp",
		"广州分行广钢新城支行":      "lw",
		"广州分行增城支行":        "zc",
		"广州分行白云支行":        "by",
		"广州分行海珠支行储蓄":      "hz",
		"广州分行番禺支行":        "py",
		"广州分行越秀支行":        "yx",
		"广州分行东风支行":        "yx",
		"广州分行花都支行":        "hd",
		"广州分行五羊支行":        "yx",
		"广州分行北京南路社区支行":    "yx",
		"华夏银行广州分行自贸区南沙分行": "ns",
	}
	DistrictCode = []string{
		"zc",
		"hz",
		"by",
		"lw",
		"ch",
		"hd",
		"hp",
		"py",
		"yx",
		"ns",
		"th",
	}
	CellSumMap    map[string]float64 = make(map[string]float64, 0)
	CellSumMap_zc map[string]float64 = make(map[string]float64, 0)
	CellSumMap_py map[string]float64 = make(map[string]float64, 0)
	CellSumMap_ns map[string]float64 = make(map[string]float64, 0)
	CellSumMap_hd map[string]float64 = make(map[string]float64, 0)
	CellSumMap_th map[string]float64 = make(map[string]float64, 0)
	CellSumMap_ch map[string]float64 = make(map[string]float64, 0)
	CellSumMap_by map[string]float64 = make(map[string]float64, 0)
	CellSumMap_hp map[string]float64 = make(map[string]float64, 0)
	CellSumMap_hz map[string]float64 = make(map[string]float64, 0)
	CellSumMap_lw map[string]float64 = make(map[string]float64, 0)
	CellSumMap_yx map[string]float64 = make(map[string]float64, 0)
)

func Excel() {
	// Open the file
	//f, err := excelize.OpenFile("C:/Users/kk/Desktop/moonTable/GF0103_0904_2024-10-31_01新.xlsx")
	f, err := excelize.OpenFile("C:/Users/kk/Desktop/moonTable/GF0103_0904_2023-10-31_01新.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	ff, err := excelize.OpenFile("C:/Users/kk/Desktop/moonTable/广州市各区银行业主要指标情况表.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if f != nil && ff != nil {
		F_unit = f
		F_sum = ff
	}
}

func GetSheetNameList() {
	//re := regexp.MustCompile(`_([^_]+)_`)
	OriginSheetNameList = append(OriginSheetNameList, F_unit.GetSheetList()...)
	/*
		for _, v := range OriginSheetNameList {
			SheetNameList = append(SheetNameList, re.FindStringSubmatch(v)[1])
		}

		if len(SheetNameList) > 0 {
			fmt.Println(SheetNameList)
		}
	*/
}

func GetOneCellSum(cellNum string, devide bool, district string) float64 {
	tmp := 0.0
	expectVal := 1.0
	if devide {
		re := regexp.MustCompile(`_([^_]+)_`)
		for _, v := range OriginSheetNameList {
			if district != "th" {
				if District[re.FindStringSubmatch(v)[1]] != district {
					continue
				}
			} else {
				if District[re.FindStringSubmatch(v)[1]] != "" || re.FindStringSubmatch(v)[1] == "广州分行地市级" {
					continue
				}
			}
			str, _ := F_unit.GetCellValue(v, cellNum)
			str = strings.ReplaceAll(str, ",", "")
			flo, _ := strconv.ParseFloat(str, 64)
			tmp += flo
		}
	} else {
		re := regexp.MustCompile(`广州分行地市级`)
		for _, v := range OriginSheetNameList {
			str, _ := F_unit.GetCellValue(v, cellNum)
			str = strings.ReplaceAll(str, ",", "")
			flo, _ := strconv.ParseFloat(str, 64)
			if re.MatchString(v) {
				expectVal = flo
				continue
			}
			tmp += flo
		}
		fmt.Printf(cellNum+" expectval: %f\n", expectVal)
		if tmp != expectVal {
			diff := tmp - expectVal
			absDiff := math.Abs(diff)
			if absDiff > 0.031 {
				fmt.Printf(cellNum+" absDiff: %f\n", absDiff)
				fmt.Println(cellNum + " error")
				fmt.Printf(cellNum+" sumval: %f\n", tmp)
			}
		}
	}

	return tmp
}

func GetAllCellSum() {
	for _, v := range CellNum {
		cellSumVal := GetOneCellSum(v, false, "hk")
		CellSumMap[v] = cellSumVal
	}

	fmt.Println(CellSumMap)
}

func GetAllEachDistrictCellSum() {
	for _, v := range DistrictCode {
		var districtCellSumMap map[string]float64
		switch v {
		case "zc":
			districtCellSumMap = CellSumMap_zc
		case "py":
			districtCellSumMap = CellSumMap_py
		case "ns":
			districtCellSumMap = CellSumMap_ns
		case "hd":
			districtCellSumMap = CellSumMap_hd
		case "th":
			districtCellSumMap = CellSumMap_th
		case "ch":
			districtCellSumMap = CellSumMap_ch
		case "by":
			districtCellSumMap = CellSumMap_by
		case "hp":
			districtCellSumMap = CellSumMap_hp
		case "hz":
			districtCellSumMap = CellSumMap_hz
		case "lw":
			districtCellSumMap = CellSumMap_lw
		case "yx":
			districtCellSumMap = CellSumMap_yx

		}
		for _, v1 := range CellNum {
			cellSumVal := GetOneCellSum(v1, true, v)
			districtCellSumMap[v1] = cellSumVal
		}
		fmt.Println(v, " cellSumMap: ", districtCellSumMap)
	}

}

// todayYear: true: this year, false: last year, up to the bool val, remember to change the src table name while open file
func FillFinalTable(todayYear bool) {
	// fix up final sum val
	re := regexp.MustCompile(`_([^_]+)_`)
	var stdSheetName string
	for _, v := range OriginSheetNameList {
		if re.FindStringSubmatch(v)[1] == "广州分行地市级" {
			stdSheetName = v
			break
		}
	}

	if todayYear {
		//sum
		//F_sum.SetCellValue("附件2广州市银行", "G16", CellSumMap["E21"])
		E21String, _ := F_unit.GetCellValue(stdSheetName, "E21")
		F_sum.SetCellValue("附件2广州市银行", "G16", string2float64(E21String))

		E22String, _ := F_unit.GetCellValue(stdSheetName, "E22")
		//F_sum.SetCellValue("附件2广州市银行", "I16", CellSumMap["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I16", string2float64(E22String))

		E31String, _ := F_unit.GetCellValue(stdSheetName, "E31")
		//F_sum.SetCellValue("附件2广州市银行", "K16", CellSumMap["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K16", string2float64(E31String))

		E23String, _ := F_unit.GetCellValue(stdSheetName, "E23")
		E32String, _ := F_unit.GetCellValue(stdSheetName, "E32")
		//F_sum.SetCellValue("附件2广州市银行", "M16", CellSumMap["E23"]+CellSumMap["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M16", string2float64(E23String)+string2float64(E32String))

		E24String, _ := F_unit.GetCellValue(stdSheetName, "E24")
		E33String, _ := F_unit.GetCellValue(stdSheetName, "E33")
		//F_sum.SetCellValue("附件2广州市银行", "O16", CellSumMap["E24"]+CellSumMap["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O16", string2float64(E24String)+string2float64(E33String))

		//district 各项存款
		F_sum.SetCellValue("附件2广州市银行", "G5", CellSumMap_zc["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G6", CellSumMap_py["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G7", CellSumMap_ns["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G8", CellSumMap_hd["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G9", CellSumMap_th["E21"]+(string2float64(E21String)-CellSumMap["E21"]))
		F_sum.SetCellValue("附件2广州市银行", "G10", CellSumMap_ch["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G11", CellSumMap_by["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G12", CellSumMap_hp["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G13", CellSumMap_hz["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G14", CellSumMap_lw["E21"])
		F_sum.SetCellValue("附件2广州市银行", "G15", CellSumMap_yx["E21"])

		//district 单位存款
		F_sum.SetCellValue("附件2广州市银行", "I5", CellSumMap_zc["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I6", CellSumMap_py["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I7", CellSumMap_ns["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I8", CellSumMap_hd["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I9", CellSumMap_th["E22"]+(string2float64(E22String)-CellSumMap["E22"]))
		F_sum.SetCellValue("附件2广州市银行", "I10", CellSumMap_ch["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I11", CellSumMap_by["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I12", CellSumMap_hp["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I13", CellSumMap_hz["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I14", CellSumMap_lw["E22"])
		F_sum.SetCellValue("附件2广州市银行", "I15", CellSumMap_yx["E22"])

		//district 个人存款
		F_sum.SetCellValue("附件2广州市银行", "K5", CellSumMap_zc["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K6", CellSumMap_py["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K7", CellSumMap_ns["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K8", CellSumMap_hd["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K9", CellSumMap_th["E31"]+(string2float64(E31String)-CellSumMap["E31"]))
		F_sum.SetCellValue("附件2广州市银行", "K10", CellSumMap_ch["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K11", CellSumMap_by["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K12", CellSumMap_hp["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K13", CellSumMap_hz["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K14", CellSumMap_lw["E31"])
		F_sum.SetCellValue("附件2广州市银行", "K15", CellSumMap_yx["E31"])

		//district 活期存款
		F_sum.SetCellValue("附件2广州市银行", "M5", CellSumMap_zc["E23"]+CellSumMap_zc["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M6", CellSumMap_py["E23"]+CellSumMap_py["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M7", CellSumMap_ns["E23"]+CellSumMap_ns["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M8", CellSumMap_hd["E23"]+CellSumMap_hd["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M9", CellSumMap_th["E23"]+CellSumMap_th["E32"]+(string2float64(E23String)+string2float64(E32String)-CellSumMap["E23"]-CellSumMap["E32"]))
		F_sum.SetCellValue("附件2广州市银行", "M10", CellSumMap_ch["E23"]+CellSumMap_ch["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M11", CellSumMap_by["E23"]+CellSumMap_by["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M12", CellSumMap_hp["E23"]+CellSumMap_hp["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M13", CellSumMap_hz["E23"]+CellSumMap_hz["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M14", CellSumMap_lw["E23"]+CellSumMap_lw["E32"])
		F_sum.SetCellValue("附件2广州市银行", "M15", CellSumMap_yx["E23"]+CellSumMap_yx["E32"])

		//district 定期存款
		F_sum.SetCellValue("附件2广州市银行", "O5", CellSumMap_zc["E24"]+CellSumMap_zc["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O6", CellSumMap_py["E24"]+CellSumMap_py["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O7", CellSumMap_ns["E24"]+CellSumMap_ns["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O8", CellSumMap_hd["E24"]+CellSumMap_hd["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O9", CellSumMap_th["E24"]+CellSumMap_th["E33"]+(string2float64(E24String)+string2float64(E33String)-CellSumMap["E24"]-CellSumMap["E33"]))
		F_sum.SetCellValue("附件2广州市银行", "O10", CellSumMap_ch["E24"]+CellSumMap_ch["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O11", CellSumMap_by["E24"]+CellSumMap_by["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O12", CellSumMap_hp["E24"]+CellSumMap_hp["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O13", CellSumMap_hz["E24"]+CellSumMap_hz["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O14", CellSumMap_lw["E24"]+CellSumMap_lw["E33"])
		F_sum.SetCellValue("附件2广州市银行", "O15", CellSumMap_yx["E24"]+CellSumMap_yx["E33"])
	} else {
		E21String, _ := F_unit.GetCellValue(stdSheetName, "E21")
		//F_sum.SetCellValue("附件2广州市银行", "H16", CellSumMap["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H16", string2float64(E21String))

		E22String, _ := F_unit.GetCellValue(stdSheetName, "E22")
		//F_sum.SetCellValue("附件2广州市银行", "J16", CellSumMap["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J16", string2float64(E22String))

		E31String, _ := F_unit.GetCellValue(stdSheetName, "E31")
		//F_sum.SetCellValue("附件2广州市银行", "L16", CellSumMap["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L16", string2float64(E31String))

		E23String, _ := F_unit.GetCellValue(stdSheetName, "E23")
		E32String, _ := F_unit.GetCellValue(stdSheetName, "E32")
		//F_sum.SetCellValue("附件2广州市银行", "N16", CellSumMap["E23"]+CellSumMap["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N16", string2float64(E23String)+string2float64(E32String))

		E24String, _ := F_unit.GetCellValue(stdSheetName, "E24")
		E33String, _ := F_unit.GetCellValue(stdSheetName, "E33")
		//F_sum.SetCellValue("附件2广州市银行", "P16", CellSumMap["E24"]+CellSumMap["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P16", string2float64(E24String)+string2float64(E33String))

		//district 负债总额
		F_sum.SetCellValue("附件2广州市银行", "H5", CellSumMap_zc["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H6", CellSumMap_py["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H7", CellSumMap_ns["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H8", CellSumMap_hd["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H9", CellSumMap_th["E21"]+(string2float64(E21String)-CellSumMap["E21"]))
		F_sum.SetCellValue("附件2广州市银行", "H10", CellSumMap_ch["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H11", CellSumMap_by["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H12", CellSumMap_hp["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H13", CellSumMap_hz["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H14", CellSumMap_lw["E21"])
		F_sum.SetCellValue("附件2广州市银行", "H15", CellSumMap_yx["E21"])

		//district 单位存款
		F_sum.SetCellValue("附件2广州市银行", "J5", CellSumMap_zc["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J6", CellSumMap_py["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J7", CellSumMap_ns["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J8", CellSumMap_hd["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J9", CellSumMap_th["E22"]+(string2float64(E22String)-CellSumMap["E22"]))
		F_sum.SetCellValue("附件2广州市银行", "J10", CellSumMap_ch["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J11", CellSumMap_by["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J12", CellSumMap_hp["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J13", CellSumMap_hz["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J14", CellSumMap_lw["E22"])
		F_sum.SetCellValue("附件2广州市银行", "J15", CellSumMap_yx["E22"])

		//district 个人存款
		F_sum.SetCellValue("附件2广州市银行", "L5", CellSumMap_zc["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L6", CellSumMap_py["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L7", CellSumMap_ns["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L8", CellSumMap_hd["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L9", CellSumMap_th["E31"]+(string2float64(E31String)-CellSumMap["E31"]))
		F_sum.SetCellValue("附件2广州市银行", "L10", CellSumMap_ch["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L11", CellSumMap_by["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L12", CellSumMap_hp["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L13", CellSumMap_hz["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L14", CellSumMap_lw["E31"])
		F_sum.SetCellValue("附件2广州市银行", "L15", CellSumMap_yx["E31"])

		//district 活期存款
		F_sum.SetCellValue("附件2广州市银行", "N5", CellSumMap_zc["E23"]+CellSumMap_zc["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N6", CellSumMap_py["E23"]+CellSumMap_py["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N7", CellSumMap_ns["E23"]+CellSumMap_ns["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N8", CellSumMap_hd["E23"]+CellSumMap_hd["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N9", CellSumMap_th["E23"]+CellSumMap_th["E32"]+(string2float64(E23String)+string2float64(E32String)-CellSumMap["E23"]-CellSumMap["E32"]))
		F_sum.SetCellValue("附件2广州市银行", "N10", CellSumMap_ch["E23"]+CellSumMap_ch["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N11", CellSumMap_by["E23"]+CellSumMap_by["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N12", CellSumMap_hp["E23"]+CellSumMap_hp["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N13", CellSumMap_hz["E23"]+CellSumMap_hz["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N14", CellSumMap_lw["E23"]+CellSumMap_lw["E32"])
		F_sum.SetCellValue("附件2广州市银行", "N15", CellSumMap_yx["E23"]+CellSumMap_yx["E32"])

		//district 定期存款
		F_sum.SetCellValue("附件2广州市银行", "P5", CellSumMap_zc["E24"]+CellSumMap_zc["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P6", CellSumMap_py["E24"]+CellSumMap_py["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P7", CellSumMap_ns["E24"]+CellSumMap_ns["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P8", CellSumMap_hd["E24"]+CellSumMap_hd["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P9", CellSumMap_th["E24"]+CellSumMap_th["E33"]+(string2float64(E24String)+string2float64(E33String)-CellSumMap["E24"]-CellSumMap["E33"]))
		F_sum.SetCellValue("附件2广州市银行", "P10", CellSumMap_ch["E24"]+CellSumMap_ch["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P11", CellSumMap_by["E24"]+CellSumMap_by["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P12", CellSumMap_hp["E24"]+CellSumMap_hp["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P13", CellSumMap_hz["E24"]+CellSumMap_hz["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P14", CellSumMap_lw["E24"]+CellSumMap_lw["E33"])
		F_sum.SetCellValue("附件2广州市银行", "P15", CellSumMap_yx["E24"]+CellSumMap_yx["E33"])
	}

	F_sum.Save()
}

func string2float64(str string) float64 {
	str = strings.ReplaceAll(str, ",", "")
	f, _ := strconv.ParseFloat(str, 64)
	return f
}
