package main

import (
	"ExcelMerge/theme"
	"flag"
	"fmt"
	"image/color"
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/driver/desktop"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
	"github.com/xuri/excelize/v2"
)

type Atabcontent struct {
	sheetname               string
	srccontent              [][]string
	dstcontent              [][]string
	lcsindex_src            map[int]int
	lcsindex_dst            map[int]int
	col_length_per_src      []int
	col_length_per_dst      []int
	current_row_diff_cursor int
	current_row_diff        []int
	initial_rows            int
	initial_cols            int
}

var sheetcontentlist []*Atabcontent
var srcidrow, srcidcol int = -1, -1
var dstidrow, dstidcol int = -1, -1
var presrcidrow, predstidrow int = -1, -1
var diff_label = widget.NewLabel("差异数：")

var curindex int
var cursrctbl, curdsttbl *widget.Table
var ctrlS = desktop.CustomShortcut{KeyName: fyne.KeyS, Modifier: desktop.ControlModifier}
var ctrlR = desktop.CustomShortcut{KeyName: fyne.KeyR, Modifier: desktop.ControlModifier} //insert a row above selected row
var ctrlL = desktop.CustomShortcut{KeyName: fyne.KeyL, Modifier: desktop.ControlModifier} //insert a colume befor selected col
var altR = desktop.CustomShortcut{KeyName: fyne.KeyR, Modifier: desktop.AltModifier}      //delete the selected row
var altL = desktop.CustomShortcut{KeyName: fyne.KeyL, Modifier: desktop.AltModifier}      //delete the selected col
var ctrlI = desktop.CustomShortcut{KeyName: fyne.KeyI, Modifier: desktop.ControlModifier} //accept remote row
var ctrlK = desktop.CustomShortcut{KeyName: fyne.KeyK, Modifier: desktop.ControlModifier} //accept remote col
var ctrlU = desktop.CustomShortcut{KeyName: fyne.KeyU, Modifier: desktop.ControlModifier} //insert remote row
var altQ = desktop.CustomShortcut{KeyName: fyne.KeyQ, Modifier: desktop.AltModifier}      //quit

func NewAtabcontent(sn string, s [][]string, ls map[int]int, d [][]string, ld map[int]int, lengthpercolsrc []int, lengthpercoldst []int,
	current_row_diff_cursor int, current_row_diff []int, initial_rows int, initial_cols int) *Atabcontent {
	return &Atabcontent{
		sheetname:               sn,
		srccontent:              s,
		lcsindex_src:            ls,
		dstcontent:              d,
		lcsindex_dst:            ld,
		col_length_per_src:      lengthpercolsrc,
		col_length_per_dst:      lengthpercoldst,
		current_row_diff_cursor: current_row_diff_cursor,
		current_row_diff:        current_row_diff,
		initial_rows:            initial_rows,
		initial_cols:            initial_cols,
	}
}

func isSame_row(s1 []string, s2 []string) bool {
	len1 := len(s1)
	len2 := len(s2)
	if len1 != len2 {
		return false
	}
	for i := 0; i < len1; i++ {
		if s1[i] != s2[i] {
			return false
		}
	}
	return true
}

func find_diff() {
	sheetcontentlist[curindex].current_row_diff = sheetcontentlist[curindex].current_row_diff[0:0]
	sheetcontentlist[curindex].current_row_diff_cursor = 0
	if srcidrow > 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) && dstidrow > 0 && dstidrow < len(sheetcontentlist[curindex].dstcontent) {
		len1 := len(sheetcontentlist[curindex].srccontent[srcidrow])
		len2 := len(sheetcontentlist[curindex].dstcontent[dstidrow])
		minlen := min(len1, len2)
		maxlen := max(len1, len2)
		for i := 0; i < minlen; i++ {
			if sheetcontentlist[curindex].srccontent[srcidrow][i] != sheetcontentlist[curindex].dstcontent[dstidrow][i] {
				sheetcontentlist[curindex].current_row_diff = append(sheetcontentlist[curindex].current_row_diff, i)
			}
		}
		for i := minlen; i < maxlen; i++ {
			sheetcontentlist[curindex].current_row_diff = append(sheetcontentlist[curindex].current_row_diff, i)
		}
		diff_label.SetText("差异数：" + strconv.Itoa(len(sheetcontentlist[curindex].current_row_diff)))
		Refresh()
	}
}

func hasElement(i int, vector []int) bool {
	if len(vector) == 0 {
		return false
	}
	for _, element := range vector {
		if element == i {
			return true
		}
	}
	return false
}

func find_next_diff() {
	if len(sheetcontentlist[curindex].current_row_diff) != 0 {
		var srcid widget.TableCellID
		srcid.Row = srcidrow
		srcid.Col = sheetcontentlist[curindex].current_row_diff[sheetcontentlist[curindex].current_row_diff_cursor]
		var dstid widget.TableCellID
		dstid.Row = dstidrow
		dstid.Col = sheetcontentlist[curindex].current_row_diff[sheetcontentlist[curindex].current_row_diff_cursor]
		cursrctbl.Select(srcid)
		curdsttbl.Select(dstid)
		sheetcontentlist[curindex].current_row_diff_cursor = (sheetcontentlist[curindex].current_row_diff_cursor + 1) % len(sheetcontentlist[curindex].current_row_diff)
	}
}

func longestCommonSubsequence(textsrc [][]string, textdst [][]string) (map[int]int, map[int]int) {
	l1 := len(textsrc)
	l2 := len(textdst)
	// dp[i][j]标识text1 到达text1[i], text2到达text2[j]的情况下，能拿到的最长子序列的长度
	// 此处length + 1为了便于后面的计算
	dp := make([][]int, l1+1)
	for i := 0; i <= l1; i++ {
		dp[i] = make([]int, l2+1)
	}
	for i := 1; i <= l1; i++ {
		for j := 1; j <= l2; j++ {
			if isSame_row(textsrc[i-1], textdst[j-1]) {
				dp[i][j] = dp[i-1][j-1] + 1
			} else {
				dp[i][j] = max(dp[i-1][j], dp[i][j-1])
			}
		}
	}
	lcsindex_src := make(map[int]int)
	lcsindex_dst := make(map[int]int)
	for i, j := l1, l2; i > 0 && j > 0; {
		if isSame_row(textsrc[i-1], textdst[j-1]) {
			lcsindex_src[i-1] = i - 1
			lcsindex_dst[j-1] = j - 1
			i--
			j--
		} else if dp[i-1][j] >= dp[i][j-1] {
			i--
		} else {
			j--
		}
	}
	return lcsindex_src, lcsindex_dst
}

func max(x, y int) int {
	if x > y {
		return x
	}
	return y
}

func min(x, y int) int {
	if x < y {
		return x
	}
	return y
}

func SheetDiff(fsrc *excelize.File, srcsheet string, fdst *excelize.File, dstsheet string) *Atabcontent {
	srcrows, err := fsrc.GetRows(srcsheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	dstrows, err := fdst.GetRows(dstsheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	lcsindex_src, lcsindex_dst := longestCommonSubsequence(srcrows, dstrows)
	lenpercolsrc := make([]int, len(srcrows[0]))
	lenpercoldst := make([]int, len(dstrows[0]))
	for i := 0; i < len(srcrows); i++ {
		for j := 0; j < len(srcrows[0]); j++ {
			if len(srcrows[i][j]) > lenpercolsrc[j] {
				lenpercolsrc[j] = len(srcrows[i][j])
			}
		}
	}
	for i := 0; i < len(dstrows); i++ {
		for j := 0; j < len(dstrows[0]); j++ {
			if len(dstrows[i][j]) > lenpercoldst[j] {
				lenpercoldst[j] = len(dstrows[i][j])
			}
		}
	}
	return NewAtabcontent(srcsheet, srcrows, lcsindex_src, dstrows, lcsindex_dst, lenpercolsrc, lenpercoldst, 0, []int{}, len(srcrows), len(srcrows[0]))
}

func ExcelDiff(filesrc string, filedst string) *excelize.File {
	fsrc, err := excelize.OpenFile(filesrc)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	srcworksheetname := fsrc.GetSheetList()
	fdst, err := excelize.OpenFile(filedst)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	dstworksheetname := fdst.GetSheetList()
	srcsheetlen := len(srcworksheetname)
	dstsheetlen := len(dstworksheetname)
	if srcsheetlen != dstsheetlen {
		fmt.Println("worksheet num is not the same")
		return nil
	}
	for i := 0; i < srcsheetlen; i++ {
		srcsheet := srcworksheetname[i]
		dstsheet := dstworksheetname[i]
		sheetcontentlist = append(sheetcontentlist, SheetDiff(fsrc, srcsheet, fdst, dstsheet))
	}
	return fsrc
}

func Table(w fyne.Window, srcentry *widget.Entry, dstentry *widget.Entry, index int) (*widget.Table, *widget.Table) {
	srctbl := widget.NewTable(nil, nil, nil)
	dsttbl := widget.NewTable(nil, nil, nil)
	srctbl.Length = func() (int, int) {
		if len(sheetcontentlist[index].srccontent) == 0 {
			fmt.Println("local content is empty")
			return 0, 0
		}
		return len(sheetcontentlist[index].srccontent), len(sheetcontentlist[index].srccontent[0])
	}
	dsttbl.Length = func() (int, int) {
		if len(sheetcontentlist[index].dstcontent) == 0 {
			fmt.Println("remote content is empty")
			return 0, 0
		}
		return len(sheetcontentlist[index].dstcontent), len(sheetcontentlist[index].dstcontent[0])
	}
	srctbl.CreateCell = func() fyne.CanvasObject {
		txt := canvas.NewText("table", color.Black)
		return txt
	}
	dsttbl.CreateCell = func() fyne.CanvasObject {
		txt := canvas.NewText("table", color.Black)
		return txt
	}
	srctbl.UpdateCell = func(id widget.TableCellID, template fyne.CanvasObject) {
		txt := template.(*canvas.Text)
		txt.Text = sheetcontentlist[index].srccontent[id.Row][id.Col]
		if _, ok := sheetcontentlist[index].lcsindex_src[id.Row]; !ok {
			txt.Color = color.RGBA{0, 255, 0, 255}
			if id.Row == srcidrow && hasElement(id.Col, sheetcontentlist[index].current_row_diff) {
				txt.Color = color.RGBA{255, 0, 0, 255}
			}
		} else {
			txt.Color = color.Black
		}
	}
	dsttbl.UpdateCell = func(id widget.TableCellID, template fyne.CanvasObject) {
		txt := template.(*canvas.Text)
		txt.Text = sheetcontentlist[index].dstcontent[id.Row][id.Col]
		if _, ok := sheetcontentlist[index].lcsindex_dst[id.Row]; !ok {
			txt.Color = color.RGBA{0, 255, 0, 255}
			if id.Row == dstidrow && hasElement(id.Col, sheetcontentlist[index].current_row_diff) {
				txt.Color = color.RGBA{255, 0, 0, 255}
			}
		} else {
			txt.Color = color.Black
		}
	}
	srctbl.OnSelected = func(id widget.TableCellID) {
		srcidrow = id.Row
		srcidcol = id.Col
		curindex = index
		cursrctbl = srctbl
		curdsttbl = dsttbl
		if presrcidrow != srcidrow {
			presrcidrow = srcidrow
			sheetcontentlist[curindex].current_row_diff = sheetcontentlist[curindex].current_row_diff[0:0]
			diff_label.SetText("差异数：" + strconv.Itoa(len(sheetcontentlist[curindex].current_row_diff)))
		}
		m := len(sheetcontentlist[index].srccontent)
		var n int
		if m == 0 {
			n = 1
		} else {
			n = len(sheetcontentlist[index].srccontent[0])
		}
		if id.Row >= m {
			for i := m; i <= id.Row; i++ {
				sheetcontentlist[index].srccontent = append(sheetcontentlist[index].srccontent, make([]string, n))
			}
			m = len(sheetcontentlist[index].srccontent)
		}
		//insert a new colume if select cols out of range
		/*if id.Col == -1 {
			for i := 0; i < m; i++ {
				sheetcontentlist[index].srccontent[i] = append(sheetcontentlist[index].srccontent[i], "")
			}
			n = len(sheetcontentlist[index].srccontent[0])
		}*/
		if id.Row < m && id.Row >= 0 && id.Col >= 0 && id.Col < n {
			srcentry.SetText(sheetcontentlist[index].srccontent[id.Row][id.Col])
		} else {
			srcentry.SetText("")
		}
		srcentry.OnSubmitted = func(s string) {
			sheetcontentlist[index].srccontent[id.Row][id.Col] = s
			if len(s) > sheetcontentlist[index].col_length_per_src[id.Col] {
				sheetcontentlist[index].col_length_per_src[id.Col] = len(s)
			}
			sheetcontentlist[index].lcsindex_src, sheetcontentlist[index].lcsindex_dst = longestCommonSubsequence(sheetcontentlist[index].srccontent, sheetcontentlist[index].dstcontent)
			for i := 0; i < len(sheetcontentlist[index].srccontent[0]); i++ {
				//srctbl.SetColumnWidth(i, float32(15*len(sheetcontentlist[index].srccontent[0][i])))
				srctbl.SetColumnWidth(i, float32(10*sheetcontentlist[index].col_length_per_src[i]))
			}
			dsttbl.Refresh()
			srctbl.Refresh()
			srctbl.Select(widget.TableCellID{Row: id.Row + 1, Col: id.Col})
		}
	}
	dsttbl.OnSelected = func(id widget.TableCellID) {
		dstidrow = id.Row
		dstidcol = id.Col
		curdsttbl = dsttbl
		cursrctbl = srctbl
		if predstidrow != dstidrow {
			predstidrow = dstidrow
			sheetcontentlist[curindex].current_row_diff = sheetcontentlist[curindex].current_row_diff[0:0]
			diff_label.SetText("差异数：" + strconv.Itoa(len(sheetcontentlist[curindex].current_row_diff)))
		}
		m := len(sheetcontentlist[index].dstcontent)
		var n int
		if m == 0 {
			n = 1
		} else {
			n = len(sheetcontentlist[index].dstcontent[0])
		}
		if id.Row >= m {
			for i := m; i <= id.Row; i++ {
				sheetcontentlist[index].dstcontent = append(sheetcontentlist[index].dstcontent, make([]string, n))
			}
			m = len(sheetcontentlist[index].dstcontent)
		}
		if id.Row < m && id.Col < n && id.Row >= 0 && id.Col >= 0 {
			dstentry.SetText(sheetcontentlist[index].dstcontent[id.Row][id.Col])
		} else {
			dstentry.SetText("")
		}
	}
	for i := 0; i < len(sheetcontentlist[index].srccontent[0]); i++ {
		//srctbl.SetColumnWidth(i, float32(15*len(sheetcontentlist[index].srccontent[0][i])))
		srctbl.SetColumnWidth(i, float32(10*sheetcontentlist[index].col_length_per_src[i]))
	}
	for i := 0; i < len(sheetcontentlist[index].dstcontent[0]); i++ {
		//dsttbl.SetColumnWidth(i, float32(15*len(sheetcontentlist[index].dstcontent[0][i])))
		dsttbl.SetColumnWidth(i, float32(10*sheetcontentlist[index].col_length_per_dst[i]))
	}
	return srctbl, dsttbl
}

func main() {
	//os.Setenv("FYNE_FONT", "giAlibaba-PuHuiTi-Medium.ttf")
	var src, dst string
	flag.StringVar(&src, "src", "", "src excel name")
	flag.StringVar(&dst, "dst", "", "dst excel name")
	flag.Parse()
	f := ExcelDiff(src, dst)
	myApp := app.New()
	myApp.Settings().SetTheme(&theme.MyTheme{})
	myWindow := myApp.NewWindow("ExcelMerge")
	//myWindow.CenterOnScreen()
	//myWindow.Resize(fyne.NewSize(1920, 1080))
	myWindow.SetFullScreen(true)
	tabnum := len(sheetcontentlist)
	var srcTabitems []*container.TabItem
	var dstTabitems []*container.TabItem
	srcentry := widget.NewEntry()
	dstentry := widget.NewEntry()
	srcname := widget.NewLabel("本地：" + src)
	dstname := widget.NewLabel("远端：" + dst)
	for i := 0; i < tabnum; i++ {
		srctbl, dsttbl := Table(myWindow, srcentry, dstentry, i)
		dstTabitems = append(dstTabitems, container.NewTabItem(sheetcontentlist[i].sheetname, dsttbl))
		srcTabitems = append(srcTabitems, container.NewTabItem(sheetcontentlist[i].sheetname, srctbl))
	}
	srctabs := container.NewAppTabs(srcTabitems...)
	srctabs.SetTabLocation(container.TabLocationBottom)
	dsttabs := container.NewAppTabs(dstTabitems...)
	dsttabs.SetTabLocation(container.TabLocationBottom)
	//btns
	bt_save := widget.NewButton("保存"+srcname.Text, func() {
		for i := 0; i < len(sheetcontentlist); i++ {
			rows := max(sheetcontentlist[i].initial_rows, len(sheetcontentlist[i].srccontent))
			cols := max(sheetcontentlist[i].initial_cols, len(sheetcontentlist[i].srccontent[0]))
			for m := 0; m < rows; m++ {
				for n := 0; n < cols; n++ {
					axis, _ := excelize.CoordinatesToCellName(n+1, m+1)
					if m < len(sheetcontentlist[i].srccontent) && n < len(sheetcontentlist[i].srccontent[0]) {
						f.SetCellValue(sheetcontentlist[i].sheetname, axis, sheetcontentlist[i].srccontent[m][n])
					} else {
						f.SetCellValue(sheetcontentlist[i].sheetname, axis, "")
					}
				}
			}
		}
		f.Save()
		dd := dialog.NewInformation("success!", "保存成功！", myWindow)
		dd.Show()
	})
	bt_local_change_row := widget.NewButton("保留本地所选行(ctrl+O)", accept_local_row)
	bt_insertarow := widget.NewButton("本地插入新行(crtl+R)", insert_a_row)
	bt_insertacol := widget.NewButton("本地插入新列(ctrl+L)", insert_a_col)
	bt_deletearow := widget.NewButton("删除本地选中行(alt+R)", delete_selected_row)
	bt_deleteacol := widget.NewButton("删除本地选中列(alt+L)", delete_selected_col)
	bt_incoming_change_row := widget.NewButton("保留远端所选行(ctrl+I)", accept_incoming_row)
	bt_incoming_change_col := widget.NewButton("保留远端所选列(ctrl+K)", accept_incoming_col)
	bt_insert_incomingrow_befor := widget.NewButton("插入远端所选行(ctrl+U)", insert_incoming_row_before)
	bt_find_diff := widget.NewButton("发现选中行差异", find_diff)
	bt_find_next_diff := widget.NewButton("下一个差异", find_next_diff)
	//layouts
	box1 := container.NewHSplit(srctabs, dsttabs)
	boxentry := container.NewHSplit(srcentry, dstentry)
	boxbuttons := container.NewVBox(bt_local_change_row, layout.NewSpacer(), bt_incoming_change_row, layout.NewSpacer(), bt_insert_incomingrow_befor,
		layout.NewSpacer(), bt_insertarow, layout.NewSpacer(), bt_deletearow, layout.NewSpacer(),
		bt_incoming_change_col, layout.NewSpacer(), bt_insertacol, layout.NewSpacer(), bt_deleteacol)
	boxleft := container.NewVBox(bt_find_diff, layout.NewSpacer(), diff_label, layout.NewSpacer(), bt_find_next_diff, layout.NewSpacer(), bt_save)
	boxname := container.NewHSplit(srcname, dstname)
	box := container.NewBorder(boxentry, boxname, boxleft, boxbuttons, box1)
	//shortcuts
	myWindow.Canvas().AddShortcut(&ctrlS, func(shortcut fyne.Shortcut) {
		for i := 0; i < len(sheetcontentlist); i++ {
			rows := max(sheetcontentlist[i].initial_rows, len(sheetcontentlist[i].srccontent))
			cols := max(sheetcontentlist[i].initial_cols, len(sheetcontentlist[i].srccontent[0]))
			for m := 0; m < rows; m++ {
				for n := 0; n < cols; n++ {
					axis, _ := excelize.CoordinatesToCellName(n+1, m+1)
					if m < len(sheetcontentlist[i].srccontent) && n < len(sheetcontentlist[i].srccontent[0]) {
						f.SetCellValue(sheetcontentlist[i].sheetname, axis, sheetcontentlist[i].srccontent[m][n])
					} else {
						f.SetCellValue(sheetcontentlist[i].sheetname, axis, "")
					}
				}
			}
		}
		f.Save()
	})
	myWindow.Canvas().AddShortcut(&ctrlL, func(shortcut fyne.Shortcut) {
		if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) {
			for i := 0; i < len(sheetcontentlist[curindex].srccontent); i++ {
				rear := append([]string{}, sheetcontentlist[curindex].srccontent[i][srcidcol:]...)
				sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i][0:srcidcol], "")
				sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i], rear...)
			}
			Refresh()
		}
	})
	myWindow.Canvas().AddShortcut(&altR, func(shortcut fyne.Shortcut) {
		if srcidrow >= 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) {
			sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent[:srcidrow], sheetcontentlist[curindex].srccontent[srcidrow+1:]...)
			/*for _, item := range sheetcontentlist[curindex].accept_local_rows_src {
				if item >= srcidrow {
					item -= 1
				}
			}*/
			Refresh()
		}
	})
	myWindow.Canvas().AddShortcut(&altL, func(shortcut fyne.Shortcut) {
		if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) {
			for i := 0; i < len(sheetcontentlist[curindex].srccontent); i++ {
				sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i][:srcidcol], sheetcontentlist[curindex].srccontent[i][srcidcol+1:]...)
			}
			Refresh()
		}
	})
	myWindow.Canvas().AddShortcut(&ctrlI, func(shortcut fyne.Shortcut) {
		if srcidrow >= 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) && dstidrow >= 0 && dstidrow < len(sheetcontentlist[curindex].dstcontent) {
			sheetcontentlist[curindex].srccontent[srcidrow] = append([]string{}, sheetcontentlist[curindex].dstcontent[dstidrow][0:]...)
			//sheetcontentlist[curindex].srccontent[srcidrow] = sheetcontentlist[curindex].dstcontent[dstidrow]
			Refresh()
		}
	})
	myWindow.Canvas().AddShortcut(&ctrlK, func(shortcut fyne.Shortcut) {
		if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) && dstidcol >= 0 && dstidcol < len(sheetcontentlist[curindex].dstcontent[0]) {
			msrc := len(sheetcontentlist[curindex].srccontent)
			mdst := len(sheetcontentlist[curindex].dstcontent)
			minm := min(msrc, mdst)
			for i := 0; i < minm; i++ {
				sheetcontentlist[curindex].srccontent[i][srcidcol] = sheetcontentlist[curindex].dstcontent[i][dstidcol]
			}
			if msrc > mdst {
				for i := minm; i < msrc; i++ {
					sheetcontentlist[curindex].srccontent[i][srcidcol] = ""
				}
			} else if msrc < mdst {
				for i := minm; i < mdst; i++ {
					sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent, make([]string, len(sheetcontentlist[curindex].srccontent[0])))
					sheetcontentlist[curindex].srccontent[i][srcidcol] = sheetcontentlist[curindex].dstcontent[i][dstidcol]
				}
			}
			Refresh()
		}
	})
	myWindow.Canvas().AddShortcut(&ctrlU, func(shortcut fyne.Shortcut) {
		insert_a_row()
		accept_incoming_row()
	})
	myWindow.Canvas().AddShortcut(&altQ, func(shortcut fyne.Shortcut) {
		myWindow.Close()
	})
	//shortcuts end
	myWindow.SetContent(box)
	myWindow.ShowAndRun()
}

func Refresh() {
	sheetcontentlist[curindex].lcsindex_src, sheetcontentlist[curindex].lcsindex_dst = longestCommonSubsequence(sheetcontentlist[curindex].srccontent, sheetcontentlist[curindex].dstcontent)
	for i := 0; i < len(sheetcontentlist[curindex].srccontent[0]); i++ {
		cursrctbl.SetColumnWidth(i, float32(10*sheetcontentlist[curindex].col_length_per_src[i]))
		//cursrctbl.SetColumnWidth(i, float32(15*len(sheetcontentlist[curindex].srccontent[0][i])))
	}
	cursrctbl.Refresh()
	curdsttbl.Refresh()
}

func accept_local_row() {
	//sheetcontentlist[curindex].accept_local_rows_src = append(sheetcontentlist[curindex].accept_local_rows_src, srcidrow)
	//sheetcontentlist[curindex].accept_local_rows_dst = append(sheetcontentlist[curindex].accept_local_rows_dst, dstidrow)
	copy(sheetcontentlist[curindex].dstcontent[dstidrow], sheetcontentlist[curindex].srccontent[srcidrow])
	Refresh()
}

func insert_a_row() {
	if srcidrow >= 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) {
		rear := append([][]string{}, sheetcontentlist[curindex].srccontent[srcidrow:]...)
		sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent[0:srcidrow], make([]string, len(rear[0])))
		sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent, rear...)
		Refresh()
	}
}

func insert_a_col() {
	if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) {
		//fmt.Println(sheetcontentlist[curindex].col_length_per_src)
		rear1 := append([]int{}, sheetcontentlist[curindex].col_length_per_src[srcidcol:]...)
		sheetcontentlist[curindex].col_length_per_src = append(sheetcontentlist[curindex].col_length_per_src[0:srcidcol], 1)
		sheetcontentlist[curindex].col_length_per_src = append(sheetcontentlist[curindex].col_length_per_src, rear1...)
		for i := 0; i < len(sheetcontentlist[curindex].srccontent); i++ {
			rear := append([]string{}, sheetcontentlist[curindex].srccontent[i][srcidcol:]...)
			sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i][0:srcidcol], "")
			sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i], rear...)
		}
		//fmt.Println(sheetcontentlist[curindex].col_length_per_src)
		Refresh()
	}
}

func delete_selected_row() {
	if srcidrow >= 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) {
		sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent[:srcidrow], sheetcontentlist[curindex].srccontent[srcidrow+1:]...)
		Refresh()
	}
}

func delete_selected_col() {
	if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) {
		//fmt.Println(sheetcontentlist[curindex].col_length_per_src)
		sheetcontentlist[curindex].col_length_per_src = append(sheetcontentlist[curindex].col_length_per_src[0:srcidcol], sheetcontentlist[curindex].col_length_per_src[srcidcol+1:]...)
		for i := 0; i < len(sheetcontentlist[curindex].srccontent); i++ {
			sheetcontentlist[curindex].srccontent[i] = append(sheetcontentlist[curindex].srccontent[i][:srcidcol], sheetcontentlist[curindex].srccontent[i][srcidcol+1:]...)
		}
		//fmt.Println(sheetcontentlist[curindex].col_length_per_src)
		Refresh()
	}
}

func accept_incoming_row() {
	if srcidrow >= 0 && srcidrow < len(sheetcontentlist[curindex].srccontent) && dstidrow >= 0 && dstidrow < len(sheetcontentlist[curindex].dstcontent) {
		//sheetcontentlist[curindex].srccontent[srcidrow] = append([]string{}, sheetcontentlist[curindex].dstcontent[dstidrow][0:]...)
		copy(sheetcontentlist[curindex].srccontent[srcidrow], sheetcontentlist[curindex].dstcontent[dstidrow])
		//sheetcontentlist[curindex].srccontent[srcidrow] = sheetcontentlist[curindex].dstcontent[dstidrow]
		Refresh()
	}
}

func insert_incoming_row_before() {
	insert_a_row()
	accept_incoming_row()
	insert_a_row()
	accept_incoming_row()
}

func accept_incoming_col() {
	if srcidcol >= 0 && srcidcol < len(sheetcontentlist[curindex].srccontent[0]) && dstidcol >= 0 && dstidcol < len(sheetcontentlist[curindex].dstcontent[0]) {
		msrc := len(sheetcontentlist[curindex].srccontent)
		insert_a_row()
		accept_incoming_row()
		mdst := len(sheetcontentlist[curindex].dstcontent)
		minm := min(msrc, mdst)
		for i := 0; i < minm; i++ {
			sheetcontentlist[curindex].srccontent[i][srcidcol] = sheetcontentlist[curindex].dstcontent[i][dstidcol]
		}
		if msrc > mdst {
			for i := minm; i < msrc; i++ {
				sheetcontentlist[curindex].srccontent[i][srcidcol] = ""
			}
		} else if msrc < mdst {
			for i := minm; i < mdst; i++ {
				sheetcontentlist[curindex].srccontent = append(sheetcontentlist[curindex].srccontent, make([]string, len(sheetcontentlist[curindex].srccontent[0])))
				sheetcontentlist[curindex].srccontent[i][srcidcol] = sheetcontentlist[curindex].dstcontent[i][dstidcol]
			}
		}
		sheetcontentlist[curindex].col_length_per_src[srcidcol] = sheetcontentlist[curindex].col_length_per_dst[dstidcol]
		Refresh()
	}
}
