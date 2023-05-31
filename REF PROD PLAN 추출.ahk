SetBatchLines, -1

xl := ComObjActive("Excel.Application")
RC := xl.Sheets("RC SEQUENCE")

Data := "Fecha`tTurno`tItem`tCantidad`n"

;표 Column 끝 값 column 변수에 담기
Loop {
	if(RC.Cells(8, 10 + A_Index).MergeArea.Cells(1).Value = "") {
		column := RC.Cells(8, 10 + A_Index).MergeArea.Cells(1).Column
		break
	}
}

;표 Row 끝 값 row 변수에 담기
Loop {
	if(RC.Cells(8 + A_Index, 6).Value = "") {
		row := RC.Cells(8 + A_Index, 6).Row
		break
	}
}

;Loop 로 Row 를 하나씩 늘려가며 Column 10 ~ 37 까지 값을 찾아서 빈값이 아닐 경우 누적 변수에 Item 명과 수량 담기
r := 10
Loop {
	Loop {
		if(9 + A_Index = column)
			break
		if(RC.Cells(r, 9 + A_Index).Value != "") {
			Shift := RC.Cells(9, 9 + A_Index).Value = "1 st" ? "Dia" : "Noche"
			Data .= RC.Cells(8, 9 + A_Index).MergeArea.Cells(1).Value "`t" Shift "`t"RC.Cells(r, 6).Value "`t" RC.Cells(r, 9 + A_Index).Value "`n"
		}
	}
	r ++
	if(r = row)
		break
}
TrayTip, Complete, 계획 취합이 완료되었습니다.
clipboard := Data