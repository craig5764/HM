SetBatchLines, -1

xl := ComObjActive("Excel.Application")

;오늘 날짜 지정
Today := A_Year "-" A_MM "-" A_DD

;재고를 저장할 변수 초기 세팅
Inv := "Codigo`tDescription`tTTL`n" 

;Sheet명 지정
s := ["WM", "REF", "P3", "LÁMINA", "SEAL"]

;~ WinMinimize, % "ahk_id" id

;Sheet명 순서대로 재고 따기
for i, v in s
{
	
	;오늘 날짜인 Column 위치 찾기
	Loop {
		if(xl.Sheets(v).Cells(4, 5 + A_Index).Value = Today) {
			column := xl.Sheets(v).Cells(4, 5 + A_Index).Column
			break
		}
	}

	;D열에서 "STOCK" 있는 행에 있는 코드, 품명 그리고 D + Column 번호값(재고) 를 가져오기
	Loop {
		if(xl.Sheets(v).Cells(4 + A_Index, 4).Value = "STOCK") {
			Codigo := xl.Sheets(v).Cells(4 + A_Index, 2).Value
			Descrip := xl.Sheets(v).Cells(4 + A_Index, 3).Value
			Cant := xl.Sheets(v).Cells(4 + A_Index, column).Value
			Inv .= Codigo "`t" Descrip "`t" Cant "`n"
		}
		if(xl.Sheets(v).Cells(4 + A_Index, 4).Value = "")
			break
	}
	ToolTip, % "시트명 : " v "`n추출 완료"
}
clipboard := Inv
ToolTip
Msgbox, 재고 추출이 완료되었습니다.