Sub ExportUIFunctionTableV()

	Dim FilePath As String
	Dim CellData As String
	Dim CellData1 As String
	Dim N As Long
	Dim SettingItemColumn As Integer
	Dim column_index As Integer
	Dim UICol As Long

	CellData = ""

	FilePath = ActiveWorkbook.Path & "\generated.cpp"

	Rem SetItemCol 13th column or M column
	SettingItemColumn = 15
	column_index = 2

	Open FilePath For Output As #2

	Rem ---------------  Start Print of Copy Table ---------------------------
		CellData = "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cpy_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "C", SettingItemColumn, 21, column_index
	Rem --------------- END Print of Copy Table ---------------------------


	Rem ---------------  Start Print of Interrupt Copy Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_icpy_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "D", SettingItemColumn, 21, column_index
	Rem --------------- END Print of Interrupt Copy Table ---------------------------



	Rem ---------------  Start Print of Send Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_snd_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "E", SettingItemColumn, 23, column_index
	Rem --------------- END Print of Send Table ---------------------------



	Rem ---------------  Start Print of FAX Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_fax_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "F", SettingItemColumn, 25, column_index
	Rem --------------- END Print of FAX Table ---------------------------



	Rem ---------------  Start Print of Custom Box Print Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbprt_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "G", SettingItemColumn, 27, column_index
	Rem --------------- END Print of Custom Box Print Table ---------------------------



	Rem ---------------  Start Print of Custom Box Send Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbsnd_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "H", SettingItemColumn, 29, column_index
	Rem --------------- END Print of Custom Box Send Table ---------------------------



	Rem ---------------  Start Print of Custom Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "I", SettingItemColumn, 31, column_index
	Rem --------------- END Print of Custom Box Store Table ---------------------------



	Rem ---------------  Start Print of USB Box Print Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_ubprt_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "J", SettingItemColumn, 27, column_index
	Rem --------------- END Print of USB Box Print Table ---------------------------



	Rem ---------------  Start Print of USB Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_ubstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "K", SettingItemColumn, 33, column_index
	Rem --------------- END Print of USB Box Store Table ---------------------------



	Rem ---------------  Start Print of Job Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_jbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "L", SettingItemColumn, 35, column_index
	Rem ---------------  END Print of Job Box Store Table ---------------------------



	Rem ---------------  Start Print of Polling Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_pbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "M", SettingItemColumn, 37, column_index
	Rem ---------------  End Print of Polling Box Store Table ---------------------------


	Rem Print Terminating Closing braces
	CellData = "};"

	Rem Print to File
	Print #2, CellData

	Rem Close File
	Close #2

	MsgBox ("Done Generating Virgo")

End Sub


Sub ExportUIFunctionTableIris2020()

	Dim FilePath As String
	Dim CellData As String
	Dim CellData1 As String
	Dim N As Long
	Dim SettingItemColumn As Integer
	Dim column_index As Integer
	Dim UICol As Long

	CellData = ""

	FilePath = ActiveWorkbook.Path & "\generated.cpp"

	Rem SetItemCol 13th column or M column
	SettingItemColumn = 15
	column_index = 1

	Open FilePath For Output As #2

	Rem ---------------  Start Print of Copy Table ---------------------------
		CellData = "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cpy_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "C", SettingItemColumn, 21, column_index
	Rem --------------- END Print of Copy Table ---------------------------


	Rem ---------------  Start Print of Interrupt Copy Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_icpy_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "D", SettingItemColumn, 21, column_index
	Rem --------------- END Print of Interrupt Copy Table ---------------------------



	Rem ---------------  Start Print of Send Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_snd_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "E", SettingItemColumn, 23, column_index
	Rem --------------- END Print of Send Table ---------------------------



	Rem ---------------  Start Print of FAX Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_fax_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "F", SettingItemColumn, 25, column_index
	Rem --------------- END Print of FAX Table ---------------------------



	Rem ---------------  Start Print of Custom Box Print Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbprt_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "G", SettingItemColumn, 27, column_index
	Rem --------------- END Print of Custom Box Print Table ---------------------------



	Rem ---------------  Start Print of Custom Box Send Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbsnd_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "H", SettingItemColumn, 29, column_index
	Rem --------------- END Print of Custom Box Send Table ---------------------------



	Rem ---------------  Start Print of Custom Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_cbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "I", SettingItemColumn, 31, column_index
	Rem --------------- END Print of Custom Box Store Table ---------------------------



	Rem ---------------  Start Print of USB Box Print Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_ubprt_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "J", SettingItemColumn, 27, column_index
	Rem --------------- END Print of USB Box Print Table ---------------------------



	Rem ---------------  Start Print of USB Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_ubstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "K", SettingItemColumn, 33, column_index
	Rem --------------- END Print of USB Box Store Table ---------------------------



	Rem ---------------  Start Print of Job Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_jbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "L", SettingItemColumn, 35, column_index
	Rem ---------------  END Print of Job Box Store Table ---------------------------



	Rem ---------------  Start Print of Polling Box Store Table ---------------------------
		CellData = "};" & vbNewLine & "const" & vbTab & "FUNCTIONAL_ID_TBL" & vbTab & "functional_id_pbstr_tbl[]" & vbTab & "= {"
		Print #2, CellData

		Iterate "M", SettingItemColumn, 37, column_index
	Rem ---------------  End Print of Polling Box Store Table ---------------------------


	Rem Print Terminating Closing braces
	CellData = "};"

	Rem Print to File
	Print #2, CellData

	Rem Close File
	Close #2

	MsgBox ("Done Generating 2020")
	ExportUIFunctionTableV

End Sub

Private Sub Iterate(string_column As String, i_table_column As Integer, i_table_uifunc As Integer, i_model_column As Integer)

	Rem N count number of valid rows
	Range("A1").Select
	N = Cells(Rows.Count, string_column).End(xlUp).Row

	Rem str_temp refers to button for sorting
	str_temp = string_column + "3"

	Rem sort ascending
	Range("A3:AK3" & N).Sort Key1:=Range(str_temp), Order1:=xlAscending, Header:=xlYes

	For i = 4 To N

	If (ActiveCell(i, string_column).Value) = "" Then
		CellData = ""
	ElseIf (InStr(1, ActiveCell(i, i_table_column).Value, "#")) > 0 And (InStr(1, ActiveCell(i, i_model_column).Value, "O")) > 0 Then
		CellData = Trim(ActiveCell(i, i_table_column).Value)
		Print #2, CellData
		CellData = ""
	Else
		If (InStr(1, ActiveCell(i, i_model_column).Value, "O")) > 0 Then
			CellData = vbTab & "{ " + CellData + Trim(ActiveCell(i, i_table_column).Value) + ", """
			CellData = CellData + Trim(ActiveCell(i, i_table_uifunc).Value) + """ },"
			Print #2, CellData
			CellData = ""
		End If
	End If

	Next i
End Sub

