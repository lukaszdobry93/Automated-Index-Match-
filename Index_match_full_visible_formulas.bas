Attribute VB_Name = "Index_match_widoczne_formuly"
Sub Makro_widoczne_formuly()

' MACRO APPLICATION:

' This macro allows users for dynamic values search from source table, using Index + Match + Match combination.
' Code includes filling only one table, but it is possible to split makro for two parts for:

'   1) Defining source table range,
'   2) Combining source data with target table(s), where look up values should be placed.

' FORUMLA CODE:
   
   
   
'Setting variable range, which will be data source for our research.
'We assuming that, data Source Table includes B6 cell.
    
Dim tbl As Range
Set tbl = Sheet1.Range("b6").CurrentRegion
        
tbl_address = tbl.Address       ' For the entire formula purposes.
        
        
'Setting variables as strings, for the entire formula purposes
        
Dim match_ax_range_address, match_ay_range_address As String
    
tbl.Select
    
    
'Fixing first row (of data table) range, for the "Match" formula purposes.
    
Dim match_ax_range As Range
Set match_ax_range = Selection.Rows(1)
    
    
match_ax_range_address = match_ax_range.Address
match_ax_range_address = Replace(match_ax_range_address, "$", "")   'Setting address as string, for the entire "Index + Match + Match" formula puproses.
    
tbl.Select
    
    
'Fixing first column (of data table) range, for the "Match" formula purposes.

Dim match_ay_range As Range
Set match_ay_range = Selection.Columns(1)
    
match_ay_range_address = match_ay_range.Address
match_ay_range_address = Replace(match_ay_range_address, "$", "")   'Setting address as string, for the entire "Index + Match + Match" formula puproses.

'Target table clearing (in case, if not empty).

Dim table_to_clear As Range
Set table_to_clear = Range("S19").CurrentRegion
table_to_clear.Offset(1, 1).Resize(table_to_clear.Rows.Count - 1, table_to_clear.Columns.Count - 1).ClearContents

Range("S19").Select     'Target table first cell to fill, with targeted values.

Dim x As Integer
Dim col_count As Integer
Dim row_count As Integer


'Number of columns and rows of Target Table, will be necassary for looping purposes, and filling all the empty cells in Target Table.

col_count = Range(ActiveCell.Offset(-1, 0), ActiveCell.Offset(-1, 0).End(xlToRight)).Columns.Count
row_count = Range(ActiveCell.Offset(0, -1), ActiveCell.Offset(0, -1).End(xlDown)).Rows.Count

Dim ay, ax, ay_value, ax_value As String

For y = 1 To row_count      'First loop, dependend of number of rows in Target Table.

ay = ActiveCell.Offset(0, -1).End(xlToLeft).Value

ay_value = Chr(34) & ActiveCell.Offset(0, -1).End(xlToLeft).Value & Chr(34)     'For the proper naming "Match" first argument in the entire formula


'Check if the table search for proper first "Match" argument.

If IsNumeric(ay) Or ActiveCell.Offset(0, -1).End(xlToLeft).Value = "" Then

ay_value = Chr(34) & ActiveCell.Offset(0, -1).Value & Chr(34)

End If

    For x = 1 To col_count      'Second loop, dependend of number of columns in Target Table.

    ax = ActiveCell.End(xlUp).Value

    ax_value = Chr(34) & ActiveCell.End(xlUp).Value & Chr(34)
    
    
    'Check if the table search for proper first "Match" argument.

    If IsNumeric(ax) Or ActiveCell.End(xlUp).Value = "" Then

    ax = ActiveCell.Offset(-1, 0).End(xlUp).Value

    ax_value = Chr(34) & ActiveCell.Offset(-1, 0).End(xlUp).Value & Chr(34)     'For the proper naming "Match" first argument in the entire formula

    End If


    ' Entire Formula. For simplicity, formula is just String type "cell value". This solution allows to view formula in Excell Sheets,
    'and still be treated as formula by Excell program:
    
    ActiveCell.Value = _
        "=INDEX(" & tbl_address & ", MATCH(" & ay_value & ", " & match_ay_range_address & ", 0), MATCH(" & ax_value & ", " & match_ax_range_address & ", 0))"


    ActiveCell.Offset(0, 1).Select

    Next

ActiveCell.Offset(1, 0).End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select

Next

End Sub
