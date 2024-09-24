# OOOLE
Object-oriented implementation of OLE2 syntax in SAP ABAP
## vba Parameter reference table 
- https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa221100(v=office.11)?redirectedfrom=MSDN
## basic syntax
```vba
Sub ExcelOperations()
    Dim EXCEL As Object
    Set EXCEL = CreateObject("Excel.Application")

    EXCEL.Visible = True

    Dim WORKBOOKS As Object
    Set WORKBOOKS = EXCEL.Workbooks
    Dim WORKBOOK As Object
    Set WORKBOOK = WORKBOOKS.Add

    Dim SHEET As Object
    Set SHEET = EXCEL.ActiveSheet

    SHEET.Cells(1, 1).Value = "TEST"
End Sub
```
```abap
DATA(EXCEL) = NEW CL_OLE_EXCEL( ).
EXCEL->SET_PROPERTY( E_PROPERTY = 'VISIBLE' E_VALUE = 1 ).
DATA(WORKBOOKS) = EXCEL->CALL_METHOD_OF( E_METHOD = 'WORKBOOKS' ).
DATA(WORKBOOK) = WORKBOOKS->CALL_METHOD_OF( E_METHOD = 'ADD' ).
DATA(SHEET) = EXCEL->CALL_METHOD_OF( E_METHOD = 'ACTIVESHEET' ).
DATA(CELL) = SHEET->CALL_METHOD_OF( E_METHOD = 'CELLS' E_ARG1 = 1 E_ARG2 = 1 ).
CELL->SET_PROPERTY( E_PROPERTY = 'VALUE' E_VALUE = 'TEST' ).
```
## encapsulated method
```vba
Sub ExcelOperations()
    Dim OOOLE As Object
    Set OOOLE = CreateObject("Excel.Application")

    OOOLE.Visible = True
    OOOLE.Workbooks.Add
    OOOLE.Sheets(2).Activate
    OOOLE.ActiveSheet.Cells(2, 2).Value = "TEST"

    With OOOLE.ActiveSheet.Cells(2, 2).Font
        .Name = "Arial"
        .Size = 15
        .Bold = True
        .Color = RGB(0, 0, 255) 
        .Italic = True
        .Underline = xlUnderlineStyleSingleAccounting
    End With

    OOOLE.ActiveSheet.Cells(2, 2).Interior.Color = RGB(255, 204, 0)

    With OOOLE.ActiveSheet.Cells(3, 3).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    With OOOLE.ActiveSheet.Cells(3, 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    With OOOLE.ActiveSheet.Cells(3, 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    With OOOLE.ActiveSheet.Cells(3, 3).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    OOOLE.ActiveSheet.Cells(4, 4).Value = 1.23
    OOOLE.ActiveSheet.Cells(4, 4).NumberFormat = "0.00"

    OOOLE.ActiveSheet.Cells(5, 4).Value = DateSerial(2012, 2, 1)
    OOOLE.ActiveSheet.Cells(5, 4).NumberFormat = "M/D/YYYY"
    
    OOOLE.ActiveSheet.Cells(6, 4).NumberFormat = "# ?/?"
    OOOLE.ActiveSheet.Cells(6, 4).Value = 1 / 2

    OOOLE.Quit
End Sub
```
```abap
DATA(OOOLE) = NEW CL_OLE_EXCEL( ).
OOOLE->SETPROPERTY( E_VISIBLE = 1 ).
OOOLE->WORKBOOKS( )->ADD( ).
OOOLE->WORKSHEETS( 2 )->ACTIVATE( ).
*WORKBOOKS->OPEN('C:\Users\ngj\Desktop\test111.xlsx').
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 2 E_COL = 2 )->VALUE( 'TEST' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 2 E_COL = 2 )->FONT( )->SETPROPERTY( E_NAME = 'Arial'
                                                                           E_SIZE = 15
                                                                           E_BOLD = 1
                                                                           E_COLOR = -16776961
                                                                           E_TINTANDSHADE = 0
                                                                           E_ITALIC = 1
                                                                           E_UNDERLINE = 2 ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 2 E_COL = 2 )->INTERIOR( )->SETPROPERTY( E_COLOR = 15773696 ).

OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 3 E_COL = 3 )->BORDERS( '7' )->SETPROPERTY( E_LINESTYLE = '1' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 3 E_COL = 3 )->BORDERS( '8' )->SETPROPERTY( E_LINESTYLE = '1' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 3 E_COL = 3 )->BORDERS( '9' )->SETPROPERTY( E_LINESTYLE = '1' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 3 E_COL = 3 )->BORDERS( '10' )->SETPROPERTY( E_LINESTYLE = '1' E_WEIGHT = 4 ).

OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 4 E_COL = 4 )->VALUE( '1.23' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 4 E_COL = 4 )->SETPROPERTY( E_NUMBERFORMAT = '0.00' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 5 E_COL = 4 )->VALUE( '02/01/2012' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 5 E_COL = 4 )->SETPROPERTY( E_NUMBERFORMAT = 'M/D/YYYY' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 6 E_COL = 4 )->SETPROPERTY( E_NUMBERFORMAT = '0.00' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 6 E_COL = 4 )->VALUE( '1/2' ).
OOOLE->ACTIVESHEET( )->CELLS( E_ROW = 6 E_COL = 4 )->SETPROPERTY( E_NUMBERFORMAT = '# ?/?' ).

OOOLE->QUIT( ).
```
