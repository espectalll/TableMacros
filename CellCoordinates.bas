Attribute VB_Name = "CellCoordinates"

'    TableMacros - Tools for handling tables on Word
'
'    Written in 2018 by Francisco G�mez Garc�a <espectalll@kydara.com>
'
'    To the extent possible under law, the author(s) have dedicated all
'    copyright and related and neighboring rights to this software to the public
'    domain worldwide.
'    This software is distributed without any warranty.
'
'    You should have received a copy of the CC0 Public Domain Dedication along
'    with this software. If not, see
'    <http://creativecommons.org/publicdomain/zero/1.0/>.

Sub showCellCoordinates()
    ' Show a popup with the coordinates of a selected cell
    
    If Selection.Information(wdWithInTable) Then
        
        Set CoordBar = CommandBars.Add(Name:="Table Coordinates", _
            Position:=msoBarPopup, Temporary:=True)
     
        Dim CellCoord As String
        Dim CellChar As Byte
        n = Selection.Cells(1).ColumnIndex
        
        Do
            
            CellChar = ((n - 1) Mod 26)
            CellCoord = Chr(CellChar + 65) & CellCoord
            n = (n - CellChar) \ 26
            
        Loop While n > 0
        
        CellCoord = CellCoord & CStr(Selection.Cells(1).RowIndex)
        
        With CoordBar
            .Controls.Add Type:=msoControlButton
            .Controls(1).Caption = "Table Coordinates: " & CellCoord
            .Controls(1).OnAction = Clipboard.CopyText(CellCoord)
            .ShowPopup
        End With
        
    End If
    
End Sub

