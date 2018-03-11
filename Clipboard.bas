Attribute VB_Name = "Clipboard"

'    TableMacros - Tools for handling tables on Word
'
'    Written in 2018 by Francisco Gómez García <espectalll@kydara.com>
'
'    To the extent possible under law, the author(s) have dedicated all
'    copyright and related and neighboring rights to this software to the public
'    domain worldwide.
'    This software is distributed without any warranty.
'
'    You should have received a copy of the CC0 Public Domain Dedication along
'    with this software. If not, see
'    <http://creativecommons.org/publicdomain/zero/1.0/>.

Function CopyText(Text As String)
    ' Copy a string into the Windows clipboard

    Dim DataObj As New DataObject
    DataObj.SetText Text
    DataObj.PutInClipboard

End Function
