Attribute VB_Name = "VB6Functions"
'    --------------------------------------------------------------------------
'    EzColorTest HTML Editor Color Coding Test
'    Copyright (C) 2000  Eric Banker
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'    --------------------------------------------------------------------------

Option Explicit
' These are vb6 functions not in vb5. If using vb 6 comment out these functions or
' remove this module from the project... note that i have edited some of these functions
' for my use and the vb 6 functions may not work the same. You could always change the
' name of these functions and replace the call in the code. These functions were taken from
' support.microsoft.com

Public Function RevInStr(ByVal sIn As String, sFind As String, Optional nStart As Long = 1, Optional bCompare As VbCompareMethod = vbBinaryCompare) As Long
Dim nPos As Long
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        RevInStr = 0
    Else
        RevInStr = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function
 
' End VB6 functions
