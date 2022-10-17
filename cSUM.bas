'-------------------------------------------------------------------------------
' Program Name: cSUM R0
' Created By:   Max Rocha (email: max.warocha@gmail.com)
' Date Created: 10-17-2022
' Function:     Simple Excel-VBA Function to Sum the Values of Cells Based on 
'               their Background Color.
'
' Usage:         =cSUM(A1:A10, B1)
'                            [Range to Sum, Cell w/ Desired BgColor]
'                =cSUM(A1:A10, "#FFFFFF")
'                            [Range to Sum, Desired BGColor in Hex]
'                =cSUM(A1:A10, "i30")    
'                            [Range to Sum, Desired BGColor in Excel ColorIndex]
'                =cGET(B1)
'                            [Cell w/ Unkown BgColor. Return BGColor in Hex]
' 
' Note:         It Ignores Conditional Formatting!
'               Refresh the Funct. after Changing the BGColor of the Cells!
'-------------------------------------------------------------------------------
' Revision History
' Rev  By     Date    Description
'-------------------------------------------------------------------------------
' 0    MR   10-17-22   Initial version
'-------------------------------------------------------------------------------


Function cSUM(Range As Range, Color As Variant)
'Sums all of the cells with a desired background color
    
	Application.Volatile
   
	Dim sum As Double
	Dim modColor As Double
    
	''Use function ProcessColor() to convert color
	modColor = ProcessColor(Color)

	''Read range. Sum if cell have the desired background color
	'''Excel ColorIndex
	If Left(Color, 1) = "i" Then
		For Each Cell In Range
			If Cell.Interior.ColorIndex = CDec(Mid(Color, 2)) And IsNumeric(Cell.Value) = True Then
				sum = sum + Cell.Value
			End If
		Next
	'''modColor
	Else
		For Each Cell In Range
			If Cell.Interior.Color = modColor And IsNumeric(Cell.Value) = True Then
				sum = sum + Cell.Value
			End If
		Next
    End If
    
    ''Return Sum
    cSUM = sum

End Function


Function cGET(r As Range)
'Get the background color in HEX["#xxxxxx"] of a cell 
    
    Application.Volatile
	
    cGET = cRGB2HEX(r.Cells(1, 1).Interior.Color)

End Function


Private Function cRGB2HEX(clr As Variant)
'Converts RGB[long] colors in HEX["#xxxxxx"]
    
    Dim clrR, clrG, clrB As String
	
    clrR = Format(Hex(clr Mod 256), "00")
    clrG = Format(Hex((clr \ 256) Mod 256), "00")
    clrB = Format(Hex(clr \ 65536), "00")

    cRGB2HEX = "#" & clrR & clrG & clrB

End Function


Private Function cHEX2RGB(clr As Variant)
'Converts HEX["#xxxxxx"] colors in RGB[long]
    
    Dim clrR, clrG, clrB As String
    
    clrR = CDec("&h0" & Mid(clr, 2, 2))
    clrG = CDec("&h0" & Mid(clr, 4, 2))
    clrB = CDec("&h0" & Mid(clr, 6, 2))
    
    cHEX2RGB = RGB(clrR, clrG, clrB)
    
End Function


Private Function ProcessColor(clr As Variant)
'Process color input

    On Error GoTo exception
    
    Dim clrR, clrG, clrB As String

	''If clr is a cell, get its background color
    If TypeOf clr Is Range Then
        ProcessColor = clr.Cells(1, 1).Interior.Color
    ''Else, convert color from HEX("#xxxxxx") to RGB(long)
	Else
        ProcessColor = cHEX2RGB(clr)
    End If
    
    Exit Function
    
exception:
    ProcessColor = RGB(0, 0, 0)
    
End Function