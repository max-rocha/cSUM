'-------------------------------------------------------------------------------
' Program Name: cSUM R0   
' Created By:   Max Rocha (email: max.warocha@gmail.com)
' Date: 10-17-2022
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
'               Refresh the Funct. after Modifying the BGColor of the Cells!
'
'-------------------------------------------------------------------------------
' Revision History
' Rev  By     Date    Description
'-------------------------------------------------------------------------------
' 0    MR   10-17-22   Initial version
'-------------------------------------------------------------------------------
