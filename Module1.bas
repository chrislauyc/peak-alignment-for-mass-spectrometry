Attribute VB_Name = "Module1"
Function Compare_Print(ByVal y As Collection, ByVal sh2 As Worksheet, ByVal sh1 As Worksheet)
'MsgBox y(1).get_dataMZ(1)
Dim iReply As Double



iReply = InputBox(Prompt:="enter the M/Z difference wish to sort (enter 0.01 )", _
            Title:="UPDATE MACRO")


Dim i As Integer, v As Integer, h As Integer
Dim diff As Double
Dim last As Boolean
last = False
' find the smallest and compare all of the data
h = 3

Do While last = False
    Dim smallest As Integer, x As Integer
    Dim r As Integer 'r is a variable used to keep track of the element of a column
    Dim count As Integer
    count = 1
    
    For v = 1 To y.count ' in here I check if all the columns are all empty (gone through)
    
    
        If y(v).j = y(v).size Then 'checking the size
        count = count + 1
        End If
        If count = y.count Then ' if number of empty columns equal number of total column
        last = True
        
        Exit Do
        End If
        
        
    Next v
    For v = 1 To y.count 'to make sure you are assigning a sample with things in it
        If y(v).j < y(v).size Then
            r = v
            smallest = y(r).j
            Exit For
        End If
    Next v
    

    'the element in question is the smallest element across the columns
    For v = 1 To y.count 'find the smallest of all the peaks
        If y(v).j < y(v).size Then
            'Dim a As Double, b As Double
            'a = y(v).get_dataMZ(y(v).j)
            'b = y(r).get_dataMZ(y(r).j)
            If y(v).get_dataMZ(y(v).j) < y(r).get_dataMZ(y(r).j) Then
            smallest = y(v).j
            r = v 'r is the sample where the smallest is in
            
           
            End If
        End If
    Next v
    If y(r).j < y(r).size Then
    y(r).Plus_j
    End If
    
    '****
    If h = 33 Then
    h = 33
    End If
    '****
    'printing the element in question
    sh2.Cells(h, r * 2 - 1) = y(r).get_dataMZ(smallest)
    sh2.Cells(h, r * 2) = y(r).get_dataAB(smallest)
    x = y(r).get_color(smallest)
    sh2.Cells(h, r * 2 - 1).Font.ColorIndex = y(r).get_color(smallest)
    sh2.Cells(h, r * 2).Font.ColorIndex = y(r).get_color(smallest)
    'starting the comparision with elements with other columns
    For v = 1 To y.count 'step through all the columns and compare
        If y(v).j < y(v).size Then
            
         
            If v = r Then
            Else
                diff = Abs(y(r).get_dataMZ(smallest) - y(v).get_dataMZ(y(v).j))
                'need to make sure it is printing the closest possible element
                'if the next element is smaller, don't print
                Dim next_diff As Double
                If smallest + 1 < y(r).size Then
                    next_diff = Abs(y(r).get_dataMZ(smallest + 1) - y(v).get_dataMZ(y(v).j))
                Else
                    next_diff = -1
                End If
                If diff < iReply Then
                    If next_diff = -1 Then
                        sh2.Cells(h, v * 2 - 1) = y(v).get_dataMZ(y(v).j)
                        sh2.Cells(h, v * 2) = y(v).get_dataAB(y(v).j)
                        sh2.Cells(h, v * 2 - 1).Font.ColorIndex = y(v).get_color(y(v).j)
                        sh2.Cells(h, v * 2).Font.ColorIndex = y(v).get_color(y(v).j)
                        y(v).Plus_j
                    Else
                        If next_diff < diff Then
                            'if the next element in question is actually closer, do nothing
                        Else
                            sh2.Cells(h, v * 2 - 1) = y(v).get_dataMZ(y(v).j)
                            sh2.Cells(h, v * 2) = y(v).get_dataAB(y(v).j)
                            sh2.Cells(h, v * 2 - 1).Font.ColorIndex = y(v).get_color(y(v).j)
                            sh2.Cells(h, v * 2).Font.ColorIndex = y(v).get_color(y(v).j)
                            y(v).Plus_j
                        End If
                    End If
                        
                ElseIf diff > iReply And diff < 0.1 Then
                    sh2.Cells(h, r * 2 - 1).Font.ColorIndex = 3
                    sh2.Cells(h, r * 2).Font.ColorIndex = 3
                    y(v).ChangeToRed (y(v).j)
                    
                    
                End If
                
            
            End If
        
        End If
    Next v
    

    
    
h = h + 1
Loop


'replacing empty cells with zeros for easier analysis. Can be commented out
Dim max_height As Integer, max_width As Integer
max_height = sh2.UsedRange.Rows.count
max_width = sh2.UsedRange.Columns.count
For h = 3 To max_height
    For v = 1 To max_width
        If IsEmpty(sh2.Cells(h, v)) Then
            sh2.Cells(h, v) = 0
        End If
    Next v
Next h
  
    
    
    
    
    

End Function

Sub sorting()
'This is the beginning of the algorithm
'The algorithm first takes all the columns and store each of them in each Samples class
'It then creates a new worksheet and copy the header. It will then prompt for a tolerance value
'using the tolerance value, it will start comparing and printing the data in the new sheet

'****
'Thanks for using! Please also cite this program.
'Author: Chris Y. Lau
'Program title: Peak Alignment for Mass Spectrometry
'Last edited: 7/22/18
'****

Dim nOfdata As Integer
nOfdata = (ActiveSheet.UsedRange.Columns.count) / 2
Dim N_ofm_z As Integer
Dim v As Integer
Dim thisSample As Samples
Dim m_z As Double, Abun As Double
Set Samples_Collection = New Collection
For v = 1 To nOfdata 'going across
    Dim i As Integer
    N_ofm_z = Columns(v * 2 - 1).SpecialCells(xlCellTypeConstants, 23).Cells.count - 1
    Set thisSample = New Samples
    
    Call thisSample.setuparray(N_ofm_z)
    For i = 1 To N_ofm_z 'going down and adding each data
        m_z = Cells(i + 2, v * 2 - 1)
        Abun = Cells(i + 2, v * 2)
        Call thisSample.Add_data(m_z, Abun, 1)
        If i = N_ofm_z Then
        Call thisSample.reset_i
        End If

    Next i
       
       
    Samples_Collection.Add Item:=thisSample, Key:=CStr(v)
    
 Next v
 



    
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Set sh1 = ActiveSheet
    Set sh2 = Sheets.Add
    
For i = 1 To 2
    For v = 1 To nOfdata * 2
    sh1.Cells(i, v).Copy sh2.Cells(i, v)
    Next v
Next i


Call Compare_Print(Samples_Collection, sh2, sh1)
End Sub

