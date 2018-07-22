VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Samples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class contains all the peaks for one column
Private listOfm_z() As New data
Private i As Integer
Public j As Integer ' j is the index of the smallest element yet to be gone through
Public size As Integer



Public Sub setuparray(ByVal N_ofm_z As Integer)
    ReDim listOfm_z(1 To N_ofm_z)
    size = N_ofm_z
    i = 1
    j = 1
End Sub
Public Sub ChangeToRed(ByVal index As Integer)
    listOfm_z(index).color = 3
End Sub
Public Sub ChangeToBlack(ByVal index As Integer)
    listOfm_z(index).color = 1
End Sub

Public Sub Add_data(ByVal MZ As Double, ByVal Ab As Double, ByVal color As Integer)
    listOfm_z(i).m_z = MZ
    listOfm_z(i).Abun = Ab
    listOfm_z(i).color = 1
    i = i + 1
    
End Sub

Public Function get_dataMZ(ByVal x As Integer)
    get_dataMZ = listOfm_z(x).m_z
    
End Function
Public Function get_dataAB(ByVal x As Integer)
    get_dataAB = listOfm_z(x).Abun

End Function
Public Function get_color(ByVal x As Integer)
    get_color = listOfm_z(x).color
End Function
Public Function reset_i()
i = 1
End Function
Public Function Plus_j()
j = j + 1
End Function

