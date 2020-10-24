VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BCProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' BCProduct
''
'' Utility class representing products from BigCommerce

Private pData As Scripting.Dictionary
Private pCustomFields As Scripting.Dictionary

' Data is a Scripting.Dictionary object as returned by
' the catalog/products API call
Public Sub Init(Data As Scripting.Dictionary)
    Dim Field As Variant
    
    Set pCustomFields = Nothing
    Set pData = Data
    If pData.Exists("custom_fields") Then
        Set pCustomFields = New Scripting.Dictionary
        For Each Field In pData.Item("custom_fields")
            pCustomFields.Add Field.Item("name"), Field.Item("value")
        Next
    End If
End Sub

Public Property Get Data() As Scripting.Dictionary
    Set Data = pData
End Property

Public Property Get CustomField(Key As String) As String
    If pCustomFields Is Nothing Then
        CustomField = ""
    Else
        CustomField = pCustomFields.Item(Key)
    End If
End Property
