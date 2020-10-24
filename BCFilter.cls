VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BCFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' Helper type, representing a filter in a BigCommerce GET request
'' See https://developer.bigcommerce.com/api-docs/getting-started/filtering

Public Enum BCFilterOp
    BCF_EQ      ' Equals
    BCF_GE      ' Greater than or equal to
    BCF_LE      ' Less than or equal to
    BCF_GT      ' Greater than
    BCF_LT      ' Less than
    BCF_LIKE    ' SQL LIKE
    BCF_IN      ' SQL IN
    BCF_NOTIN   ' SQL NOT IN
End Enum

Private pField As String
Private pOperation As BCFilterOp
Private pValue As Variant


Public Sub Init(Field As String, Operation As BCFilterOp, Value As Variant)
    pField = Field
    pOperation = Operation
    pValue = Value
End Sub

Public Property Get Text() As String
    Dim Value As String
    Dim ValueType As Integer
    Dim DateValue As Date
    
    ValueType = VarType(pValue)
    
    If ValueType >= vbArray Then
        Value = Join(pValue, ",")
    ElseIf ValueType = vbDate Then
        DateValue = CDate(pValue)
        Value = JsonConverter.ConvertToIso(DateValue)
    Else
        Value = pValue
    End If
    
    Select Case pOperation
        Case BCF_EQ      ' Equals
            Text = pField & "=" & Value
        Case BCF_GE      ' Greater than or equal to
            Text = pField & ":min=" & Value
        Case BCF_LE      ' Less than or equal to
            Text = pField & ":max=" & Value
        Case BCF_GT      ' Greater than
            Text = pField & ":greater=" & Value
        Case BCF_LT      ' Less than
            Text = pField & ":less=" & Value
        Case BCF_LIKE    ' SQL LIKE
            Text = pField & ":like=" & Value
        Case BCF_IN      ' SQL IN
            Text = pField & ":in=" & Value
        Case BCF_NOTIN   ' SQL NOT IN
            Text = pField & ":not_in=" & Value
    End Select
End Property

