VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent AddText[(arg1, arg2, ... , argn)]
Public Event AddText()
'local variable(s) to hold property value(s)
Private mvarPagesNumber As Long 'local copy
Private mvarFormCaption As String 'local copy
Private mvarPrintingMode As Long 'local copy
Private mvarColumns As Integer
Private mvarSizes As Variant
Private mvarSums As Variant
Private mvarDataFormats As Variant
Private mvarText As String
Private mvarColumnBlankText As String
Private mvarColumnHeader As Variant
Private mvarAllow As Boolean
Private isHeader As Boolean
Public Enum Alignment
        prnRight = 1
        prnLeft = 2
        prnCenter = 3
End Enum
Public Enum DataFormats
       prnText = 1
       prnDate = 2
       prnCurrency = 3
       prnNumber = 4
       prnBarCode = 5       'That's it! You can put barcode on your
End Enum                    'reports! I've made the TT font myself!
            
Public Sub DrawLine()
For i% = 0 To mvarColumns - 1
    linha = linha + mvarSizes(i%)
Next i%
WriteText String(linha, "-"), True
End Sub

Public Sub AddTextToColumn(ColumnIndex As Integer, theText As String, prnAlign As Alignment, Optional isHeader As Boolean = False)
start = 0

If ColumnIndex = 1 Then
   start = 1
Else
   For i% = 0 To ColumnIndex - 2
       start = start + mvarSizes(i%)
   Next i%
End If
If prnAlign = prnLeft Then
   start = start + 1
ElseIf prnAlign = prnRight Then
   start = start + (mvarSizes(ColumnIndex - 1) - Len(theText)) - 1
ElseIf prnAlign = prnCenter Then
   start = start + ((mvarSizes(ColumnIndex - 1) - Len(theText)) / 2)
End If
If mvarDataFormats(ColumnIndex - 1) = prnCurrency Or _
   mvarDataFormats(ColumnIndex - 1) = prnNumber Then
   If Not isHeader And mvarAllow Then
      mvarSums(ColumnIndex - 1) = CStr(CCur(mvarSums(ColumnIndex - 1)) + CCur(theText))
   End If
End If
Mid(mvarColumnBlankText, IIf(mvarDataFormats(ColumnIndex - 1) = prnNumber And Not isHeader, start - 3, start), mvarSizes(ColumnIndex - 1)) = IIf(mvarDataFormats(ColumnIndex - 1) = prnNumber, "R$ " + theText, theText)
End Sub

Public Sub DefineColumns(Columns As Integer, Sizes(), prnDataFormats(), prnHeaders())
mvarColumns = Columns
ReDim mvarSizes(Columns)
ReDim mvarSums(Columns - 1)
ReDim mvarDataFormats(Columns - 1)
ReDim mvarColumnHeader(Columns - 1)
mvarAllow = True
mvarSizes = Sizes
For i% = 0 To Columns - 1
    mvarColumnBlankText = mvarColumnBlankText + String(mvarSizes(i%), " ")
    mvarSums(i%) = "0"
    mvarDataFormats(i%) = prnDataFormats(i%)
    mvarColumnHeader(i%) = prnHeaders(i%)
Next i%
End Sub

Public Property Let PrintingMode(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrintingMode = 5
    mvarPrintingMode = vData
End Property


Public Property Get PrintingMode() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrintingMode
    PrintingMode = mvarPrintingMode
End Property
Public Property Let FormCaption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FormCaption = 5
    mvarFormCaption = vData
    frmExibicao.Caption = mvarFormCaption
End Property
Public Property Get FormCaption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FormCaption
    FormCaption = mvarFormCaption
    
End Property
Public Sub SendToPrinter(Mode As Long)

End Sub

Public Property Let PagesNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PagesNumber = 5
    mvarPagesNumber = vData
End Property
Public Property Get PagesNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PagesNumber
    PagesNumber = mvarPagesNumber
End Property
Public Sub InsertBlankLine()
frmExibicao.cmpDados.Text = frmExibicao.cmpDados.Text + vbNewLine + vbNewLine
End Sub
Public Sub WriteText(Texto As String, Optional NovaLinha As Boolean = False)
frmExibicao.cmpDados.Text = frmExibicao.cmpDados.Text + Texto + IIf(NovaLinha, vbNewLine, "")
End Sub
Public Sub ViewOnScreen()
frmExibicao.Show 1
End Sub
Public Sub RowReady()
WriteText mvarColumnBlankText, True
mvarColumnBlankText = String(Len(mvarColumnBlankText), " ")
End Sub
Public Sub WriteHeader()
isHeader = True
DrawLine
For i% = 1 To 4
    AddTextToColumn i%, CStr(mvarColumnHeader(i% - 1)), prnLeft, True
Next i%
RowReady
DrawLine
isHeader = False
End Sub
Public Sub WriteSubTotal(ColumnIndex)
mvarAllow = False
AddTextToColumn (ColumnIndex), CStr(mvarSums(ColumnIndex - 1)), prnRight
RowReady
mvarAllow = True
End Sub
