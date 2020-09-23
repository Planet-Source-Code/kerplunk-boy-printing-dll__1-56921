VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim prn As New Printing
Set prn = prn

prn.WriteText "Testando", False
prn.InsertBlankLine

prn.WriteText "Testando linha 3"
prn.InsertBlankLine

For i% = 1 To 118
    prn.WriteText "A", False
Next i%
prn.InsertBlankLine

prn.FormCaption = "Testando caption"

Dim tamanhos(0 To 3)
Dim formatos(0 To 3)
Dim cabecalhodelinha(0 To 3)

tamanhos(0) = 6
tamanhos(1) = 70
tamanhos(2) = 10
tamanhos(3) = 11

formatos(0) = 3
formatos(1) = 1
formatos(2) = 4
formatos(3) = 2

cabecalhodelinha(0) = "Cód"
cabecalhodelinha(1) = "Produto"
cabecalhodelinha(2) = "Valor"
cabecalhodelinha(3) = "Data"


prn.DefineColumns 4, tamanhos, formatos, cabecalhodelinha

prn.WriteHeader

prn.AddTextToColumn 1, "1234", prnRight
prn.AddTextToColumn 2, "Camisa flanela estampada sem mangas e com botoes", prnCenter
prn.AddTextToColumn 3, "10,00", prnRight
prn.AddTextToColumn 4, CStr(Date), prnRight
prn.RowReady

prn.WriteSubTotal 3

prn.AddTextToColumn 1, "3654", prnRight
prn.AddTextToColumn 2, "Meia-calça nylon Tryfill", prnCenter
prn.AddTextToColumn 3, "3,65", prnRight
prn.AddTextToColumn 4, CStr(Date + 9), prnRight
prn.RowReady

prn.WriteSubTotal 3

prn.ViewOnScreen
End Sub
