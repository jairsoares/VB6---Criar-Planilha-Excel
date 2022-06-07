VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Caption         =   "commanfd"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
GeraExcel True
End Sub


Private Sub GeraExcel(geraMeses As Boolean)

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object


    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)

oSheet.Range("D2").Value = "Demonstrativo Anual"
oSheet.Range("D2").Font.Bold = True
oSheet.Range("D2").Font.Name = "Calibri"
oSheet.Range("D2").Font.Size = 20
oSheet.Range("D3").Value = "Pagamentos Efetuados para os Corretores de Imóveis"
oSheet.Range("D4").Value = "Noêmia Margareth Rokebargk Bezerra - CRECI 18.258"
oSheet.Range("D5").Value = "Adir Arocha Pedroso - CRECI 31.760"

' Dados do Proprietario
'-----------------------------------------------
oSheet.Range("A6").Value = "Proprietário:"
oSheet.Range("B6").Value = "<CODIGO>"
oSheet.Range("C6").Value = "<NOME DO PROPRIETARIO>"
oSheet.Range("G6").Value = "CPF:"
oSheet.Range("H6").Value = "<CPF>"

' Dados do Imovel
'-----------------------------------------------
oSheet.Range("A7").Value = "Imóvel:"
oSheet.Range("B7").Value = "<CODIGO>"
oSheet.Range("C7").Value = "<ENDERECO DO IMOVEL>"

'Dados do Locatário
'-----------------------------------------------
oSheet.Range("A8").Value = "Locatário:"
oSheet.Range("B8").Value = "<CODIGO>"
oSheet.Range("C8").Value = "<NOME DO INQUILINO>"
oSheet.Range("G8").Value = "CPF:"
oSheet.Range("H8").Value = "<CPF>"

' Dados do Contrato
'----------------------------------------------
oSheet.Range("F9").Value = "Data Contrato:"
oSheet.Range("H9").Value = "<Data>"



' Formatacao
'
'   VerticalAlignment = 1 = Topo
'   VerticalAlignment = 2 = Centro
'   VerticalAlignment = 3 = Embaixo (Padrao)
'
'   HorizontalAlignment =  2 = Esquerda
'   HorizontalAlignment =  3 = Centro
'   HorizontalAlignment =  4 = Direita
'---------------------------------------------


oSheet.Columns("A:A").ColumnWidth = 14.3
oSheet.Columns("B:B").ColumnWidth = 4.43
oSheet.Columns("C:C").ColumnWidth = 12
oSheet.Columns("D:D").ColumnWidth = 15.71
oSheet.Columns("E:E").ColumnWidth = 15
oSheet.Columns("G:G").ColumnWidth = 7

oSheet.Range("A6").HorizontalAlignment = 4
oSheet.Range("A7").HorizontalAlignment = 4
oSheet.Range("A8").HorizontalAlignment = 4
oSheet.Range("G6").HorizontalAlignment = 4
oSheet.Range("G8").HorizontalAlignment = 4

oSheet.Range("F9:G9").Merge
oSheet.Range("F9").HorizontalAlignment = 4


' Criação do Cabeçalho dos Alugueis
'---------------------------------------------

oSheet.Range("B10:C10").Merge
oSheet.Range("F10:G10").Merge
oSheet.Range("H10:I10").Merge
oSheet.Range("A10").Value = "Mês Referência"
oSheet.Range("B10").Value = "Valor Aluguel R$"
oSheet.Range("D10").Value = "Administração R$"
oSheet.Range("E10").Value = "Condomínio R$"
oSheet.Range("F10").Value = "Taxa Contrato R$"
oSheet.Range("H10").Value = "Total R$"

oSheet.Range("A10").Font.Bold = True
oSheet.Range("B10").Font.Bold = True
oSheet.Range("D10").Font.Bold = True
oSheet.Range("E10").Font.Bold = True
oSheet.Range("F10").Font.Bold = True
oSheet.Range("H10").Font.Bold = True


oSheet.Range("A10").HorizontalAlignment = 3
oSheet.Range("B10").HorizontalAlignment = 3
oSheet.Range("D10").HorizontalAlignment = 3
oSheet.Range("E10").HorizontalAlignment = 3
oSheet.Range("F10").HorizontalAlignment = 3
oSheet.Range("H10").HorizontalAlignment = 3

Dim i As Integer
Dim nMes As Integer
Dim mesRef As String

nMes = 1
For i = 11 To 22
    oSheet.Range("B" & Trim(Str(i)) & ":C" & Trim(Str(i))).Merge
    oSheet.Range("F" & Trim(Str(i)) & ":G" & Trim(Str(i))).Merge
    oSheet.Range("H" & Trim(Str(i)) & ":I" & Trim(Str(i))).Merge
    If geraMeses Then
       Select Case nMes
              Case 1
                   mesRef = "Janeiro/2020"
              Case 2
                   mesRef = "Fevereiro/2020"
              Case 3
                   mesRef = "Março/2020"
              Case 4
                   mesRef = "Abril/2020"
              Case 5
                   mesRef = "Maio/2020"
              Case 6
                   mesRef = "Junho/2020"
              Case 7
                   mesRef = "Julho/2020"
              Case 8
                   mesRef = "Agosto/2020"
              Case 9
                   mesRef = "Setembro/2020"
              Case 10
                   mesRef = "Outubro/2020"
              Case 11
                   mesRef = "Novembro/2020"
              Case 12
                   mesRef = "Dezembro/2020"
       End Select
       oSheet.Range("A" & Trim(Str(i))).Value = mesRef
    End If
    oSheet.Range("A" & Trim(Str(i))).HorizontalAlignment = 3
    oSheet.Range("B" & Trim(Str(i))).HorizontalAlignment = 3
    oSheet.Range("D" & Trim(Str(i))).HorizontalAlignment = 3
    oSheet.Range("E" & Trim(Str(i))).HorizontalAlignment = 3
    oSheet.Range("F" & Trim(Str(i))).HorizontalAlignment = 3
    oSheet.Range("H" & Trim(Str(i))).HorizontalAlignment = 3
    nMes = nMes + 1
    
Next


CDLG.DialogTitle = "Onde Salvar a Planilha ?"
    CDLG.DefaultExt = ".mdb"
    CDLG.Filter = "Microsoft Excel|*.xls"
    CDLG.InitDir = App.Path
    CDLG.ShowSave


'Save the Workbook and Quit Excel
oBook.SaveAs CDLG.FileName
oExcel.Quit
MsgBox "Feito."

End Sub
