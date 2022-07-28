VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim Cliente As Cliente
    
    Set Cliente = New Cliente
    
    With Cliente
    
        .Endereco = "Rua Elaney, 70 - Jardim dos comerciarios "
        .Cep = "31.650-040"
        .Pais = "Brasil"
        
        MsgBox .Endereco & "(" & .Cep & ") " & .Pais
    
    End With
    
    

End Sub
