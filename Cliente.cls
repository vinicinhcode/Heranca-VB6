VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oEndereco As Endereco

Public Sub class_Initialize()
    Set oEndereco = New Endereco
End Sub

Public Sub class_terminate()
     Set oEndereco = Nothing
End Sub

Public Property Get Endereco() As String
    Endereco = oEndereco.Endereco
End Property

Public Property Let Endereco(novoValor As String)
    oEndereco.Endereco = novoValor
End Property

Public Property Get Cep() As String
    Cep = oEndereco.Cep
End Property

Public Property Let Cep(novoValor As String)
    oEndereco.Cep = novoValor
End Property

Public Property Let Pais(novoValor As String)
    oEndereco.Pais = novoValor
End Property

Public Property Get Pais() As String
    Pais = oEndereco.Pais
End Property
