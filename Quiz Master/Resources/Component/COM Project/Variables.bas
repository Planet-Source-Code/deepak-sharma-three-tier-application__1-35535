Attribute VB_Name = "Variables"
'********* ADO VARIABLE ***********'

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public AddScore As New ADODB.Recordset
Public Find As New ADODB.Recordset



'********* INTEGER VARIABLE ***********'

Public Quest_No As Integer
Public I As Integer
Public num As Integer
Public Incre As Integer
Public J As Integer
Public Total As Integer
Public k As Integer
Public B As Integer
Public D As Integer
Public Correct_Answer As Integer

'********* STRING VARIABLE ***********'

Public Total_Quest As String
Public Quest_Already_Came As String
Public Path As String
Public Comment As String * 100
Public Remark As String * 100

'********* BOOLEAN VARIABLE ***********'

Public AddQuestion As Boolean


