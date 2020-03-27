VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public masterRng As Range
Public wiersz As Long
Public qty As Long

Public yyyycw As Long
Public monday As Date
Public CodEntrega As String
Public dateAfterOffset As Date

Public delDate As Date
Public delTime As Date

Public puDate As Date
Public puTime As Date

Public tmc As String

Public pn As String
Public cofor As String


Public sourceRow As Long
Public sourceRng As Range

Public found As Boolean

