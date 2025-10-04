VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceCoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ticker As String
Public name As String
Public isFiat As Boolean
Public bases As Long
Public quotes As Long
Public precision As Long
Public markets As New Dictionary

