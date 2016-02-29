Attribute VB_Name = "modPublicVar"

Option Explicit

'Public InvalidDB                    As Boolean
Public CurrUser                     As USER_INFO
Public DBPath                       As String
Public Enc                          As New clsBlowfish
Public CurrBiz                      As BUSINESS_INFO

Public CN                           As New Connection

Public ReportType                   As String
