VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class ConsoleLogger
Implements Logger
Implements LoggerPrototype
' We are required to implement both interfaces, even if LoggerPrototype also implements Logger.

' This logger appends to the 'immediate window' in the Excel VBA IDE.


Private name As String ' The name of this logger


Private Function Logger_whereIsMyLog() As String
    Logger_whereIsMyLog = "Log output is printed to the immediate window."
End Function


Private Sub Logger_info(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.info, message, msg2, msg3
End Sub



