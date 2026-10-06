VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerPrototype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Interface Class LoggerPrototype
Implements Logger


' Returns a new instance of this class with exact same configuration as this object.
Public Function clone() As LoggerPrototype
End Function


' Depending on the implementation, the given loggerName could be used to determine location where the log messages are written to
' or can be used inside the messages.
Public Sub setName(loggerName As String)
End Sub


Private Sub Logger_fatal(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub

Private Sub Logger_info(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub

Private Sub Logger_warn(message As String, Optional msg2 As String, Optional msg3 As String)
End Sub

Private Function Logger_whereIsMyLog() As String
End Function



Private Sub Logger_warn(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.warn, message, msg2, msg3
End Sub


Private Sub Logger_fatal(message As String, Optional msg2 As String, Optional msg3 As String)
    logMessage LogFactory.fatal, message, msg2, msg3
End Sub


Public Sub LoggerPrototype_setName(loggerName As String)
    name = loggerName
End Sub


Private Function LoggerPrototype_clone() As LoggerPrototype
    Dim clone As ConsoleLogger
    Set clone = New ConsoleLogger
    ' the clone's name will be overwritten, so no need to: clone.LoggerPrototype_setName name
    Set LoggerPrototype_clone = clone
End Function


' The actual implementation of the 'info' , 'warn' and 'fatal' subs.
Private Sub logMessage(status As String, message As String, Optional msg2 As String, Optional msg3 As String)
    Dim formatted As String
    formatted = LogFactory.formatLogMessage(status & "|" & name, message, msg2, msg3)
    Debug.Print formatted
End Sub
