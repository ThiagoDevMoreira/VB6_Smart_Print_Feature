Option Explicit

' -- API DECLARATIONS --
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" ( _
    ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" ( _
    ByVal hPrinter As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" ( _
    ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, _
    ByVal cbBuf As Long, pcbNeeded As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

' -- CONSTANTS --
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
' (Add others as needed)

' -- STRUCT --
Public Type PRINTER_INFO_2
    pServerName As Long
    pPrinterName As Long
    pShareName As Long
    pPortName As Long
    pDriverName As Long
    pComment As Long
    pLocation As Long
    pDevMode As Long
    pSepFile As Long
    pPrintProcessor As Long
    pDatatype As Long
    pParameters As Long
    pSecurityDescriptor As Long
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Public Function PrinterHasError(ByRef pPrinter As Printer) As Boolean
   Dim hPrinter As Long
   Dim bytesNeeded As Long
   Dim pi2() As Byte
   Dim status As Long
   Dim result As Long
   Dim ErrorMessage As String

   On Error GoTo HandleError
   PrinterHasError = False
   ErrorMessage = ""

   '--- Try to open the printer (including network printers)
   result = OpenPrinter(pPrinter.DeviceName, hPrinter, 0)
   If result = 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Could not access printer: " & pPrinter.DeviceName & vbCrLf & _
         "  - Check if it is turned on, properly installed, and connected to the network." & _
         vbCrLf & vbCrLf
      Exit Function ' Cannot proceed if printer could not be opened
   End If

   '--- Get required buffer size
   GetPrinter hPrinter, 2, ByVal 0, 0, bytesNeeded
   If bytesNeeded = 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Failed to query the printer." & vbCrLf & _
         "  - Possible spooler or driver issue." & vbCrLf & vbCrLf
   End If

   '--- Allocate buffer and get data
   ReDim pi2(0 To bytesNeeded - 1) As Byte
   result = GetPrinter(hPrinter, 2, pi2(0), bytesNeeded, bytesNeeded)
   If result = 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Could not retrieve printer status." & vbCrLf & _
         "  - Check the driver and restart the print spooler." & vbCrLf & vbCrLf
   End If

   '--- Extract status from structure
   Dim pi2Struct As PRINTER_INFO_2
   CopyMemory pi2Struct, pi2(0), Len(pi2Struct)
   status = pi2Struct.status

   '--- Close printer handle
   ClosePrinter hPrinter

   '--- Evaluate status bits
   If (status And PRINTER_STATUS_OFFLINE) <> 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Printer is offline." & vbCrLf
   End If

   If (status And PRINTER_STATUS_NOT_AVAILABLE) <> 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Printer is not available." & vbCrLf
   End If

   If (status And PRINTER_STATUS_ERROR) <> 0 Then
      PrinterHasError = True
      ErrorMessage = ErrorMessage & "* Printer has an error (e.g. paper, ink, jam)." & vbCrLf
   End If

   '--- Display accumulated error messages
   If PrinterHasError And Len(ErrorMessage) > 0 Then
      MsgBox "Printer issues detected:" & vbCrLf & vbCrLf & ErrorMessage, _
         vbExclamation, "Print Error"
   End If

   Exit Function

HandleError:
   PrinterHasError = True
   MsgBox "Unexpected error while checking the printer. There may be system or " & _
      "communication issues.", vbCritical, "Error"
End Function
