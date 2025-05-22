Option Explicit
Implements iPrintStrategy

' # Class cPrintStrategy_Refinanced (concrete strategy)
Public Sub iPrintStrategy_execute( _
            ByRef pData As cAggregatedOrder, _
            ByRef pFormat As PrintFormat_Enu, _
            ByRef pLayout As cPrintLayout _
            )
   On Error GoTo EH

   Dim Receipts() As ctlReceipt
   Dim Receipt As Variant
   Dim RescheduledReceipts() As ctlReceipt
   Dim i As Byte
   Dim RenegotiatedAmount As Double

   '--- Get receipts
   Receipts = pData.getReceipts

   With Receipt
      i = 0
      For Each Receipt In Receipts
         '--- Collect receipts that belong to the rescheduling
         If Not IsNull(.getField("Reschedule_Date")) Then
            ReDim RescheduledReceipts(i)
            RescheduledReceipts(i) = Receipt
            i = i + 1
         End If
         '--- Sum of renegotiated value
         RenegotiatedAmount = RenegotiatedAmount + .getField("Value")
      Next Receipt
      i = 0
   End With

   With pLayout
      .separator
      .title "RENEGOTIATED ACCOUNT"
      .installments_and_payment_methods pData.getReceipts
      .separator
      .total_renegotiated RenegotiatedAmount
      .separator
      .thank_you_note
      .signature
   End With

   Printer.EndDoc
   Exit Sub

EH:
   MsgBox Err.Description, vbInformation, "cSaleReceiptStrategy.execute"
   Screen.MousePointer = 0
End Sub
