Option Explicit
'# Print Functionality.
'
'## class cPrint:
' From the MVC architecture perspective, this is the main `Controller` for the printing
' functionality. (The `View` is the printed paper itself, and the `Model` is the class
' cAggregatedOrder – which loads, processes, and provides data for the printing process).
'
' From the internal structure point of view of the printing functionality, this class is
' the `Context` of the `strategy` pattern.
' Reference: https://refactoring.guru/design-patterns/strategy
'
' From the perspective of this feature's relation to the rest of the system, this class
' acts as a `facade` that handles all printing requests.
'
' Responsibility: receive the various printing requests and delegate to the internal
' mechanisms that fulfill the request.
'
'## Properties
Private PrintStrategy     As iPrintStrategy
Private Data              As cAggregatedOrder
Private Layout            As cPrintLayout

'## Method exposed as the main entry point for print requests.
Public Sub Print( _
                 ByRef pType As CouponType_Enu, _
                 ByRef pOrder As Long, _
                 Optional ByRef pFormat As PrintFormat_Enu = formatCoupon, _
                 Optional ByRef pCopies As Integer = 1 _
                 )
   On Error GoTo ERR_cPrint
   '--- Validate printer availability before doing anything else
   If PrinterError(Printer) Then Exit Sub

   Dim i As Integer
   '--- Instantiate data object and load its contents. This is the data source for
   ' generating the printed content.
   Set Data = New cAggregatedOrder
   Data.Ndoc = pOrder ' populate
   
   '--- Instantiate a layout object, which is a component library responsible for formatting
   ' each printed line.
   Set Layout = New cPrintLayout
   
   '--- Define the printing strategy.
   Select Case pType
      Case couponTypeSale
      Set PrintStrategy = New cPrintStrategy_Sale
      
      Case couponTypeExchange
      Set PrintStrategy = New cPrintStrategy_Exchange
      
      Case couponTypePromotional
      Set PrintStrategy = New cPrintStrategy_Promotional
      
      Case couponTypeRemainingInstallments
      Set PrintStrategy = New cPrintStrategy_Remaining
      
      Case couponTypeRefinanced
      Set PrintStrategy = New cPrintStrategy_Refinanced
      
      Case Else
      GoTo ERR_cPrint
   End Select
   
   '--- Let the system breathe while printing is being processed
   DoEvents
   
   '--- Request the printing operation.
   ' The `.execute` method receives the dependency injection: `data`, `layout`,
   ' and the print `format` parameter.
   For i = 1 To pCopies
      PrintStrategy.execute Data, pFormat, Layout
   Next i
   
   '--- Free memory
   Set PrintStrategy = Nothing
   Set Layout = Nothing
   Set Data = Nothing
   
   Exit Sub
ERR_cPrint:
   MsgBox "cPrint.Print" & vbCrLf & Err.Description, , Err.Source
End Sub
