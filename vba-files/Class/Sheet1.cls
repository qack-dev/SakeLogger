VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    Call M_SheetUtils.UpdateSummarySheet
    Call M_SheetUtils.CreateOrUpdateGraph
    Call M_SheetUtils.AddTotalFields
End Sub
