Private Sub Workbook_BeforeClose(Cancel As Boolean)
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("&;MyFunction").Delete
      On Error GoTo 0
   End With
End Sub
Private Sub Workbook_Open()
   Dim objPopUp As CommandBarPopup
   Dim objBtn As CommandBarButton
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("CORWELL").Delete
      On Error GoTo 0
      Set objPopUp = .Controls.Add( _
         Type:=msoControlPopup, _
         before:=.Controls.Count, _
         temporary:=True)
   End With
   
   objPopUp.Caption = "&CORWELL"
   
   Set objBtn = objPopUp.Controls.Add
   
   With objBtn
      .Caption = "Képadatbázis másoló"
      .OnAction = "CallKepMasolo"
      .Style = msoButtonCaption
      .FaceId = 331
   End With

End Sub