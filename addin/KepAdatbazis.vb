
'Subroutine amely a mégse gombra kattintva
'bezárja a userformot.
Private Sub cancelButton_Click()
    Unload Me
End Sub

'Subroutine amely leköveti a cikkszám
'spin gombok változását
Private Sub cikkszamSpinbutton_Change()
    cikkszamTextbox.Value = cikkszamSpinbutton.Value
End Sub

'Subroutine amely az 'export mappa helye' gombra
'hívódik meg. Figyelmeztetés után megjeleníti a Windows
'beépített tallózó menüjét.
Private Sub endPathCommandbutton_Click()
    MsgBox "Az export mappa beállításainál figyelj arra, hogy a program felülír minden fájlt," & _
            "ha a munkalapon szerepel az adott névvel. Ez az érték ne legyen a szerverünk!", _
            vbOKOnly
    
    Dim diaFolder As FileDialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
        If diaFolder.SelectedItems.Count > 0 Then
            KepAdatbazis.endPathTextbox.Value = diaFolder.SelectedItems(1)
        End If
End Sub

'Subroutine amely leköveti a hiányzó kép
'spin gombok változását
Private Sub hianyzoKepSpinbutton_Change()
    hianyzoKepTextbox.Value = hianyzoKepSpinbutton.Value
End Sub

'Subroutine amely leköveti a márkához tartozó
'spin gombok változását
Private Sub markaSpinbutton_Change()
    markaTextbox.Value = markaSpinbutton.Value
End Sub

'Mappa létezését ellenőrző function. 
'Meghívva és átadva a mappa elérési útvonalát, visszaadja, hogy
'Az adott mappa létezik-e vagy nem.
Function DirectoryExists(Directory As String) As Boolean
    DirectoryExists = False
    If Not Dir(Directory, vbDirectory) = "" Then
        If GetAttr(Directory) = vbDirectory Then
            DirectoryExists = True
        End If
    End If
End Function

'Ékezeteket megszűntetű function.
'Meghívva és átadva a Stringet, visszaadja azt olyan formában, hogy 
'Ne legyenek benne ékezetek.
Function StripAccents(folderName As String) As String
    Dim A As String * 1
    Dim B As String * 1
    Dim i As Integer

    Const AccChars = "ÁÉÍÓÖŐÜÚŰáéíóöőüúű"
    Const RegChars = "AEIOOOUUUaeiooouuu"

    For i = 1 To 18
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        folderName = Replace(folderName, A, B)
    Next i

    StripAccents = Trim(folderName)
End Function

'A userform 'interaktivitását' befolyásoló function.
'Meghívva és átadva, hogy milyen állapotba állítsa át 
'A userformot, megváltoztatja annak gombjainak, szövegmezőinek kattinthatóságát
'és színeit.
Function userFormTrigger(trigger As Boolean)
    
    Me.originPathTextbox.Enabled = trigger
    Me.endPathTextbox.Enabled = trigger
    Me.originPathOptionCheckbox.Enabled = trigger
    Me.endPathOptionCheckbox.Enabled = trigger
    Me.cikkszamTextbox.Enabled = trigger
    Me.markaTextbox.Enabled = trigger
    Me.hianyzoKepTextbox.Enabled = trigger
    Me.originPathCommandbutton.Enabled = trigger
    Me.endPathCommandbutton.Enabled = trigger
    Me.cikkszamSpinbutton.Enabled = trigger
    Me.markaSpinbutton.Enabled = trigger
    Me.hianyzoKepSpinbutton.Enabled = trigger
    Me.resetButton.Enabled = trigger
    Me.okButton.Enabled = trigger
    
    If trigger Then
    
    Me.originPathTextbox.BackColor = &H80000005
    Me.endPathTextbox.BackColor = &H80000005
    Me.originPathOptionCheckbox.BackColor = &H80000005
    Me.endPathOptionCheckbox.BackColor = &H80000005
    Me.hianyzoKepTextbox.BackColor = &H80000005
    Me.markaTextbox.BackColor = &H80000005
    Me.cikkszamTextbox.BackColor = &H80000005
    
    Else
    
    Me.originPathTextbox.BackColor = &H80000016
    Me.endPathTextbox.BackColor = &H80000016
    Me.originPathOptionCheckbox.BackColor = &H80000016
    Me.endPathOptionCheckbox.BackColor = &H80000016
    Me.hianyzoKepTextbox.BackColor = &H80000016
    Me.markaTextbox.BackColor = &H80000016
    Me.cikkszamTextbox.BackColor = &H80000016
    
    End If
End Function

'Subroutine amelyet az OK gombra kattintás hív meg.
'Ez a subroutine tartalmazza a bővítmény fő logikáját.
Private Sub okButton_Click()

'Változók deklarálása
Dim originPath As String
Dim originPath_Completed As String
Dim endPath As String
Dim endPath_Completed As String
Dim folderName As String
Dim fileName As String
Dim strCheck As String

Dim originPathOption As Boolean
Dim endPathOption As Boolean

Dim lastRow As Long
Dim cikkszamColumn As Long
Dim markaColumn As Long
Dim hianyzoKepColumn As Long
Dim maxProgressBarWidth As Long
Dim life As Long
Dim i As Long
Dim n As Long
Dim x As Long

Dim coll As New Collection

Dim pbIncrement As Single

'Progressbar alapbeállítása
maxProgressBarWidth = 312
Me.pBarLabel.Height = 24
Me.pBarLabel.Width = 0
Me.pBarLabel.ForeColor = &H8000000D
Me.pBarLabel.BackColor = &H8000000D

'Hibakeresés a következő blokkra
On Error GoTo err_handler2 

    'Userformból átvezetjük a változókat
originPath = originPathTextbox.Value
endPath = endPathTextbox.Value
originPathOption = originPathOptionCheckbox.Value
endPathOption = endPathOptionCheckbox.Value
cikkszamColumn = cikkszamTextbox.Value
markaColumn = markaTextbox.Value
hianyzoKepColumn = hianyzoKepTextbox.Value

'Hibakeresés alapértelmezettre állítása
On Error GoTo 0 

'Az elérési útvonal javítása és hibakeresés üres útvonalra
If Right(originPath, 1) <> "\" Then
    originPath = originPath & "\"
End If
  
If Right(endPath, 1) <> "\" Then
    endPath = endPath & "\"
End If

If Dir(originPath, vbDirectory) = "" Then
    GoTo err_handler4
End If

If Dir(endPath, vbDirectory) = "" Then
    GoTo err_handler4
End If

'Aktív munkalapot a jelenlegire állítjuk
Set ws = ActiveSheet

With ws
    'Hibakeresés a következő sorra
    On Error GoTo err_handler1
        'Utolsó sor sorszámának meghatározása
        lastRow = .Cells(.Rows.Count, cikkszamColumn).End(xlUp).Row
    On Error GoTo 0 
    
    'Progressbar növekedési egységének kiszámolása
    pbIncrement = maxProgressBarWidth / lastRow 

    'Interaktivitás kikapcsolása               
    userFormTrigger (False) 
  
    For i = 2 To lastRow
        'Progressbar növelése
        Me.pBarLabel.Width = Format(Me.pBarLabel.Width + pbIncrement, "#.##")
        'Progressbar állapotjelző frissítése
        Me.pBarFrame.Caption = "Ellenőrzés: " & CStr(i) & " ennyiből: " & CStr(lastRow)
        
        'Átadjuk az execution jogot az operációs rendszernek. Enélkül a felhasználó
        'Azt a benyomást kapja, hogy 'kifagyott' az Excel
        DoEvents

        'Gyűjtő létrehozása    
        Set coll = New Collection 
        
        'Hibakeresés a következő sorra
        On Error GoTo err_handler3
            'Mappanév létrehozása 
            folderName = .Cells(i, markaColumn).Value 
        On Error GoTo 0
        
        'Mappanévben ékezetes betűk javítása
        folderName = StripAccents(folderName)
        
        'Fájlnév létrehozása
        fileName = .Cells(i, cikkszamColumn).Value
        n = 1

        'Az élet határozza meg, hogy hány kép "űr" lehet 2 kép közt mielőtt továbblépünk
        life = 2
                
        If originPathOption Then
            strCheck = Dir(originPath & folderName & "\" & fileName & ".*")

            If Len(strCheck) = 0 Then
                On Error GoTo err_handler5
                .Cells(i, hianyzoKepColumn).Value = "Nem találtam képet"
                On Error GoTo 0
            Else
'==== Első logikai kapu ====
                If endPathOption Then
                    
                    originPath_Completed = originPath & folderName & "\" & strCheck
                    If DirectoryExists(endPath & folderName) Then
                        endPath_Completed = endPath & folderName & "\" & strCheck
                    Else
                        MkDir (endPath & folderName)
                        endPath_Completed = endPath & folderName & "\" & strCheck
                    End If
                    
                    FileCopy originPath_Completed, endPath_Completed
                    
                    Do
                        strCheck = Dir(originPath & folderName & "\" & fileName & "_" & n & ".*")
                        
                        If Len(strCheck) > 0 Then
                            n = n + 1
                            coll.Add strCheck
                        Else
                            n = n + 1
                            life = life - 1
                        End If
                    Loop Until life = 0
                    
                    If coll.Count > 0 Then
                    
                        For x = 1 To coll.Count
                            originPath_Completed = originPath & folderName & "\" & coll.Item(x)
                            If DirectoryExists(endPath & folderName) Then
                                endPath_Completed = endPath & folderName & "\" & coll.Item(x)
                            Else
                                MkDir (endPath & folderName)
                                endPath_Completed = endPath & folderName & "\" & coll.Item(x)
                            End If
                            
                            FileCopy originPath_Completed, endPath_Completed
                        Next x
                    End If
'==== Első logikai kapu =====
'==== Második logikai kapu=====
                Else
                
                    originPath_Completed = originPath & folderName & "\" & strCheck
                    endPath_Completed = endPath & strCheck
                    FileCopy originPath_Completed, endPath_Completed
                                  
                    Do
                        strCheck = Dir(originPath & folderName & "\" & fileName & "_" & n & ".*")
                    
                        If Len(strCheck) > 0 Then
                            n = n + 1
                            coll.Add strCheck
                        Else
                            n = n + 1
                            life = life - 1
                        End If
                
                    Loop Until life = 0
                    
                    If coll.Count > 0 Then
                    
                        For x = 1 To coll.Count
                            originPath_Completed = originPath & folderName & "\" & coll.Item(x)
                            endPath_Completed = endPath & coll.Item(x)
                            FileCopy originPath_Completed, endPath_Completed
                        Next x
                    End If
                End If
            End If
'==== Második logikai kapu =====
        Else
            strCheck = Dir(originPath & fileName & ".*")
            
            If Len(strCheck) = 0 Then
                On Error GoTo err_handler5
                .Cells(i, hianyzoKepColumn).Value = "Nem találtam képet"
                On Error GoTo 0
            Else
'==== Harmadik logikai kapu ====
                If endPathOption Then
                    originPath_Completed = originPath & strCheck
                    If DirectoryExists(endPath & folderName) Then
                        endPath_Completed = endPath & folderName & "\" & strCheck
                    Else
                        MkDir (endPath & folderName)
                        endPath_Completed = endPath & folderName & "\" & strCheck
                    End If
                    
                    FileCopy originPath_Completed, endPath_Completed
                                        
                    Do
                        strCheck = Dir(originPath & fileName & "_" & n & ".*")
                        
                        If Len(strCheck) > 0 Then
                            n = n + 1
                            coll.Add strCheck
                        Else
                            n = n + 1
                            life = life - 1
                        End If
                    Loop Until life = 0
                    
                    If coll.Count > 0 Then
                    
                        For x = 1 To coll.Count
                            originPath_Completed = originPath & coll.Item(x)
                            If DirectoryExists(endPath & folderName) Then
                                endPath_Completed = endPath & folderName & "\" & coll.Item(x)
                            Else
                                MkDir (endPath & folderName)
                                endPath_Completed = endPath & folderName & "\" & coll.Item(x)
                            End If
                            
                            FileCopy originPath_Completed, endPath_Completed
                        Next x
                    End If
'==== Harmadik logikai kapu ====
'==== Negyedik logikai kapu ====
                Else
                    originPath_Completed = originPath & strCheck
                    endPath_Completed = endPath & strCheck
                    FileCopy originPath_Completed, endPath_Completed
                    
                    Do
                        strCheck = Dir(originPath & fileName & "_" & n & ".*")
                        
                        If Len(strCheck) > 0 Then
                            n = n + 1
                            coll.Add strCheck
                                                                            
                        Else
                            n = n + 1
                            life = life - 1
                        End If
                    Loop Until life = 0
                    
                    If coll.Count > 0 Then
                        
                        For x = 1 To coll.Count
                            originPath_Completed = originPath & coll.Item(x)
                            endPath_Completed = endPath & coll.Item(x)
                            FileCopy originPath_Completed, endPath_Completed
                            
                        Next x
                    End If
                End If
            End If
        End If
'==== Negyedik logikai kapu ====
    'Gyűjtő alapértelmezettre állítása
    Set coll = Nothing
    Next i
End With

'Interaktivitás bekapcsolása
userFormTrigger (True) 

Done:
    Me.pBarFrame.Caption = "Kész!"
    Me.pBarLabel.Width = 0
    Exit Sub

err_handler1:
    MsgBox ("Objektum hiba: Hiba a cikkszám oszlopban vagy annak beállításában.")
    userFormTrigger (True) 'Interaktivitás bekapcsolása
    Me.pBarFrame.Caption = "Állapot - Várakozás"
    Exit Sub

err_handler2:
    MsgBox "Adatfeljegyzési típushiba: nem megfelelő változó típus." & vbNewLine & "Pl.: Az egyik bemeneti mezőben szám helyett szöveg lett megadva."
    userFormTrigger (True) 'Interaktivitás bekapcsolása
    Me.pBarFrame.Caption = "Állapot - Várakozás"
    Exit Sub

err_handler3:
    MsgBox ("Objektum hiba: Hiba a márka oszlopban vagy annak beállításában.")
    userFormTrigger (True) 'Interaktivitás bekapcsolása
    Me.pBarFrame.Caption = "Állapot - Várakozás"
    Exit Sub

err_handler4:
    MsgBox ("Elérési út hiba: Nem létező elérési út lett megadva.")
    userFormTrigger (True) 'Interaktivitás bekapcsolása
    Me.pBarFrame.Caption = "Állapot - Várakozás"
    Exit Sub

err_handler5:
    MsgBox ("Objektum hiba: Hiba a hiányzó kép oszlopban vagy annak beállításában.")
    userFormTrigger (True) 'Interaktivitás bekapcsolása
    Me.pBarFrame.Caption = "Állapot - Várakozás"
    Exit Sub
End Sub

'Subroutine amely a 'webes képek helye' gombra
'hívódik meg. Megjeleníti a Windows beépített tallózó menüjét.
Private Sub originPathCommandbutton_Click()

    Dim diaFolder As FileDialog

    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
       If diaFolder.SelectedItems.Count > 0 Then
           KepAdatbazis.originPathTextbox.Value = diaFolder.SelectedItems(1)
       End If
End Sub

'Subroutine amely az alapértelmezett gombra kattintva hívódik meg.
Private Sub resetButton_Click()

    Call UserForm_Initialize

End Sub

'Subroutine amely a userform betöltéséhez kapcsolódik
Private Sub UserForm_Initialize()

    'Webes képek helye útvonal default beállítása
    originPathTextbox.Value = "[Biztonságtechnikai okok végett eltávolítva GITHUB feltöltéskor]"
    'Export mappa útvonal default beállítása
    endPathTextbox.Value = ""
    'Webes képek mappastruktúra opció default beállítása
    originPathOptionCheckbox.Value = True
    'Export mappa mappastruktúra opció default beállítása
    endPathOptionCheckbox.Value = True
    'Cikkszám default beállítása
    cikkszamTextbox.Value = 1
    'Márka default beállítása
    markaTextbox.Value = 2
    'Hiba default beállítása
    hianyzoKepTextbox.Value = 3
    'Spinbuttonok default beállítása
    cikkszamSpinbutton.Value = cikkszamTextbox.Value
    hianyzoKepSpinbutton.Value = hianyzoKepTextbox.Value
    markaSpinbutton.Value = markaTextbox.Value
    'Progressbar beállítások
    Me.pBarLabel.Height = 24
    Me.pBarLabel.Width = 0
    Me.pBarLabel.ForeColor = &H8000000D
    Me.pBarLabel.BackColor = &H8000000D
    Me.pBarFrame.Caption = "Állapot - Várakozás"
End Sub