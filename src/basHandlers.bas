Attribute VB_Name = "basHandlers"
Option Compare Database
Option Explicit

'==============================================================
'  دوال معالجة الأحداث (تستدعى من أزرار النماذج بصيغة =دالة())
'  التنقل - الحفظ - الحذف - البحث - التصفية - الفرز
'==============================================================

'================= التنقل بين السجلات ========================

Public Function NavRec(sDir As String)
    On Error GoTo ErrH
    Select Case sDir
        Case "first": DoCmd.GoToRecord , , acFirst
        Case "prev":  DoCmd.GoToRecord , , acPrevious
        Case "next":  DoCmd.GoToRecord , , acNext
        Case "last":  DoCmd.GoToRecord , , acLast
    End Select
    Exit Function
ErrH:
    Beep    ' الوصول إلى بداية أو نهاية السجلات
End Function

'================= إدارة السجل الحالي ========================

Public Function RecNew()
    On Error Resume Next
    DoCmd.GoToRecord , , acNewRec
    Screen.ActiveForm!RefNumber.SetFocus
End Function

Public Function RecSave()
    On Error GoTo ErrH
    If Screen.ActiveForm.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
        MsgBox "تم حفظ السجل بنجاح.", vbInformation, "حفظ"
    Else
        MsgBox "لا توجد تغييرات تحتاج إلى حفظ.", vbInformation, "حفظ"
    End If
    Exit Function
ErrH:
    MsgBox "تعذر الحفظ: " & Err.Description & vbCrLf & _
           "تأكد من إدخال رقم المكاتبة (حقل إلزامي).", vbExclamation, "خطأ في الحفظ"
End Function

Public Function RecDelete()
    On Error GoTo ErrH
    If MsgBox("هل أنت متأكد من حذف هذه المكاتبة نهائياً؟", _
              vbYesNo + vbQuestion + vbDefaultButton2, "تأكيد الحذف") = vbYes Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    Exit Function
ErrH:
    If Err.Number <> 2501 Then _
        MsgBox "لا يوجد سجل يمكن حذفه.", vbExclamation, "حذف"
End Function

Public Function RecUndo()
    On Error Resume Next
    If Screen.ActiveForm.Dirty Then
        Screen.ActiveForm.Undo
    Else
        Beep
    End If
End Function

'================= البحث والتصفية ============================

' بناء شرط البحث الموحد (يشمل الرقم والموضوع والملاحظات والجهة والتصنيف)
Public Function BuildSearchWhere(sText As String) As String
    Dim s As String
    s = Replace(Trim$(sText), "'", "''")
    BuildSearchWhere = _
        "[RefNumber] LIKE '*" & s & "*' " & _
        "OR [Subject] LIKE '*" & s & "*' " & _
        "OR [ReplyRef] LIKE '*" & s & "*' " & _
        "OR [Notes] LIKE '*" & s & "*' " & _
        "OR [EntityID] IN (SELECT EntityID FROM tblEntities WHERE EntityName LIKE '*" & s & "*') " & _
        "OR [CategoryID] IN (SELECT CategoryID FROM tblCategories WHERE CategoryName LIKE '*" & s & "*')"
End Function

Public Function DoSearch()
    On Error Resume Next
    Dim f As Form, s As String
    Set f = Screen.ActiveForm
    s = Trim$(Nz(f!txtSearch, ""))
    If Len(s) = 0 Then
        f.FilterOn = False
        UpdateCounter
        Exit Function
    End If
    f.Filter = BuildSearchWhere(s)
    f.FilterOn = True
    If f.Recordset.RecordCount = 0 Then
        MsgBox "لا توجد نتائج مطابقة لكلمة البحث: " & s, vbInformation, "بحث"
        f.FilterOn = False
    End If
    UpdateCounter
End Function

Public Function ApplyFilters()
    On Error Resume Next
    Dim f As Form, sWhere As String
    Set f = Screen.ActiveForm

    If Not IsNull(f!cboFType) Then _
        sWhere = sWhere & " AND [CorrType]='" & f!cboFType & "'"
    If Not IsNull(f!cboFCat) Then _
        sWhere = sWhere & " AND [CategoryID]=" & f!cboFCat
    If Not IsNull(f!cboFStatus) Then _
        sWhere = sWhere & " AND [Status]='" & f!cboFStatus & "'"

    If Len(sWhere) = 0 Then
        MsgBox "اختر قيمة واحدة على الأقل من قوائم التصفية.", vbInformation, "تصفية"
        Exit Function
    End If

    f.Filter = Mid$(sWhere, 6)   ' إزالة AND الأولى
    f.FilterOn = True
    If f.Recordset.RecordCount = 0 Then _
        MsgBox "لا توجد مكاتبات مطابقة لشروط التصفية.", vbInformation, "تصفية"
    UpdateCounter
End Function

Public Function ClearFilters()
    On Error Resume Next
    Dim f As Form
    Set f = Screen.ActiveForm
    f.FilterOn = False
    f!txtSearch = Null
    f!cboFType = Null
    f!cboFCat = Null
    f!cboFStatus = Null
    UpdateCounter
End Function

'================= الفرز =====================================

' يبدل اتجاه الفرز (تصاعدي/تنازلي) عند كل ضغطة
Public Function SortToggle(sField As String)
    On Error Resume Next
    Dim f As Form
    Set f = Screen.ActiveForm
    If f.OrderBy = "[" & sField & "] DESC" Then
        f.OrderBy = "[" & sField & "]"
    Else
        f.OrderBy = "[" & sField & "] DESC"
    End If
    f.OrderByOn = True
End Function

'================= عداد السجلات ==============================

Public Function UpdateCounter()
    On Error Resume Next
    Dim f As Form, rs As DAO.Recordset, n As Long
    Set f = Screen.ActiveForm
    Set rs = f.RecordsetClone
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        n = rs.RecordCount
    End If
    If f.NewRecord Then
        f!lblCounter.Caption = "سجل جديد (الإجمالي: " & n & ")"
    Else
        f!lblCounter.Caption = "سجل " & f.CurrentRecord & " من " & n
    End If
End Function

'================= فتح النماذج والتقارير =====================

Public Function OpenCorr(sMode As String)
    On Error Resume Next
    Select Case sMode
        Case "new"
            DoCmd.OpenForm "frmCorrespondence", , , , acFormAdd
        Case "in"
            DoCmd.OpenForm "frmCorrespondence", , , "[CorrType]='وارد'"
        Case "out"
            DoCmd.OpenForm "frmCorrespondence", , , "[CorrType]='صادر'"
        Case "late"
            DoCmd.OpenForm "frmCorrespondence", , , "[Status]<>'منجز' AND [DueDate]<Date()"
        Case Else
            DoCmd.OpenForm "frmCorrespondence"
    End Select
End Function

Public Function OpenObj(sKind As String, sName As String)
    On Error GoTo ErrH
    Select Case sKind
        Case "form":   DoCmd.OpenForm sName
        Case "report": DoCmd.OpenReport sName, acViewPreview
        Case "query":  DoCmd.OpenQuery sName
    End Select
    Exit Function
ErrH:
    If Err.Number <> 2501 Then _
        MsgBox "تعذر فتح " & sName & ": " & Err.Description, vbExclamation, "خطأ"
End Function

' بحث سريع من الواجهة الرئيسية
Public Function QuickSearch()
    Dim s As String
    s = InputBox("أدخل كلمة البحث (رقم المكاتبة / الموضوع / الجهة / الملاحظات):", "بحث سريع")
    If Len(Trim$(s)) = 0 Then Exit Function
    DoCmd.OpenForm "frmCorrespondence", , , BuildSearchWhere(s)
    If Screen.ActiveForm.Recordset.RecordCount = 0 Then
        MsgBox "لا توجد نتائج مطابقة.", vbInformation, "بحث سريع"
        DoCmd.Close acForm, "frmCorrespondence"
    End If
End Function

'================= الملفات المرفقة ===========================

Public Function PickFile()
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    If fd Is Nothing Then
        Dim p As String
        p = InputBox("اكتب مسار الملف المرفق:", "إرفاق ملف")
        If Len(p) > 0 Then Screen.ActiveForm!FilePath = p
        Exit Function
    End If
    With fd
        .Title = "اختر الملف المرفق"
        .AllowMultiSelect = False
        If .Show Then Screen.ActiveForm!FilePath = .SelectedItems(1)
    End With
End Function

Public Function OpenAttachment()
    On Error GoTo ErrH
    Dim p As String
    p = Nz(Screen.ActiveForm!FilePath, "")
    If Len(p) = 0 Then
        MsgBox "لا يوجد ملف مرفق لهذه المكاتبة.", vbInformation, "الملف المرفق"
        Exit Function
    End If
    Application.FollowHyperlink p
    Exit Function
ErrH:
    MsgBox "تعذر فتح الملف: " & p & vbCrLf & Err.Description, vbExclamation, "خطأ"
End Function

'================= إغلاق وخروج ===============================

Public Function CloseForm()
    On Error Resume Next
    DoCmd.Close acForm, Screen.ActiveForm.Name
End Function

Public Function ExitApp()
    If MsgBox("هل تريد إنهاء البرنامج؟", vbYesNo + vbQuestion, "خروج") = vbYes Then
        Application.Quit acQuitSaveAll
    End If
End Function
