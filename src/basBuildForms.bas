Attribute VB_Name = "basBuildForms"
Option Compare Database
Option Explicit

'==============================================================
'  بناء النماذج برمجياً: الواجهة الرئيسية + نموذج المكاتبات
'  + نماذج الجهات والتصنيفات، مع أزرار تنقل وحفظ وبحث وتصفية
'==============================================================

Private mFrm As String          ' اسم النموذج الجاري بناؤه

Public Sub Build_Forms()
    DropFormIfExists "frmCorrespondence"
    DropFormIfExists "frmMain"
    DropFormIfExists "frmCategories"
    DropFormIfExists "frmEntities"

    BuildForm_Correspondence
    BuildLookupForm "frmCategories", "tblCategories", "التصنيفات", _
        Array("CategoryName"), Array("اسم التصنيف"), Array(4500&)
    BuildLookupForm "frmEntities", "tblEntities", "الجهات", _
        Array("EntityName", "ContactInfo"), Array("اسم الجهة", "بيانات الاتصال"), Array(4500&, 4500&)
    BuildForm_Main
End Sub

'==============================================================
'  نموذج المكاتبات الرئيسي (إدخال / تعديل / بحث / تصفية / تنقل)
'==============================================================

Private Sub BuildForm_Correspondence()
    Dim frm As Form
    Set frm = NewForm("سجل المكاتبات")
    frm.RecordSource = "tblCorrespondence"
    frm.Width = 12700
    frm.Section(acDetail).Height = 8400
    frm.Cycle = 1                       ' التنقل بالحقول داخل السجل الحالي فقط
    frm.OnCurrent = "=UpdateCounter()"  ' تحديث عداد السجلات

    ' ----- العنوان -----
    AddLabel "سجل المكاتبات - إضافة وتعديل وبحث", 150, 100, 12400, 500, 15, True, RGB(31, 78, 121)

    ' ----- شريط البحث والفرز -----
    AddLabel "بحث سريع:", 150, 810, 1100, 330, 10, True
    Dim t As TextBox
    Set t = AddText("txtSearch", "", 1300, 780, 3000, 380)
    t.AfterUpdate = "=DoSearch()"
    AddButton "btnSearch", "بحث", "=DoSearch()", 4400, 780, 1000, 400
    AddButton "btnShowAll", "عرض الكل", "=ClearFilters()", 5500, 780, 1300, 400
    AddButton "btnSortDate", "فرز بالتاريخ", "=SortToggle('CorrDate')", 6900, 780, 1600, 400
    AddButton "btnSortRef", "فرز بالرقم", "=SortToggle('RefNumber')", 8600, 780, 1600, 400

    ' ----- شريط التصفية -----
    AddLabel "النوع:", 150, 1330, 700, 330, 10, True
    Dim c As ComboBox
    Set c = AddCombo("cboFType", "", "Value List", "وارد;صادر", 1, "", 900, 1300, 1300, 380)
    c.AfterUpdate = "=ApplyFilters()"
    AddLabel "التصنيف:", 2300, 1330, 900, 330, 10, True
    Set c = AddCombo("cboFCat", "", "Table/Query", _
        "SELECT CategoryID, CategoryName FROM tblCategories ORDER BY CategoryName", _
        2, "0cm;5cm", 3250, 1300, 2100, 380)
    c.AfterUpdate = "=ApplyFilters()"
    AddLabel "الحالة:", 5450, 1330, 750, 330, 10, True
    Set c = AddCombo("cboFStatus", "", "Value List", "جديد;قيد الإجراء;منجز;محفوظ", 1, "", 6250, 1300, 1700, 380)
    c.AfterUpdate = "=ApplyFilters()"
    AddButton "btnApplyF", "تطبيق التصفية", "=ApplyFilters()", 8100, 1300, 1600, 400

    AddLine 150, 1810, 12400

    ' ----- حقول البيانات (عمودان) -----
    Dim y As Long
    y = 1980

    ' الصف 1: رقم المكاتبة | تاريخ المكاتبة
    AddLabel "رقم المكاتبة:", 150, y + 30, 1350, 330
    AddText "RefNumber", "RefNumber", 1550, y, 3900, 380
    AddLabel "تاريخ المكاتبة:", 7000, y + 30, 1350, 330
    Set t = AddText("CorrDate", "CorrDate", 8400, y, 3900, 380)
    t.Format = "yyyy/mm/dd"

    ' الصف 2: النوع | الأولوية
    y = y + 570
    AddLabel "النوع:", 150, y + 30, 1350, 330
    AddCombo "CorrType", "CorrType", "Value List", "وارد;صادر", 1, "", 1550, y, 3900, 380
    AddLabel "الأولوية:", 7000, y + 30, 1350, 330
    AddCombo "Priority", "Priority", "Value List", "عادي;هام;عاجل;سري", 1, "", 8400, y, 3900, 380

    ' الصف 3: التصنيف | الحالة
    y = y + 570
    AddLabel "التصنيف:", 150, y + 30, 1350, 330
    AddCombo "CategoryID", "CategoryID", "Table/Query", _
        "SELECT CategoryID, CategoryName FROM tblCategories ORDER BY CategoryName", _
        2, "0cm;5cm", 1550, y, 3900, 380
    AddLabel "الحالة:", 7000, y + 30, 1350, 330
    AddCombo "Status", "Status", "Value List", "جديد;قيد الإجراء;منجز;محفوظ", 1, "", 8400, y, 3900, 380

    ' الصف 4: الجهة | تاريخ الاستحقاق
    y = y + 570
    AddLabel "الجهة:", 150, y + 30, 1350, 330
    AddCombo "EntityID", "EntityID", "Table/Query", _
        "SELECT EntityID, EntityName FROM tblEntities ORDER BY EntityName", _
        2, "0cm;6cm", 1550, y, 3900, 380
    AddLabel "تاريخ الاستحقاق:", 7000, y + 30, 1350, 330
    Set t = AddText("DueDate", "DueDate", 8400, y, 3900, 380)
    t.Format = "yyyy/mm/dd"

    ' الصف 5: الموضوع (عرض كامل)
    y = y + 570
    AddLabel "الموضوع:", 150, y + 30, 1350, 330
    AddText "Subject", "Subject", 1550, y, 10750, 380

    ' الصف 6: رقم الإشارة | الملف المرفق
    y = y + 570
    AddLabel "رقم الإشارة/الرد:", 150, y + 30, 1350, 330
    AddText "ReplyRef", "ReplyRef", 1550, y, 3900, 380
    AddLabel "الملف المرفق:", 7000, y + 30, 1350, 330
    AddText "FilePath", "FilePath", 8400, y, 2500, 380
    AddButton "btnPickFile", "..", "=PickFile()", 10950, y, 450, 380
    AddButton "btnOpenFile", "فتح", "=OpenAttachment()", 11450, y, 850, 380

    ' الصف 7: الملاحظات
    y = y + 570
    AddLabel "ملاحظات:", 150, y + 30, 1350, 330
    Set t = AddText("Notes", "Notes", 1550, y, 10750, 1450)
    t.ScrollBars = 2
    t.EnterKeyBehavior = True

    AddLine 150, 7050, 12400

    ' ----- أزرار التنقل بين السجلات -----
    y = 7200
    AddButton "btnFirst", "|<  الأول", "=NavRec('first')", 150, y, 1150, 430
    AddButton "btnPrev", "<  السابق", "=NavRec('prev')", 1350, y, 1150, 430
    Dim lb As Label
    Set lb = AddNamedLabel("lblCounter", "سجل 1 من 1", 2600, y + 50, 2400, 330, 10, True)
    lb.TextAlign = 2
    AddButton "btnNext", "التالي  >", "=NavRec('next')", 5100, y, 1150, 430
    AddButton "btnLast", "الأخير  >|", "=NavRec('last')", 6350, y, 1150, 430

    ' ----- أزرار الإجراءات -----
    y = 7750
    AddButton "btnNew", "جديد", "=RecNew()", 150, y, 1400, 430
    AddButton "btnSave", "حفظ", "=RecSave()", 1600, y, 1400, 430
    AddButton "btnDelete", "حذف", "=RecDelete()", 3050, y, 1400, 430
    AddButton "btnUndo", "تراجع", "=RecUndo()", 4500, y, 1400, 430
    AddButton "btnPrint", "تقرير", "=OpenObj('report','rptCorrespondence')", 5950, y, 1400, 430
    AddButton "btnClose", "إغلاق", "=CloseForm()", 7400, y, 1400, 430

    SaveFormAs "frmCorrespondence"
End Sub

'==============================================================
'  الواجهة الرئيسية (لوحة التحكم)
'==============================================================

Private Sub BuildForm_Main()
    Dim frm As Form
    Set frm = NewForm("نظام إدارة المكاتبات - الواجهة الرئيسية")
    frm.Width = 10800
    frm.Section(acDetail).Height = 7000

    ' العنوان والتاريخ
    Dim lb As Label
    Set lb = AddNamedLabel("lblTitle", "نظام إدارة وتنظيم المكاتبات", 150, 200, 10500, 550, 16, True, RGB(31, 78, 121))
    lb.TextAlign = 2

    Dim t As TextBox
    Set t = AddText("txtToday", "", 150, 800, 10500, 350)
    t.ControlSource = "=""اليوم: "" & Format(Date(),""yyyy/mm/dd"")"
    t.Locked = True
    t.TabStop = False
    t.BorderStyle = 0
    t.BackStyle = 0
    t.TextAlign = 2

    ' ----- لوحة الإحصائيات -----
    AddStatTile "txtStatAll", "=DCount(""*"",""tblCorrespondence"")", "إجمالي المكاتبات", 150
    AddStatTile "txtStatIn", "=DCount(""*"",""tblCorrespondence"",""CorrType='وارد'"")", "الوارد", 2750
    AddStatTile "txtStatOut", "=DCount(""*"",""tblCorrespondence"",""CorrType='صادر'"")", "الصادر", 5350
    AddStatTile "txtStatLate", "=DCount(""*"",""tblCorrespondence"",""[Status]<>'منجز' AND [DueDate]<Date()"")", "متأخرة", 7950

    ' ----- أزرار التشغيل -----
    Dim y As Long
    y = 2700
    AddButton "btnNewCorr", "تسجيل مكاتبة جديدة", "=OpenCorr('new')", 800, y, 4300, 540, 11
    AddButton "btnAllCorr", "سجل المكاتبات", "=OpenCorr('all')", 5700, y, 4300, 540, 11

    y = y + 680
    AddButton "btnIn", "المكاتبات الواردة", "=OpenCorr('in')", 800, y, 4300, 540, 11
    AddButton "btnOut", "المكاتبات الصادرة", "=OpenCorr('out')", 5700, y, 4300, 540, 11

    y = y + 680
    AddButton "btnLate", "المكاتبات المتأخرة", "=OpenCorr('late')", 800, y, 4300, 540, 11
    AddButton "btnQuick", "بحث سريع", "=QuickSearch()", 5700, y, 4300, 540, 11

    y = y + 680
    AddButton "btnEntities", "الجهات", "=OpenObj('form','frmEntities')", 800, y, 4300, 540, 11
    AddButton "btnCats", "التصنيفات", "=OpenObj('form','frmCategories')", 5700, y, 4300, 540, 11

    y = y + 680
    AddButton "btnReport", "تقرير المكاتبات (طباعة)", "=OpenObj('report','rptCorrespondence')", 800, y, 4300, 540, 11
    AddButton "btnStats", "إحصائية شهرية", "=OpenObj('query','qryMonthlyStats')", 5700, y, 4300, 540, 11

    y = y + 680
    AddButton "btnExit", "إنهاء البرنامج", "=ExitApp()", 3250, y, 4300, 540, 11

    SaveFormAs "frmMain"
End Sub

' بطاقة إحصائية: رقم كبير + تسمية أسفله
Private Sub AddStatTile(sName As String, sSrc As String, sCaption As String, x As Long)
    Dim t As TextBox
    Set t = AddText(sName, "", x, 1350, 2400, 550)
    t.ControlSource = sSrc
    t.Locked = True
    t.TabStop = False
    t.TextAlign = 2
    t.FontSize = 16
    t.FontWeight = 700
    t.ForeColor = RGB(31, 78, 121)
    Dim lb As Label
    Set lb = AddNamedLabel("lbl" & sName, sCaption, x, 1930, 2400, 320, 10, True)
    lb.TextAlign = 2
End Sub

'==============================================================
'  نماذج البيانات المساعدة (جداول مرجعية بعرض ورقة بيانات)
'==============================================================

Private Sub BuildLookupForm(sForm As String, sTable As String, sCaption As String, _
                            vFields As Variant, vCaptions As Variant, vWidths As Variant)
    Dim frm As Form
    Set frm = NewForm(sCaption)
    frm.RecordSource = sTable
    frm.DefaultView = 2            ' ورقة بيانات
    frm.NavigationButtons = True
    frm.RecordSelectors = True

    Dim i As Integer, t As TextBox, lb As Label
    For i = LBound(vFields) To UBound(vFields)
        Set t = AddText(CStr(vFields(i)), CStr(vFields(i)), 150, 150 + i * 450, CLng(vWidths(i)), 380)
        ' تسمية مرفقة لتظهر كعنوان العمود في ورقة البيانات
        Set lb = CreateControl(mFrm, acLabel, acDetail, t.Name, "", 150, 150 + i * 450, 100, 100)
        lb.Caption = CStr(vCaptions(i))
    Next i

    SaveFormAs sForm
End Sub

'==============================================================
'  أدوات بناء النماذج
'==============================================================

Private Function NewForm(sCaption As String) As Form
    Dim frm As Form
    Set frm = CreateForm
    mFrm = frm.Name
    With frm
        .Caption = sCaption
        On Error Resume Next
        .Properties("Orientation") = 1      ' اتجاه من اليمين إلى اليسار
        On Error GoTo 0
        .DefaultView = 0                    ' نموذج مفرد
        .NavigationButtons = False
        .RecordSelectors = False
        .DividingLines = False
        .AutoCenter = True
    End With
    Set NewForm = frm
End Function

Private Sub SaveFormAs(sName As String)
    DoCmd.Save acForm, mFrm
    DoCmd.Close acForm, mFrm, acSaveYes
    On Error Resume Next
    DoCmd.DeleteObject acForm, sName
    On Error GoTo 0
    DoCmd.Rename sName, acForm, mFrm
End Sub

Private Sub AddLabel(sCaption As String, x As Long, y As Long, w As Long, h As Long, _
                     Optional nSize As Integer = 10, Optional bBold As Boolean = False, _
                     Optional lColor As Long = -1)
    AddNamedLabel "", sCaption, x, y, w, h, nSize, bBold, lColor
End Sub

Private Function AddNamedLabel(sName As String, sCaption As String, x As Long, y As Long, _
                               w As Long, h As Long, Optional nSize As Integer = 10, _
                               Optional bBold As Boolean = False, _
                               Optional lColor As Long = -1) As Label
    Dim lb As Label
    Set lb = CreateControl(mFrm, acLabel, acDetail, "", "", x, y, w, h)
    lb.Caption = sCaption
    If Len(sName) > 0 Then lb.Name = sName
    lb.FontSize = nSize
    If bBold Then lb.FontWeight = 700
    If lColor <> -1 Then lb.ForeColor = lColor
    Set AddNamedLabel = lb
End Function

Private Function AddText(sName As String, sSource As String, x As Long, y As Long, _
                         w As Long, h As Long) As TextBox
    Dim t As TextBox
    Set t = CreateControl(mFrm, acTextBox, acDetail, "", sSource, x, y, w, h)
    t.Name = sName
    t.FontSize = 10
    Set AddText = t
End Function

Private Function AddCombo(sName As String, sSource As String, sRowType As String, _
                          sRowSource As String, nCols As Integer, sColWidths As String, _
                          x As Long, y As Long, w As Long, h As Long) As ComboBox
    Dim c As ComboBox
    Set c = CreateControl(mFrm, acComboBox, acDetail, "", sSource, x, y, w, h)
    c.Name = sName
    c.FontSize = 10
    c.RowSourceType = sRowType
    c.RowSource = sRowSource
    c.ColumnCount = nCols
    If Len(sColWidths) > 0 Then c.ColumnWidths = sColWidths
    If nCols > 1 Then c.LimitToList = True
    Set AddCombo = c
End Function

Private Sub AddButton(sName As String, sCaption As String, sOnClick As String, _
                      x As Long, y As Long, w As Long, h As Long, _
                      Optional nSize As Integer = 10)
    Dim b As CommandButton
    Set b = CreateControl(mFrm, acCommandButton, acDetail, "", "", x, y, w, h)
    b.Name = sName
    b.Caption = sCaption
    b.OnClick = sOnClick
    b.FontSize = nSize
End Sub

Private Sub AddLine(x As Long, y As Long, w As Long)
    Dim c As Control
    Set c = CreateControl(mFrm, acLine, acDetail, "", "", x, y, w, 0)
End Sub
