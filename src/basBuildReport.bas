Attribute VB_Name = "basBuildReport"
Option Compare Database
Option Explicit

'==============================================================
'  بناء تقرير سجل المكاتبات (جدولي - أفقي)
'==============================================================

Private mRpt As String      ' اسم التقرير الجاري بناؤه

Public Sub Build_Report()
    DropReportIfExists "rptCorrespondence"

    Dim rpt As Report
    Set rpt = CreateReport
    mRpt = rpt.Name

    rpt.RecordSource = "qryAllCorrespondence"
    rpt.Caption = "تقرير سجل المكاتبات"
    On Error Resume Next
    rpt.Properties("Orientation") = 1          ' من اليمين إلى اليسار
    rpt.Printer.Orientation = 2                ' طباعة أفقية (قد يفشل بدون طابعة)
    On Error GoTo 0
    rpt.Width = 14000

    ' التأكد من وجود أقسام الرأس والتذييل
    EnsureSection rpt, acPageHeader, acCmdPageHdrFtr
    EnsureSection rpt, acHeader, acCmdReportHdrFtr

    rpt.Section(acHeader).Height = 900
    rpt.Section(acPageHeader).Height = 500
    rpt.Section(acDetail).Height = 420
    rpt.Section(acPageFooter).Height = 450
    rpt.Section(acFooter).Height = 500

    ' ----- رأس التقرير -----
    RLabel acHeader, "تقرير سجل المكاتبات", 150, 100, 13700, 500, 15, True
    Dim t As TextBox
    Set t = RText(acHeader, "", 150, 620, 13700, 250)
    t.ControlSource = "=""تاريخ الطباعة: "" & Format(Now(),""yyyy/mm/dd hh:nn"")"
    t.FontSize = 9
    t.BorderStyle = 0

    ' ----- أعمدة التقرير -----
    Dim vCaps As Variant, vSrc As Variant, vW As Variant
    vCaps = Array("رقم المكاتبة", "التاريخ", "النوع", "التصنيف", "الجهة", "الموضوع", "الحالة")
    vSrc = Array("[رقم المكاتبة]", "[التاريخ]", "[النوع]", "[التصنيف]", "[الجهة]", "[الموضوع]", "[الحالة]")
    vW = Array(1500&, 1300&, 900&, 1900&, 2300&, 4500&, 1300&)

    Dim i As Integer, x As Long
    x = 150
    For i = LBound(vCaps) To UBound(vCaps)
        ' عناوين الأعمدة في رأس الصفحة
        RLabel acPageHeader, CStr(vCaps(i)), x, 80, CLng(vW(i)), 330, 10, True
        ' البيانات في قسم التفصيل
        Set t = RText(acDetail, CStr(vSrc(i)), x, 30, CLng(vW(i)), 330)
        If i = 5 Then t.CanGrow = True     ' عمود الموضوع يتمدد
        x = x + CLng(vW(i)) + 50
    Next i

    ' خط أسفل عناوين الأعمدة
    Dim c As Control
    Set c = CreateReportControl(mRpt, acLine, acPageHeader, "", "", 150, 460, 13550, 0)

    ' ----- تذييل الصفحة: ترقيم الصفحات -----
    Set t = RText(acPageFooter, "", 150, 80, 13700, 300)
    t.ControlSource = "=""صفحة "" & [Page] & "" من "" & [Pages]"
    t.TextAlign = 2
    t.BorderStyle = 0

    ' ----- تذييل التقرير: الإجمالي -----
    Set t = RText(acFooter, "", 150, 100, 13700, 350)
    t.ControlSource = "=""إجمالي عدد المكاتبات: "" & Count(*)"
    t.FontSize = 11
    t.FontWeight = 700
    t.BorderStyle = 0

    DoCmd.Save acReport, mRpt
    DoCmd.Close acReport, mRpt, acSaveYes
    On Error Resume Next
    DoCmd.DeleteObject acReport, "rptCorrespondence"
    On Error GoTo 0
    DoCmd.Rename "rptCorrespondence", acReport, mRpt
End Sub

' تفعيل قسم في التقرير إذا لم يكن موجوداً
Private Sub EnsureSection(rpt As Report, iSec As Integer, lCmd As Long)
    Dim h As Long
    On Error Resume Next
    h = rpt.Section(iSec).Height
    If Err.Number <> 0 Then
        Err.Clear
        DoCmd.RunCommand lCmd
    End If
    On Error GoTo 0
End Sub

Private Sub RLabel(iSec As Integer, sCaption As String, x As Long, y As Long, _
                   w As Long, h As Long, Optional nSize As Integer = 10, _
                   Optional bBold As Boolean = False)
    Dim lb As Label
    Set lb = CreateReportControl(mRpt, acLabel, iSec, "", "", x, y, w, h)
    lb.Caption = sCaption
    lb.FontSize = nSize
    If bBold Then lb.FontWeight = 700
End Sub

Private Function RText(iSec As Integer, sSource As String, x As Long, y As Long, _
                       w As Long, h As Long) As TextBox
    Dim t As TextBox
    Set t = CreateReportControl(mRpt, acTextBox, iSec, "", sSource, x, y, w, h)
    t.FontSize = 10
    Set RText = t
End Function
