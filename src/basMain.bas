Attribute VB_Name = "basMain"
Option Compare Database
Option Explicit

'==============================================================
'  نظام إدارة وتنظيم المكاتبات - الوحدة الرئيسية
'  Correspondence Management System - Main builder module
'--------------------------------------------------------------
'  طريقة الاستخدام:
'  1) أنشئ قاعدة بيانات Access فارغة (.accdb)
'  2) استورد جميع ملفات الوحدات (.bas) من مجلد dist
'  3) افتح النافذة الفورية (Ctrl+G) واكتب:  BuildAll
'  4) اضغط Enter وسيتم بناء البرنامج كاملاً تلقائياً
'==============================================================

Public Sub BuildAll()
    On Error GoTo ErrH
    DoCmd.Hourglass True

    BuildCore

    DoCmd.Hourglass False
    MsgBox "تم بناء نظام إدارة المكاتبات بنجاح." & vbCrLf & vbCrLf & _
           "سيتم الآن فتح الواجهة الرئيسية." & vbCrLf & _
           "عند فتح الملف مستقبلاً ستظهر الواجهة الرئيسية تلقائياً.", _
           vbInformation, "اكتمل البناء"
    DoCmd.OpenForm "frmMain"
    Exit Sub

ErrH:
    DoCmd.Hourglass False
    MsgBox "حدث خطأ أثناء البناء:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "خطأ"
End Sub

' بناء صامت (بدون رسائل) - يُستدعى من سكربت Build-Accdb.ps1
Public Sub BuildAllSilent()
    BuildCore
End Sub

' خطوات البناء المشتركة
Private Sub BuildCore()
    Build_Tables      ' الجداول + العلاقات + البيانات الأولية
    Build_Queries     ' الاستعلامات الجاهزة
    Build_Forms       ' النماذج وأزرار التنقل
    Build_Report      ' تقرير الطباعة
    SetStartupForm "frmMain"
End Sub

'================= أدوات مساعدة عامة =========================

Public Sub DropTableIfExists(sName As String)
    On Error Resume Next
    CurrentDb.TableDefs.Delete sName
End Sub

Public Sub DropQueryIfExists(sName As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete sName
End Sub

Public Sub DropFormIfExists(sName As String)
    On Error Resume Next
    DoCmd.Close acForm, sName, acSaveNo
    DoCmd.DeleteObject acForm, sName
End Sub

Public Sub DropReportIfExists(sName As String)
    On Error Resume Next
    DoCmd.Close acReport, sName, acSaveNo
    DoCmd.DeleteObject acReport, sName
End Sub

' تعيين نموذج بدء التشغيل بحيث تفتح الواجهة الرئيسية تلقائياً
Public Sub SetStartupForm(sName As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties("StartupForm") = sName
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("StartupForm", dbText, sName)
    End If
    On Error GoTo 0
End Sub

' تنسيق تاريخ لاستخدامه داخل جمل SQL
Public Function SqlDate(d As Date) As String
    SqlDate = Format$(d, "\#yyyy-mm-dd\#")
End Function
