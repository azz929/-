Attribute VB_Name = "basBuildTables"
Option Compare Database
Option Explicit

'==============================================================
'  بناء الجداول والعلاقات والفهارس والبيانات الأولية
'  وفق قواعد التطبيع (Normalization) مع تكامل مرجعي كامل
'==============================================================

Public Sub Build_Tables()
    Dim db As DAO.Database
    Set db = CurrentDb

    ' حذف الجداول القديمة (الجدول الرئيسي أولاً بسبب العلاقات)
    DropTableIfExists "tblCorrespondence"
    DropTableIfExists "tblCategories"
    DropTableIfExists "tblEntities"

    '---------- جدول التصنيفات ----------
    db.Execute _
        "CREATE TABLE tblCategories (" & _
        "CategoryID AUTOINCREMENT CONSTRAINT PK_Cat PRIMARY KEY, " & _
        "CategoryName TEXT(100) NOT NULL)", dbFailOnError

    '---------- جدول الجهات ----------
    db.Execute _
        "CREATE TABLE tblEntities (" & _
        "EntityID AUTOINCREMENT CONSTRAINT PK_Ent PRIMARY KEY, " & _
        "EntityName TEXT(150) NOT NULL, " & _
        "ContactInfo TEXT(255))", dbFailOnError

    '---------- جدول المكاتبات (الرئيسي) ----------
    db.Execute _
        "CREATE TABLE tblCorrespondence (" & _
        "CorrID AUTOINCREMENT CONSTRAINT PK_Corr PRIMARY KEY, " & _
        "RefNumber TEXT(50) NOT NULL, " & _
        "CorrDate DATETIME, " & _
        "CorrType TEXT(20), " & _
        "CategoryID LONG CONSTRAINT FK_Corr_Cat REFERENCES tblCategories (CategoryID), " & _
        "EntityID LONG CONSTRAINT FK_Corr_Ent REFERENCES tblEntities (EntityID), " & _
        "Subject TEXT(255), " & _
        "Priority TEXT(20), " & _
        "Status TEXT(30), " & _
        "DueDate DATETIME, " & _
        "ReplyRef TEXT(50), " & _
        "FilePath TEXT(255), " & _
        "Notes MEMO, " & _
        "CreatedAt DATETIME)", dbFailOnError

    '---------- الفهارس لتسريع البحث والفرز ----------
    db.Execute "CREATE INDEX idxCorrDate ON tblCorrespondence (CorrDate)", dbFailOnError
    db.Execute "CREATE INDEX idxRefNumber ON tblCorrespondence (RefNumber)", dbFailOnError
    db.Execute "CREATE INDEX idxSubject ON tblCorrespondence (Subject)", dbFailOnError
    db.Execute "CREATE INDEX idxStatus ON tblCorrespondence (Status)", dbFailOnError

    '---------- القيم الافتراضية ----------
    SetDefault "tblCorrespondence", "CorrDate", "Date()"
    SetDefault "tblCorrespondence", "CorrType", """وارد"""
    SetDefault "tblCorrespondence", "Priority", """عادي"""
    SetDefault "tblCorrespondence", "Status", """قيد الإجراء"""
    SetDefault "tblCorrespondence", "CreatedAt", "Now()"

    '---------- التسميات العربية للحقول ----------
    SetCaption "tblCategories", "CategoryID", "م"
    SetCaption "tblCategories", "CategoryName", "اسم التصنيف"

    SetCaption "tblEntities", "EntityID", "م"
    SetCaption "tblEntities", "EntityName", "اسم الجهة"
    SetCaption "tblEntities", "ContactInfo", "بيانات الاتصال"

    SetCaption "tblCorrespondence", "CorrID", "م"
    SetCaption "tblCorrespondence", "RefNumber", "رقم المكاتبة"
    SetCaption "tblCorrespondence", "CorrDate", "تاريخ المكاتبة"
    SetCaption "tblCorrespondence", "CorrType", "النوع"
    SetCaption "tblCorrespondence", "CategoryID", "التصنيف"
    SetCaption "tblCorrespondence", "EntityID", "الجهة"
    SetCaption "tblCorrespondence", "Subject", "الموضوع"
    SetCaption "tblCorrespondence", "Priority", "الأولوية"
    SetCaption "tblCorrespondence", "Status", "الحالة"
    SetCaption "tblCorrespondence", "DueDate", "تاريخ الاستحقاق"
    SetCaption "tblCorrespondence", "ReplyRef", "رقم الإشارة/الرد"
    SetCaption "tblCorrespondence", "FilePath", "الملف المرفق"
    SetCaption "tblCorrespondence", "Notes", "ملاحظات"
    SetCaption "tblCorrespondence", "CreatedAt", "تاريخ الإدخال"

    SeedData db
    db.TableDefs.Refresh
End Sub

'================= بيانات أولية للتجربة ======================

Private Sub SeedData(db As DAO.Database)
    Dim arr As Variant, i As Integer

    ' التصنيفات
    arr = Array("تعاميم", "قرارات إدارية", "خطابات رسمية", "مذكرات داخلية", _
                "عقود واتفاقيات", "شؤون الموظفين", "الشؤون المالية", "أخرى")
    For i = LBound(arr) To UBound(arr)
        db.Execute "INSERT INTO tblCategories (CategoryName) VALUES ('" & arr(i) & "')", dbFailOnError
    Next i

    ' الجهات
    db.Execute "INSERT INTO tblEntities (EntityName, ContactInfo) VALUES ('ديوان المحافظة', 'هاتف: 0112345678')", dbFailOnError
    db.Execute "INSERT INTO tblEntities (EntityName, ContactInfo) VALUES ('وزارة المالية', 'هاتف: 0119876543')", dbFailOnError
    db.Execute "INSERT INTO tblEntities (EntityName, ContactInfo) VALUES ('إدارة الموارد البشرية', 'داخلي: 105')", dbFailOnError
    db.Execute "INSERT INTO tblEntities (EntityName, ContactInfo) VALUES ('الشؤون القانونية', 'داخلي: 210')", dbFailOnError
    db.Execute "INSERT INTO tblEntities (EntityName, ContactInfo) VALUES ('شركة الاتصالات', 'بريد: info@telecom.example')", dbFailOnError

    ' مكاتبات تجريبية
    AddSample db, "245/و/26", Date - 20, "وارد", 3, 1, _
        "طلب موافاة ببيانات الموظفين المنتدبين", "عادي", "منجز", Date - 15, "112/ص/26"
    AddSample db, "112/ص/26", Date - 18, "صادر", 3, 1, _
        "الرد على طلب بيانات الموظفين المنتدبين", "عادي", "منجز", 0, ""
    AddSample db, "246/و/26", Date - 10, "وارد", 1, 2, _
        "تعميم بشأن مواعيد العمل الرسمية", "هام", "محفوظ", 0, ""
    AddSample db, "247/و/26", Date - 6, "وارد", 5, 5, _
        "مسودة عقد صيانة أجهزة الحاسب الآلي للاعتماد", "عاجل", "قيد الإجراء", Date + 2, ""
    AddSample db, "113/ص/26", Date - 3, "صادر", 4, 3, _
        "مذكرة بشأن تحديث لائحة الجزاءات الإدارية", "عادي", "قيد الإجراء", Date + 7, ""
    AddSample db, "248/و/26", Date - 15, "وارد", 7, 2, _
        "مطالبة مالية مستحقة عن الربع الأول", "عاجل", "قيد الإجراء", Date - 4, ""
End Sub

Private Sub AddSample(db As DAO.Database, sRef As String, dDate As Date, sType As String, _
                      lCat As Long, lEnt As Long, sSubject As String, sPriority As String, _
                      sStatus As String, dDue As Date, sReply As String)
    Dim sSql As String
    sSql = "INSERT INTO tblCorrespondence " & _
           "(RefNumber, CorrDate, CorrType, CategoryID, EntityID, Subject, Priority, Status, DueDate, ReplyRef, CreatedAt) VALUES ('" & _
           sRef & "', " & SqlDate(dDate) & ", '" & sType & "', " & lCat & ", " & lEnt & ", '" & _
           sSubject & "', '" & sPriority & "', '" & sStatus & "', " & _
           IIf(dDue = 0, "NULL", SqlDate(dDue)) & ", " & _
           IIf(sReply = "", "NULL", "'" & sReply & "'") & ", " & SqlDate(Date) & ")"
    db.Execute sSql, dbFailOnError
End Sub

'================= أدوات مساعدة ==============================

Private Sub SetDefault(sTable As String, sField As String, sExpr As String)
    CurrentDb.TableDefs(sTable).Fields(sField).DefaultValue = sExpr
End Sub

Private Sub SetCaption(sTable As String, sField As String, sCaption As String)
    Dim f As DAO.Field
    Set f = CurrentDb.TableDefs(sTable).Fields(sField)
    On Error Resume Next
    f.Properties("Caption") = sCaption
    If Err.Number <> 0 Then
        Err.Clear
        f.Properties.Append f.CreateProperty("Caption", dbText, sCaption)
    End If
    On Error GoTo 0
End Sub
