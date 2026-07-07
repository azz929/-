Attribute VB_Name = "basBuildQueries"
Option Compare Database
Option Explicit

'==============================================================
'  بناء الاستعلامات الجاهزة
'==============================================================

Public Sub Build_Queries()
    Dim db As DAO.Database
    Set db = CurrentDb

    DropQueryIfExists "qryAllCorrespondence"
    DropQueryIfExists "qryIncoming"
    DropQueryIfExists "qryOutgoing"
    DropQueryIfExists "qryOverdue"
    DropQueryIfExists "qrySearchParam"
    DropQueryIfExists "qryMonthlyStats"
    DropQueryIfExists "qryByCategory"

    Dim sBase As String
    sBase = "SELECT c.CorrID AS [م], c.RefNumber AS [رقم المكاتبة], c.CorrDate AS [التاريخ], " & _
            "c.CorrType AS [النوع], cat.CategoryName AS [التصنيف], e.EntityName AS [الجهة], " & _
            "c.Subject AS [الموضوع], c.Priority AS [الأولوية], c.Status AS [الحالة], " & _
            "c.DueDate AS [تاريخ الاستحقاق], c.ReplyRef AS [رقم الإشارة] " & _
            "FROM (tblCorrespondence AS c LEFT JOIN tblCategories AS cat ON c.CategoryID = cat.CategoryID) " & _
            "LEFT JOIN tblEntities AS e ON c.EntityID = e.EntityID "

    ' سجل شامل بجميع المكاتبات
    db.CreateQueryDef "qryAllCorrespondence", sBase & "ORDER BY c.CorrDate DESC;"

    ' المكاتبات الواردة
    db.CreateQueryDef "qryIncoming", sBase & "WHERE c.CorrType = 'وارد' ORDER BY c.CorrDate DESC;"

    ' المكاتبات الصادرة
    db.CreateQueryDef "qryOutgoing", sBase & "WHERE c.CorrType = 'صادر' ORDER BY c.CorrDate DESC;"

    ' المكاتبات المتأخرة (تجاوزت تاريخ الاستحقاق ولم تنجز)
    db.CreateQueryDef "qryOverdue", sBase & _
        "WHERE c.Status <> 'منجز' AND c.DueDate < Date() ORDER BY c.DueDate;"

    ' استعلام بحث معلمي (يسأل المستخدم عن كلمة البحث عند التشغيل)
    db.CreateQueryDef "qrySearchParam", _
        "PARAMETERS [اكتب كلمة البحث] TEXT (255); " & sBase & _
        "WHERE c.RefNumber LIKE '*' & [اكتب كلمة البحث] & '*' " & _
        "OR c.Subject LIKE '*' & [اكتب كلمة البحث] & '*' " & _
        "OR c.Notes LIKE '*' & [اكتب كلمة البحث] & '*' " & _
        "OR e.EntityName LIKE '*' & [اكتب كلمة البحث] & '*' " & _
        "OR cat.CategoryName LIKE '*' & [اكتب كلمة البحث] & '*' " & _
        "ORDER BY c.CorrDate DESC;"

    ' إحصائية شهرية (وارد/صادر لكل شهر)
    db.CreateQueryDef "qryMonthlyStats", _
        "SELECT Format(c.CorrDate,'yyyy/mm') AS [الشهر], c.CorrType AS [النوع], Count(*) AS [العدد] " & _
        "FROM tblCorrespondence AS c " & _
        "WHERE c.CorrDate Is Not Null " & _
        "GROUP BY Format(c.CorrDate,'yyyy/mm'), c.CorrType " & _
        "ORDER BY Format(c.CorrDate,'yyyy/mm') DESC;"

    ' عدد المكاتبات حسب التصنيف
    db.CreateQueryDef "qryByCategory", _
        "SELECT cat.CategoryName AS [التصنيف], Count(c.CorrID) AS [عدد المكاتبات] " & _
        "FROM tblCategories AS cat LEFT JOIN tblCorrespondence AS c ON cat.CategoryID = c.CategoryID " & _
        "GROUP BY cat.CategoryName " & _
        "ORDER BY Count(c.CorrID) DESC;"

    db.QueryDefs.Refresh
End Sub
