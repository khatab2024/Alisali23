Attribute VB_Name = "بناء_نظام_الفحص"
' =====================================================================
'  بانِي نظام "استمارة الفحص الفني الشامل" لقاعدة بيانات Microsoft Access
'  ينشئ: الجداول + العلاقات + الاستعلام + بيانات تجريبية
'
'  طريقة الاستخدام:
'   1) في Access: اضغط Alt+F11 لفتح محرر VBA.
'   2) File > Import File...  واختر هذا الملف (اكسس_بناء_النظام.bas)
'      أو: Insert > Module ثم الصق الكود.
'   3) شغّل الإجراء:  تشغيل_كامل
'      (ضع المؤشر داخله واضغط F5).
'  ملاحظة: يتطلب مرجع "Microsoft Office x.x Access Database Engine Object"
'          وهو مفعّل افتراضياً (DAO).
' =====================================================================
Option Compare Database
Option Explicit

' ---------- الإجراء الرئيسي: ينفّذ كل شيء ----------
Public Sub تشغيل_كامل()
    On Error GoTo خطأ
    انشاء_الجداول
    انشاء_العلاقات
    انشاء_الاستعلام
    ادراج_بيانات_تجريبية
    MsgBox "تم إنشاء نظام الفحص بالكامل (جداول + علاقات + استعلام + بيانات تجريبية)." & vbCrLf & _
           "افتح استعلام: qryتفاصيل_الفحص لرؤية النتيجة.", vbInformation, "اكتمل"
    Exit Sub
خطأ:
    MsgBox "حدث خطأ: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ---------- 1) الجداول ----------
Public Sub انشاء_الجداول()
    Dim db As DAO.Database
    Set db = CurrentDb

    حذف_جدول db, "tblالمراحل"
    حذف_جدول db, "tblالاستمارات"
    حذف_جدول db, "tblالقطع"

    Dim td As DAO.TableDef, idx As DAO.Index, f As DAO.Field

    ' --- جدول القطع ---
    Set td = db.CreateTableDef("tblالقطع")
    td.Fields.Append td.CreateField("رقم_القطعة", dbText, 50)
    td.Fields.Append td.CreateField("اسم_القطعة", dbText, 100)
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("رقم_القطعة")
    idx.Primary = True
    td.Indexes.Append idx
    db.TableDefs.Append td

    ' --- جدول الاستمارات (ترويسة لكل قطعة) ---
    Set td = db.CreateTableDef("tblالاستمارات")
    Set f = td.CreateField("معرف_الاستمارة", dbLong)
    f.Attributes = dbAutoIncrField
    td.Fields.Append f
    td.Fields.Append td.CreateField("رقم_القطعة", dbText, 50)
    td.Fields.Append td.CreateField("اسم_المشروع", dbText, 150)
    td.Fields.Append td.CreateField("التاريخ", dbText, 20)
    td.Fields.Append td.CreateField("المجموعة", dbText, 50)
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("معرف_الاستمارة")
    idx.Primary = True
    td.Indexes.Append idx
    db.TableDefs.Append td

    ' --- جدول المراحل (تفاصيل الفحص) ---
    Set td = db.CreateTableDef("tblالمراحل")
    Set f = td.CreateField("المعرف", dbLong)
    f.Attributes = dbAutoIncrField
    td.Fields.Append f
    td.Fields.Append td.CreateField("معرف_الاستمارة", dbLong)
    td.Fields.Append td.CreateField("م", dbInteger)
    td.Fields.Append td.CreateField("مرحلة_العمل", dbText, 150)
    td.Fields.Append td.CreateField("عدد_القطع", dbInteger)
    td.Fields.Append td.CreateField("الفني", dbText, 100)
    td.Fields.Append td.CreateField("القرار", dbText, 20)      ' مقبول / مرفوض
    td.Fields.Append td.CreateField("السماحية", dbText, 50)
    td.Fields.Append td.CreateField("السبب", dbText, 255)
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("المعرف")
    idx.Primary = True
    td.Indexes.Append idx
    db.TableDefs.Append td
End Sub

' ---------- 2) العلاقات ----------
Public Sub انشاء_العلاقات()
    Dim db As DAO.Database
    Set db = CurrentDb
    حذف_علاقة db, "rel_القطع_الاستمارات"
    حذف_علاقة db, "rel_الاستمارات_المراحل"

    Dim rel As DAO.Relation
    ' القطع (1) --- (∞) الاستمارات
    Set rel = db.CreateRelation("rel_القطع_الاستمارات", "tblالقطع", "tblالاستمارات", _
              dbRelationUpdateCascade + dbRelationDeleteCascade)
    rel.Fields.Append rel.CreateField("رقم_القطعة")
    rel.Fields("رقم_القطعة").ForeignName = "رقم_القطعة"
    db.Relations.Append rel

    ' الاستمارات (1) --- (∞) المراحل
    Set rel = db.CreateRelation("rel_الاستمارات_المراحل", "tblالاستمارات", "tblالمراحل", _
              dbRelationUpdateCascade + dbRelationDeleteCascade)
    rel.Fields.Append rel.CreateField("معرف_الاستمارة")
    rel.Fields("معرف_الاستمارة").ForeignName = "معرف_الاستمارة"
    db.Relations.Append rel
End Sub

' ---------- 3) الاستعلام ----------
Public Sub انشاء_الاستعلام()
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.QueryDefs.Delete "qryتفاصيل_الفحص"
    On Error GoTo 0

    Dim sql As String
    sql = "SELECT s.معرف_الاستمارة, p.رقم_القطعة, p.اسم_القطعة, s.اسم_المشروع, " & _
          "s.التاريخ, s.المجموعة, d.م, d.مرحلة_العمل, d.عدد_القطع, d.الفني, " & _
          "d.القرار, d.السماحية, d.السبب " & _
          "FROM (tblالقطع AS p INNER JOIN tblالاستمارات AS s ON p.رقم_القطعة = s.رقم_القطعة) " & _
          "INNER JOIN tblالمراحل AS d ON s.معرف_الاستمارة = d.معرف_الاستمارة " & _
          "ORDER BY s.معرف_الاستمارة, d.م;"
    db.CreateQueryDef "qryتفاصيل_الفحص", sql
End Sub

' ---------- 4) بيانات تجريبية ----------
Public Sub ادراج_بيانات_تجريبية()
    Dim db As DAO.Database
    Set db = CurrentDb

    ' تفريغ القديم
    db.Execute "DELETE FROM tblالمراحل", dbFailOnError
    db.Execute "DELETE FROM tblالاستمارات", dbFailOnError
    db.Execute "DELETE FROM tblالقطع", dbFailOnError

    ' القطع
    db.Execute "INSERT INTO tblالقطع (رقم_القطعة, اسم_القطعة) VALUES ('P-1001','عمود إدارة رئيسي')", dbFailOnError
    db.Execute "INSERT INTO tblالقطع (رقم_القطعة, اسم_القطعة) VALUES ('P-1002','غطاء علبة التروس')", dbFailOnError

    Dim id1 As Long, id2 As Long

    ' استمارة القطعة الأولى
    db.Execute "INSERT INTO tblالاستمارات (رقم_القطعة, اسم_المشروع, التاريخ, المجموعة) " & _
               "VALUES ('P-1001','مشروع الخطوط الناقلة','01/07/2026','المجموعة أ')", dbFailOnError
    id1 = db.OpenRecordset("SELECT @@IDENTITY")(0)
    اضف_مرحلة db, id1, 1, "خراطة أولية", 5, "أحمد علي", "مقبول", "±0.05", ""
    اضف_مرحلة db, id1, 2, "تفريز مجاري", 5, "أحمد علي", "مقبول", "±0.02", ""
    اضف_مرحلة db, id1, 3, "تجليخ نهائي", 5, "سالم محمد", "مرفوض", "±0.01", "خدش سطحي خارج الحد"

    ' استمارة القطعة الثانية
    db.Execute "INSERT INTO tblالاستمارات (رقم_القطعة, اسم_المشروع, التاريخ, المجموعة) " & _
               "VALUES ('P-1002','مشروع الصيانة الدورية','02/07/2026','المجموعة ب')", dbFailOnError
    id2 = db.OpenRecordset("SELECT @@IDENTITY")(0)
    اضف_مرحلة db, id2, 1, "قص وتشكيل", 8, "خالد يحيى", "مقبول", "±0.10", ""
    اضف_مرحلة db, id2, 2, "ثقب المسامير", 8, "خالد يحيى", "مقبول", "±0.05", ""
    اضف_مرحلة db, id2, 3, "تسوية الأسطح", 8, "ماجد سعيد", "مقبول", "±0.03", ""
    اضف_مرحلة db, id2, 4, "طلاء واقٍ", 8, "ماجد سعيد", "مرفوض", "-", "تفاوت في سماكة الطلاء"
End Sub

' =====================================================================
'  إجراءات مساعدة
' =====================================================================
Private Sub اضف_مرحلة(db As DAO.Database, معرف As Long, رقم As Integer, _
                       مرحلة As String, عدد As Integer, فني As String, _
                       قرار As String, سماحية As String, سبب As String)
    Dim sql As String
    sql = "INSERT INTO tblالمراحل " & _
          "(معرف_الاستمارة, م, مرحلة_العمل, عدد_القطع, الفني, القرار, السماحية, السبب) VALUES (" & _
          معرف & ", " & رقم & ", '" & نظّف(مرحلة) & "', " & عدد & ", '" & نظّف(فني) & "', '" & _
          نظّف(قرار) & "', '" & نظّف(سماحية) & "', '" & نظّف(سبب) & "')"
    db.Execute sql, dbFailOnError
End Sub

Private Function نظّف(s As String) As String
    نظّف = Replace(s, "'", "''")   ' تأمين علامة التنصيص المفردة
End Function

Private Sub حذف_جدول(db As DAO.Database, اسم As String)
    On Error Resume Next
    db.TableDefs.Delete اسم
    On Error GoTo 0
End Sub

Private Sub حذف_علاقة(db As DAO.Database, اسم As String)
    On Error Resume Next
    db.Relations.Delete اسم
    On Error GoTo 0
End Sub
