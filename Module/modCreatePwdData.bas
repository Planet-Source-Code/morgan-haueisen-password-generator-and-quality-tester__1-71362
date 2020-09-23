Attribute VB_Name = "modCreatePwdData"
Option Explicit

Private CAT As ADOX.Catalog

Private Sub CreateIndexes()

  Dim IDX As ADOX.Index

   On Error GoTo ErrTrap
   ' ===[Create Index 'PrimaryKey']===
   Set IDX = New ADOX.Index

   With IDX
      .Name = "PrimaryKey"
      .Columns.Append "pID"
      .PrimaryKey = True
      .Unique = True
      .Clustered = False
      .IndexNulls = adIndexNullsDisallow
   End With

   CAT.Tables("PwdData").Indexes.Append IDX
   ' ===[Create Index 'pID']===
   Set IDX = New ADOX.Index

   With IDX
      .Name = "pID"
      .Columns.Append "pID"
      .PrimaryKey = False
      .Unique = False
      .Clustered = False
      .IndexNulls = adIndexNullsAllow
   End With

   CAT.Tables("PwdData").Indexes.Append IDX

   Set IDX = Nothing

   Exit Sub

ErrTrap:
   'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateIndexes"
   'Exit Sub
   'Resume

End Sub

Public Sub CreateMDB(ByVal dbPathFilename As String)

   On Error GoTo ErrTrap

   Set CAT = New ADOX.Catalog

   '/* Engine Type = 4; (Access97)
   '/* Engine Type = 5; (Access2000)

   CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPathFilename & ";Jet OLEDB:Database Password=;Jet OLEDB:Engine" & _
      " Type=5;"

   Call CreateTables
   Call CreateIndexes

   Set CAT = Nothing

   '  MsgBox "Database created.", vbApplicationModal + vbInformation, App.Title
   Exit Sub

ErrTrap:
   '  MsgBox Err.Number & " / " & Err.Description
   Exit Sub
   Resume

End Sub

Private Sub CreateTables()

  Dim TBL As ADOX.Table

   On Error GoTo ErrTrap
   ' ===[Create Table 'PwdData']===
   Set TBL = New ADOX.Table
   Set TBL.ParentCatalog = CAT

   With TBL
      .Name = "PwdData"
      .Columns.Append "pID", adInteger, 0
      .Columns("pID").Properties("AutoIncrement") = True
      .Columns("pID").Properties("NullAble") = True

      .Columns.Append "pFor", adVarWChar, 255
      .Columns("pFor").Properties("NullAble") = True
      .Columns("pFor").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "pName", adVarWChar, 255
      .Columns("pName").Properties("NullAble") = True
      .Columns("pName").Properties("Jet OLEDB:Allow Zero Length") = True

      .Columns.Append "pLength", adUnsignedTinyInt, 0
      .Columns("pLength").Properties("NullAble") = True
      .Columns("pLength").Properties("Default") = 0

      .Columns.Append "pUppercase", adUnsignedTinyInt, 0
      .Columns("pUppercase").Properties("NullAble") = True
      .Columns("pUppercase").Properties("Default") = 0

      .Columns.Append "pNumbers", adUnsignedTinyInt, 0
      .Columns("pNumbers").Properties("NullAble") = True
      .Columns("pNumbers").Properties("Default") = 0

      .Columns.Append "pSpecial", adUnsignedTinyInt, 0
      .Columns("pSpecial").Properties("NullAble") = True
      .Columns("pSpecial").Properties("Default") = 0

      .Columns.Append "pFirstNumber", adUnsignedTinyInt, 0
      .Columns("pFirstNumber").Properties("NullAble") = True
      .Columns("pFirstNumber").Properties("Default") = 0

      .Columns.Append "pLastNumber", adUnsignedTinyInt, 0
      .Columns("pLastNumber").Properties("NullAble") = True
      .Columns("pLastNumber").Properties("Default") = 0

   End With

   CAT.Tables.Append TBL

   Set TBL = Nothing

   Exit Sub

ErrTrap:
   'MsgBox Err.Number & " / " & Err.Description,,"Error In CreateTables"
   'Exit Sub
   'Resume

End Sub

