Attribute VB_Name = "Export mdJSON"
Option Compare Database
Public Function printf(mask As String, ParamArray tokens()) As String
    Dim i As Long
    For i = 0 To UBound(tokens)
        mask = Replace$(mask, "{" & i & "}", tokens(i))
    Next
    printf = mask
End Function
''
'Convert full datetime from VB to ISO
'Adapted from http://website.lineone.net/~saphena/iso8601.vbs
'@param vDatetime   Standard VB date/time variant
'@return    Fully qualified ISO datetime string
'
Public Function toIsoDatetime(ByVal vDatetime)
    toIsoDatetime = toIsoDate(vDatetime) & "T" & toIsoTime(vDatetime) & CurrentTimezone
End Function
''
'Extract time part of VB datetime and return only ISO formatted date string; no time or timezone
'@param vDatetime   Standard VB date/time variant
'@return    yyyy-mm-dd
'
Public Function toIsoDate(ByVal vDatetime)

    Dim yy, mm, dd

    yy = Year(vDatetime)
    mm = Month(vDatetime)
    dd = Day(vDatetime)
    toIsoDate = CStr(yy) & "-" & strN2(mm) & "-" & strN2(dd)
    
End Function
''
'Extract time part of VB datetime and return only ISO formatted timestring without timezone appended
'@param vDatetime   Standard VB date/time variant
'@return    hh:mm:ss
'
Public Function toIsoTime(ByVal vDatetime)

    Dim hh, mins, ss

    hh = Hour(vDatetime)
    mins = Minute(vDatetime)
    ss = Second(vDatetime)
    toIsoTime = strN2(hh) & ":" & strN2(mins) & ":" & strN2(ss)
    
End Function
''
'Returns a two digit number
'@param intN
'@return 2 digit string
Public Function strN2(intN)

Dim x

    x = CStr(intN)
    If Len(x) < 2 Then
        x = "0" & x
    End If
    strN2 = x
    
End Function
Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = Replace(Replace(obj.Properties("Description"), vbCrLf, ""), """", "\""") 'Strip line breaks, escape quotes
    
End Function
Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    'http://allenbrowne.com/func-06.html
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Boolean"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function
Sub JsonTablesAndFields()
    'Macro Purpose:  Write all tables, indexes, relationships, and fields to mdJSON file
    'Original Source:  vbaexpress.com/kb/getarticle.php?kb_id=707
    'Updates by Josh Bradley, jbradley@arcticlcc.org
    
    Dim lTbl As Long
    Dim lFld As Long
    Dim dBase As Database
    Dim fso As Object
    Dim path As String
    Dim oFile As Object
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim strGUID As String
    Dim json As String
    Dim arr() As String
    Dim arr2() As String
    Dim tbl As Object
    Dim idx As DAO.Index
    Dim idxArr As String
    Dim i As Integer
    Dim ii As Integer
    
    'Set current database to a variable and create a new FileSystemObject instance
    Set dBase = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = CurrentProject.path & "\dataDictionary_mdJSON" & Application.CurrentProject.Name & ".json"
    Set oFile = fso.CreateTextFile(path)
    
    'Set on error in case there are no tables
    On Error Resume Next
    
    oFile.WriteLine printf("{""dictionaryInfo"": {""citation"": {""title"": ""{0}"", ""date"": [{""date"": ""{1}"", ""dateType"": ""creation""}]},", _
         "Data Dictionary for " & Application.CurrentProject.Name & ", " & dBase.Containers!Databases.Documents!SummaryInfo.Properties("Title") _
            & ", MSAccess Version " & Application.Version, _
         toIsoDatetime(Now))
    oFile.WriteLine """description"": ""This data dictionary was automatically generated from database via Visual Basic script."","
    oFile.WriteLine """resourceType"": ""database"",""language"": ""eng; US""},"
    oFile.WriteLine """entity"":["
    
    'Loop through all tables
    For lTbl = 0 To dBase.TableDefs.Count
    'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).Name, 1) = "~" Or _
        Left(dBase.TableDefs(lTbl).Name, 4) = "MSYS" Then
        '~ indicates a temporary table
        'MSYS indicates a system level table
        Else
            'Otherwise, loop through each table, writing the table and field names
            'to mdJSON
            Set tbl = dBase.TableDefs(lTbl)
            strGUID = Mid(StringFromGUID(tbl.Properties("GUID").Value), 8, 36)
            
            'Write Table Info
            oFile.WriteLine printf("{""entityId"": ""{0}"",""commonName"": """",""codeName"": ""{1}"",""definition"": """"", strGUID, tbl.Name)
            
            'Write primary key
            idxArr = ""
            i = 0
            For Each idx In tbl.Indexes
                ReDim arr(idx.Fields.Count - 1)
                For lFld = 0 To idx.Fields.Count - 1
                    arr(lFld) = idx.Fields(lFld).Name
                Next
                If idx.Primary Then
                    oFile.WriteLine ",""primaryKeyAttributeCodeName"": [""" & Join(arr, ", ") & """]"
                Else
                    If i > 0 Then
                        idxArr = idxArr & ","
                    End If
                    idxArr = idxArr & printf("{""codeName"": ""{0}"",""allowDuplicates"": {1},""attributeCodeName"": [""{2}""]}", idx.Name, StrConv(Not idx.Unique, vbLowerCase), Join(arr, ", "))
                    i = i + 1
                End If
            Next
            
            'Write indexes
            If i > 0 Then
                oFile.WriteLine ",""index"": [" & idxArr & "]"
            End If
            
            'Write attributes
            oFile.WriteLine ",""attribute"": ["
            
            For lFld = 0 To tbl.Fields.Count - 1
                If lFld > 0 Then
                    oFile.WriteLine ","
                End If
                
                Set fld = tbl.Fields(lFld)
                oFile.WriteLine printf("{""commonName"": ""{0}"",""codeName"": ""{1}"",""definition"": ""{2}"",""dataType"": ""{3}"",""allowNull"": {4}}", fld.Name, _
                    fld.Name, GetDescrip(fld), FieldTypeName(fld), StrConv(Not fld.Required, vbLowerCase))
            Next lFld
            
            'Close attribute
            oFile.WriteLine "],"
            
            oFile.WriteLine """foreignKey"": ["
            i = 0
            For Each rel In dBase.Relations
                If rel.ForeignTable = tbl.Name Then
                    ii = 0
                    ReDim arr(rel.Fields.Count - 1)
                    ReDim arr2(rel.Fields.Count - 1)
                    
                    For Each fld In rel.Fields
                        arr(ii) = fld.Name
                        arr2(ii) = fld.ForeignName
                        ii = ii + 1
                    Next
                    
                    If i > 0 Then
                        oFile.WriteLine ","
                    End If
                    
                    oFile.WriteLine printf("{""localAttributeCodeName"": [""{0}""],""referencedEntityCodeName"": ""{1}"",""referencedAttributeCodeName"": [""{2}""]}", _
                        Join(arr2, ", "), rel.Table, Join(arr, ", "))
                    
                    i = i + 1
                End If
            Next
            oFile.WriteLine "]" 'Close FK
            
            'Close Entity
            If lTbl = dBase.TableDefs.Count - 1 Then
                oFile.WriteLine "}"
            Else
                oFile.WriteLine "},"
            End If
        End If
        
        Next lTbl
        
        oFile.WriteLine "]}"
        
        'Resume error breaks
        On Error GoTo 0
        
        oFile.Close
        'Set release fso from memory
        Set fso = Nothing
        Set oFile = Nothing
        
        'Release database object from memory
        Set dBase = Nothing
        
    End Sub
