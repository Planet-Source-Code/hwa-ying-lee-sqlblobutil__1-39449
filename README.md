<div align="center">

## sqlBlobUtil


</div>

### Description

1) make blob compress, uncompress, retriving, updating, inserting as simple as normal select

statement.

2) support multiple blob fields in 1 table

support insert, retrive, update from file or memory

3) sample call

RetrieveBlob(ADODB.Connection, sSql, sOutput) As String : retrieve single blob from database,

RetrieveBlobToFile(ADODB.Connection, sSql, blobfiles, sOutput,recordSet) As recordCount

InsertBlobFromFile(ADODB.Connection, sSql, vBlobfiles)

vBlobfiles as array of files or fileName

InsertBlob ADODB.Connection, sSql, vBlobs

vBlobs as string or array of string

UpdateBlobFromFile ADODB.Connection, sSql, vBlobfiles

UpdateBlob ADODB.Connection, sSql, vBlobs

sample input sql clause:

insert into sTableName values(1,"ke,dfs,y",?,?,getdate())

insert into sTableName values(1,'ke,df"s,y',?,?,22-May-1992)

select a, c from sTableName where key = 1

update parameterxml set a='ere',b='?' where d ="dfds"

NOTE!) compress and uncompress are invisible to user

!!!compress tool is not publish with this tool.

(in order to use) comment out the compress part or replace with your own compress tool
 
### More Info
 
adodb.connection, sql Statement(to update, insert, select)

standard sql statement

return blob in files


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[hwa ying lee](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hwa-ying-lee.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hwa-ying-lee-sqlblobutil__1-39449/archive/master.zip)





### Source Code

```
Option Explicit
' *********************************************************************************
'
'  Name:    SqlBlobUtil
'
'  Purpose:  make blob compress, uncompress, retriving, updating, inserting
'        as simple as normal select statement.
'        support multiple blob fields in 1 table
'        support insert, retrive, update from file or memory
'
'        compress and uncompress are invisible to user
'        !!!compress tool is not publish with this tool.
'        (in order to use) comment out the compress part or replace with your own compress tool
'
'        sample input sql clause:
'          insert into sTableName values(1,"ke,dfs,y",?,?,getdate())
'          insert into sTableName values(1,'ke,df"s,y',?,?,22-May-1992)
'          select a, c from sTableName where key = 1
'          update parameterxml set a='ere',b='?' where d ="dfds"
'        interface:
'
'        RetrieveBlob(ADODB.Connection, sSql, sOutput) As String : retrieve single blob from database,
'        RetrieveBlobToFile(ADODB.Connection, sSql, blobfiles, sOutput,recordSet) As recordCount
'        InsertBlobFromFile(ADODB.Connection, sSql, vBlobfile)
'                  vBlobs as array of files or fileName
'        InsertBlob ADODB.Connection, sSql, vBlobs
'                  vBlobs as string or array of string
'        UpdateBlobFromFile ADODB.Connection, sSql, vBlobfiles
'        UpdateBlob ADODB.Connection, sSql, vBlobs
'
'
'  History:  18-Jul-02 hwa ying - created
'     :  Use this at your own risk
'     :  Bug reporting welcome
'     :  hwayinglee@hotmail.com
'
' *********************************************************************************
Public Enum eSQLUtilErrors
  eSqlCannotOpenConnection = vbObjectError + 1300
  eSqlCannotRecogniseSQL
  eSqlNoInputFilename
End Enum
Private cmp As COMPRESSLIBLib.Compress
Private regEx As VBScript_RegExp_55.RegExp
Private Matches As VBScript_RegExp_55.MatchCollection
Private Match As VBScript_RegExp_55.Match
Public Function OpenAdoDbConnection(ByVal sDSN As String, ByVal sDB As String, ByVal sUserId As String, ByVal sUserPwd As String) As ADODB.Connection
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "OpenAdoDbConnection"
  Dim sConString As String
  Dim oConn As ADODB.Connection
  Set oConn = New ADODB.Connection
  sConString = "DRIVER={Sybase System 11};SRVR=" & sDSN & ";DATABASE=" & sDB & ";"
  oConn.Open sConString, sUserId, sUserPwd
  If oConn.State <> adStateOpen Then
    Err.Raise Err.Number, eSQLUtilErrors.eSqlCannotOpenConnection, , " unable to open connection! "
  End If
  Set OpenAdoDbConnection = oConn
Exit_Handler:
  Set oConn = Nothing
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & "ConString=" & sConString & " " & Err.Description
  End If
End Function
' ---------------------------------------------------------------------------------
'  Name:    RetrieveBlobToFile
'
'  Purpose:  retrieve all blob field and save it in file. If no dir specify, output to temp dir
'        select * from blobTable where ..... ( any valid select statement),
'        will output field with image type as blob
'        if more that 1 row retrieved, outfile will be indexed.
'          example: a.xml, a1.xml,a2.xml,b.xml,b1.xml,b2.xml
'        return records affected
'  History:  27-Sep-02 leeh - created
' ---------------------------------------------------------------------------------
Public Function RetrieveBlobToFile(ByRef oConn As ADODB.Connection, _
                  ByVal sSql As String, _
                  ByRef vBlobfile As Variant, _
                  Optional ByRef sFilesName As String = "", _
                  Optional ByRef sOutput As String = "", _
                  Optional ByRef oRs As ADODB.Recordset) As Long
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "RetrieveBlobToFile"
  Dim rs As New ADODB.Recordset
  Dim fso As New Scripting.FileSystemObject
  Dim f As Scripting.TextStream
  Dim i As Integer, fileIdx As Integer, j As Integer
  Dim chunk() As Byte
  Dim sOutFileName As String
  sFilesName = ""
  rs.Open sSql, oConn, adOpenKeyset, adLockOptimistic
  If Not fso.FolderExists(vBlobfile(0)) Then 'if dir is not specified, output to temp dir
    vBlobfile(0) = "c:\temp"
  End If
  sOutFileName = vBlobfile(0)
  ReDim Preserve vBlobfile(rs.Fields.Count)
   j = 0
  For i = 0 To rs.Fields.Count - 1
    If rs.Fields(i).Type = adLongVarBinary Then
       vBlobfile(j) = sOutFileName & "\" & rs.Fields(i).Name & ".xml"
       j = j + 1
    End If
  Next
  fileIdx = 0
  While Not rs.EOF
    j = 0
    For i = 0 To rs.Fields.Count - 1
      If rs.Fields(i).Type = adLongVarBinary Then
        If fileIdx <> 0 Then
          regEx.Pattern = "\."
          sOutFileName = regEx.Replace(CStr(vBlobfile(j)), CStr(fileIdx) & ".")
        Else
          sOutFileName = vBlobfile(j)
        End If
        Set f = fso.CreateTextFile(sOutFileName, True)
        Dim flen As Long
        flen = rs.Fields(i).ActualSize
        If flen > 1 Then
          ReDim chunk(1 To flen)
          chunk() = rs.Fields(i).GetChunk(flen)
          f.Write cmp.UncompressVariant(chunk())
          sOutput = sOutput & cmp.UncompressVariant(chunk()) & vbNewLine
        End If
        f.Close
        sFilesName = sFilesName & IIf(j = 0, sOutFileName, "," & sOutFileName)
        j = j + 1
      End If
    Next
    rs.MoveNext
    fileIdx = fileIdx + 1
    sFilesName = sFilesName & vbCrLf 'carriage return , line feed
  Wend
  RetrieveBlobToFile = fileIdx
  If fileIdx > 0 Then rs.MoveFirst
  Set oRs = rs
Exit_Handler:
  Set rs = Nothing
  Set fso = Nothing
  Set f = Nothing
  If Err.Number <> 0 Then
    Err.Raise 103, sPROC_NAME, TypeName(Me) & " : " & "SQL=" & sSql & " " & Err.Description
  End If
End Function
' ---------------------------------------------------------------------------------
'  Name:    RetrieveBlob
'
'  Purpose:  retrive blobs from db and return as string and recordset
'
'  History:  27-Sep-02 leeh - created
' ---------------------------------------------------------------------------------
Public Function RetrieveBlob(ByRef oConn As ADODB.Connection, _
               ByVal sSql As String, _
               Optional ByRef sOutput As String = vbNullString, _
               Optional ByRef oRs As ADODB.Recordset) As String
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "RetrieveBlob"
  Dim rs As New ADODB.Recordset
  Dim i As Integer
  Dim chunk() As Byte
  Dim flen As Long
  rs.Open sSql, oConn, adOpenKeyset, adLockOptimistic
  While Not rs.EOF
    For i = 0 To rs.Fields.Count - 1
      If rs.Fields(i).Type = adLongVarBinary Then
        flen = rs.Fields(i).ActualSize
        If flen > 1 Then
          ReDim chunk(1 To flen)
          chunk() = rs.Fields(i).GetChunk(flen)
          sOutput = sOutput & cmp.UncompressVariant(chunk()) & vbNewLine
          rs.Fields(i) = sOutput
        End If
      End If
    Next
    sOutput = sOutput & vbNewLine
    rs.MoveNext
  Wend
  If flen > 1 Then rs.MoveFirst
  Set oRs = rs
  RetrieveBlob = sOutput
Exit_Handler:
  rs.Close
  Set rs = Nothing
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & "SQL=" & sSql & " " & Err.Description
  End If
End Function
' ---------------------------------------------------------------------------------
'  Name:    InsertBlob
'
'  Purpose:  insert blob from memory
'        "insert into sTableName values(1,"?","key",?,"22-May-1992"")", arrayBlobs as array of string
'        "insert into sTableName values(1,"?","key",?,"22-May-1992"")", singleBlob as string
'
'  History:  27-Sep-02 leeh - created
Public Function InsertBlob(ByRef oConn As ADODB.Connection, _
               ByVal sSql As String, ByRef vBlobs As Variant)
  insertBlobPrivate oConn, sSql, vBlobs, False
End Function
' ---------------------------------------------------------------------------------
'  Name:    InsertBlobFromFile
'
'  Purpose:  insert blobs from file given connection, valid insert statement, input blob file as array of string or string.
'        "insert into sTableName values(1,"key",?,?,"22-May-1992"")", arrayFileNames
'
'  History:  27-Sep-02 leeh - created
' ---------------------------------------------------------------------------------
Public Sub InsertBlobFromFile(ByRef oConn As ADODB.Connection, _
               ByVal sSql As String, ByRef vBlobfile As Variant)
  insertBlobPrivate oConn, sSql, vBlobfile, True
End Sub
Private Function insertBlobPrivate(ByRef oConn As ADODB.Connection, _
               ByVal sSql As String, ByRef vBlobs As Variant, ByVal blAsFile As Boolean)
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "InsertBlob"
  Dim sTableName As String
  Dim sFields() As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer, j As Integer
  Dim vOutBlob As Variant
  validateAndExtractInsertSql sSql, vBlobs, sTableName, sFields
  rs.Open "select * from " & sTableName & " where 1=0", oConn, adOpenKeyset, adLockOptimistic
  rs.AddNew
  j = 0
  For i = 0 To UBound(sFields)
    vOutBlob = vbNullString
    If rs.Fields(i).Type <> adLongVarBinary Then
      rs.Fields(i) = sFields(i)
    Else
      If sFields(i) = "?" Then
        'the reason for duplicating code below is to avoid string deep copy as vb doesn't support pointer
        If VarType(vBlobs) <> vbString Then 'input as array
          If blAsFile Then
            ReadBlobFromFile CStr(vBlobs(j)), vOutBlob
            rs.Fields(i).AppendChunk cmp.CompressVariant(vOutBlob)
          Else
            rs.Fields(i).AppendChunk cmp.CompressVariant(vBlobs(j))
          End If
        Else
          If blAsFile Then
            ReadBlobFromFile CStr(vBlobs), vOutBlob
            rs.Fields(i).AppendChunk cmp.CompressVariant(vOutBlob)
          Else
            rs.Fields(i).AppendChunk cmp.CompressVariant(vBlobs)
          End If
        End If
        j = j + 1
      End If
    End If
  Next
  rs.Update
  rs.Close
Exit_Handler:
  Set rs = Nothing
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & Err.Description
  End If
End Function
Private Sub Class_Initialize()
  Set cmp = New COMPRESSLIBLib.Compress
  Set regEx = New VBScript_RegExp_55.RegExp
End Sub
Private Sub Class_Terminate()
  Set cmp = Nothing
  Set regEx = Nothing
End Sub
Private Function validateAndExtractInsertSql(ByVal sSql As String, ByRef vBlobfile As Variant, ByRef sTableName, ByRef sFields() As String) As Boolean
Dim vBlobs As Variant
   If VarType(vBlobfile) = vbString Then
    Dim tmp(0) As Variant
    tmp(0) = vBlobfile
    vBlobs = tmp
   Else
    vBlobs = vBlobfile
   End If
   sSql = Replace(sSql, "convert(char(26),getdate(),109)", Format(Date, "mmm-dd-yyyy"))
   sSql = Replace(sSql, "getdate()", Format(Date, "mmm-dd-yyyy"))
   With regEx
    If UBound(vBlobs) = -1 Then
      Err.Raise Err.Number, eSqlNoInputFilename, , "insertBlob " & sSql & " has no input file name"
    End If
    .Pattern = "^\s*insert\s+into\s+(\w+)\s+values\s*\((.*)\)"
    .IgnoreCase = True
    Set Matches = .Execute(sSql)
    If Matches.Count = 1 Then
      sTableName = Matches(0).SubMatches(0)
      sFields = SplitQuoted(Matches(0).SubMatches(1), ",")
    Else
      Err.Raise Err.Number, eSqlCannotRecogniseSQL, , "sql clause unrecognised: sample 'insert into sTableName values(1,'key',?,?,'22-May-1992')"
    End If
  End With
End Function
Private Function validateAndExtractSelectSql(ByVal sSql As String, ByRef sTableName, ByRef sFields As Variant, ByRef sWhereClause As String) As Boolean
  Dim sTableNameAndTheRest As String
  sSql = Replace(sSql, "convert(char(26),getdate(),109)", Format(Date, "mmm-dd-yyyy"))
  sSql = Replace(sSql, "getdate()", Format(Date, "mmm-dd-yyyy"))
  With regEx
   .Pattern = "^\s*select\s+(.*)\s+from\s+(.*)"
   .IgnoreCase = True
   Set Matches = .Execute(sSql)
   If Matches.Count = 1 Then
     sFields = ParseRecValue(Matches(0).SubMatches(0), ",", "=")
     sTableName = Matches(0).SubMatches(1)
   Else
     Err.Raise Err.Number, eSqlCannotRecogniseSQL, , "sql clause unrecognised: sample 'update parameterxml set id='ere',xml='?' where id='we'"
   End If
   .Pattern = "(.*)\s+where\s+(.*)"
   Set Matches = .Execute(sTableName)
   If Matches.Count = 1 Then
    sTableName = Matches(0).SubMatches(0)
    sWhereClause = Matches(0).SubMatches(1)
   End If
  End With
End Function
Public Sub ReadBlobFromFile(ByVal sInFileName As String, ByRef outBlobData As Variant)
On Error GoTo Exit_Handler
Const sPROC_NAME As String = "ReadBlobFromFile"
  ' Open the file for reading
  Dim fso As New Scripting.FileSystemObject
  Dim f As Scripting.File
  Dim ts As Scripting.TextStream
  Dim flen As Long
  Set f = fso.GetFile(sInFileName)
  Set ts = f.OpenAsTextStream(1)
  outBlobData = ""
  Do While ts.AtEndOfStream <> True
    outBlobData = outBlobData & ts.Read(100000)
  Loop
  ts.Close
Exit_Handler:
  Set f = Nothing
  Set ts = Nothing
  Set fso = Nothing
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & "fileName=" & sInFileName & " " & Err.Description
  End If
End Sub
Private Function validateAndExtractUpdateSql(ByVal sSql As String, ByRef vBlobfile As Variant, ByRef sTableName, ByRef sFields As Variant, ByRef sWhereClause As String) As Boolean
Dim sQuote As String
Dim vBlobs As Variant
Dim lDblQuotePos As Long, lSingleQuote As Long
   If VarType(vBlobfile) = vbString Then
    Dim tmp(0) As Variant
    tmp(0) = vBlobfile
    vBlobs = tmp
   Else
    vBlobs = vBlobfile
   End If
   sSql = Replace(sSql, "convert(char(26),getdate(),109)", """" & Format(Date, "mmm-dd-yyyy") & """")
   sSql = Replace(sSql, "getdate()", """" & Format(Date, "mmm-dd-yyyy") & """")
   With regEx
    If UBound(vBlobs) = -1 Then
      Err.Raise Err.Number, eSqlNoInputFilename, , "updateBlob " & sSql & " has no input file name"
    End If
    .Pattern = "^\s*update\s+(\w+)\s+set\s*(.*)"
    .IgnoreCase = True
    Set Matches = .Execute(sSql)
    If Matches.Count = 1 Then
      sTableName = Matches(0).SubMatches(0)
      sFields = Matches(0).SubMatches(1)
    Else
      Err.Raise Err.Number, eSqlCannotRecogniseSQL, , "sql clause unrecognised: sample 'update parameterxml set id='ere',xml='?' where id='we'"
    End If
    .Pattern = "(.*)\s+where\s+(.*)"
    Set Matches = .Execute(sFields)
    If Matches.Count = 1 Then
      sFields = ParseRecValue(Matches(0).SubMatches(0), ",", "=")
      sWhereClause = Matches(0).SubMatches(1)
    Else
      sFields = ParseRecValue(sFields, ",", "=")
      sWhereClause = vbNullString
    End If
  End With
End Function
Public Function UpdateBlobFromFile(ByRef oConn As ADODB.Connection, _
                  ByVal sSql As String, ByRef vBlobs As Variant)
  updateBlobPrivate oConn, sSql, vBlobs, True
End Function
Public Function UpdateBlob(ByRef oConn As ADODB.Connection, _
              ByVal sSql As String, ByRef vBlobs As Variant)
  updateBlobPrivate oConn, sSql, vBlobs, False
End Function
' ---------------------------------------------------------------------------------
'  Name:    updateBlobPrivate
'
'  Purpose:  only update 1 row of record. at the moment
'
'  History:  24-Sep-02 leeh - created
' ---------------------------------------------------------------------------------
Private Function updateBlobPrivate(ByRef oConn As ADODB.Connection, _
                  ByVal sSql As String, ByRef vBlobs As Variant, ByVal blAsFile As Boolean)
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "updateBlob"
  Dim sTableName As String, sWhereClause As String
  Dim sFields As Variant
  Dim rs As New ADODB.Recordset
  Dim i As Integer, j As Integer
  Dim vOutBlob As Variant
  validateAndExtractUpdateSql sSql, vBlobs, sTableName, sFields, sWhereClause
  If sWhereClause = vbNullString Then
    If MsgBox("Are you sure to update the whole table without where clause?", vbYesNo) = vbYes Then
      rs.Open "select * from " & sTableName, oConn, adOpenKeyset, adLockOptimistic
    Else
      Exit Function
    End If
  Else
    rs.Open "select * from " & sTableName & " where " & sWhereClause, oConn, adOpenKeyset, adLockOptimistic
  End If
  While Not rs.EOF
    j = 0
    For i = 0 To UBound(sFields)
      vOutBlob = vbNullString
      If rs.Fields(Trim(CStr(sFields(i)(0)))).Type <> adLongVarBinary Then
        rs.Fields(Trim(CStr(sFields(i)(0)))) = sFields(i)(1)
      Else
        If Trim(CStr(sFields(i)(1))) = "?" Then
        'the reason for duplicating code below is to avoid string deep copy as vb doesn't support pointer
          If VarType(vBlobs) <> vbString Then     'input an array
            If blAsFile Then
              ReadBlobFromFile CStr(vBlobs(j)), vOutBlob
              rs.Fields(Trim(CStr(sFields(i)(0)))).AppendChunk cmp.CompressVariant(vOutBlob)
            Else
              rs.Fields(Trim(CStr(sFields(i)(0)))).AppendChunk cmp.CompressVariant(vBlobs(j))
            End If
          Else
            If blAsFile Then
              ReadBlobFromFile CStr(vBlobs), vOutBlob
              rs.Fields(Trim(CStr(sFields(i)(0)))).AppendChunk cmp.CompressVariant(vOutBlob)
            Else
              rs.Fields(Trim(CStr(sFields(i)(0)))).AppendChunk cmp.CompressVariant(vBlobs)
            End If
          End If
          j = j + 1
        End If
      End If
    Next
    rs.Update
    rs.MoveNext
  Wend
  rs.Close
Exit_Handler:
  Set rs = Nothing
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & Err.Description
  End If
End Function
' ---------------------------------------------------------------------------------
'  Name:    parseRecValue
'
'  Purpose:  return doubleArray
'        sample input "name1='val1,wr',name2='val2,3r'"
'        parseRecValue(input, "," ,"=") will return
'        tmp(0)(0) = "name1"
'        tmp(0)(1) = "val1,wr"
'        tmp(1)(0) = "name2"
'        tmp(1)(1) = "val2,3r"
'
'  History:  20-Aug-02 leeh - created
' ---------------------------------------------------------------------------------
Public Function ParseRecValue(strRecVal, rowDelimeter, colDelimeter)
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "parseRecValue"
Dim tmpRow, tmpCol, noRow, noCol
Dim tmpAllData()
  tmpRow = SplitQuoted(strRecVal, rowDelimeter)
  If UBound(tmpRow) > -1 Then
    ReDim Preserve tmpAllData(UBound(tmpRow))
    For noRow = 0 To UBound(tmpRow)
      tmpAllData(noRow) = SplitQuoted(tmpRow(noRow), colDelimeter)
    Next
    ParseRecValue = tmpAllData
  Else
    ParseRecValue = tmpRow
  End If
Exit_Handler:
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & Err.Description
  End If
End Function
' ---------------------------------------------------------------------------------
'  Name:    SplitQuoted
'
'  Purpose:  get from http://www.vb2themax.com/Item.asp?PageID=CodeBank&ID=187
'        for example you can split the following string into 3 items
'          arr() = SplitQuoted("[one,two],three,[four,five]", , "[]")
'
'  History:  20-Aug-02 enhance by hwaying
' ---------------------------------------------------------------------------------
Public Function SplitQuoted(ByVal Text As String, _
              Optional ByVal Separator As String = ",") As String()
On Error GoTo Exit_Handler
  Const sPROC_NAME As String = "SplitQuoted"
  ReDim res(100) As String
  Dim resCount As Long
  Dim Index As Long
  Dim startIndex As Long
  Dim endIndex As Long
  Dim length As Long
  Dim sepCode As Integer
  Dim bIsSpace As Boolean
  length = Len(Text)
  ' a null string is a special case
  ' return the same uninitialized error that Split would return
  If length = 0 Then
    SplitQuoted = Split(vbNullString)
    Exit Function
  End If
  ' integer ASCII codes of separators
  sepCode = Asc(Separator)
  startIndex = 1
  Index = 0
  endIndex = 0
  Const sSingleQuote = "'"
  Const sDoubleQuote = """"
  bIsSpace = True
  Do While Index < length
    Index = Index + 1
    Select Case Asc(Mid$(Text, Index, 1))
      Case sepCode
        ' we've found the end of an item
        ' if endIndex<>0 then the item is quoted
        If endIndex = 0 Then endIndex = Index
        ' make room in the array, if necessary
        If resCount > UBound(res) Then
          ReDim Preserve res(0 To resCount + 99) As String
        End If
        'store the element
        res(resCount) = Mid$(Text, startIndex, endIndex - startIndex)
        bIsSpace = True
        resCount = resCount + 1
        ' prepare for next element
        startIndex = Index + 1
        endIndex = 0
      Case Asc(sSingleQuote), Asc(sDoubleQuote)
        If Index = 1 Then
          startIndex = Index + 1
        Else
          If Asc(Mid$(Text, Index - 1, 1)) = sepCode Or bIsSpace Then
           startIndex = Index + 1
          End If
        End If
        ' search for the closing quote
        endIndex = InStr(Index + 1, Text, Right$(Mid$(Text, Index, 1), 1))
        If endIndex <> 0 Then
          Index = endIndex
        End If
      Case Asc(" ")
      Case Else
        bIsSpace = False
    End Select
  Loop
  ' store the last item
  If endIndex = 0 Then endIndex = length + 1
  ' trim or expand the array, as necessary
  ReDim Preserve res(0 To resCount) As String
  ' store the element
  res(resCount) = Mid$(Text, startIndex, endIndex - startIndex)
  SplitQuoted = res()
Exit_Handler:
  ' Tidy up code here
  If Err.Number <> 0 Then
    Err.Raise Err.Number, sPROC_NAME, TypeName(Me) & " : " & Err.Description
  End If
End Function
```

