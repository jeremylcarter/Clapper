<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' SQL Server Dynamic ASP Classic SQL Connnection
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright 2012 Jeremy Child  jeremychild@gmail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
'to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
'and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
'WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Uses additional libraries (Remove if required): 
' Js Linq http://jslinq.codeplex.com/
' Squel.Js http://hiddentao.github.com/squel/
' Fluent Html Builder http://jshtmlbuilder.codeplex.com/

' SQL Server Connection class

Const SqlDateTimeFormatCode = 103

Class SqlServerConnection
	
	Private m_Trace
	Private m_ConnectionString
	Private m_Connection
	
	Public Sub Class_Initialize
	  m_Trace = False
	  m_ConnectionString = "default connection string here"
	End Sub 
	
	Public Sub TraceSql(inputSql)
		If m_Trace = True Then
			Response.Write("<div style=""background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;"">" & inputSql & "</div>")
		End If 
		
	End Sub
	
	Public Function DataTypeNameFromAdoCode(code) 
		
		If IsNumeric(code) Then

			Select Case Cint(code)
				Case 205
					DataTypeNameFromAdoCode = "Binary"
				Case 11
					DataTypeNameFromAdoCode = "Boolean"
				Case 3
					DataTypeNameFromAdoCode = "Int32"	
				Case 4
					DataTypeNameFromAdoCode = "Int32"						
				Case 5
					DataTypeNameFromAdoCode = "Float"	
				Case 6
					DataTypeNameFromAdoCode = "Float"						
				Case 14
					DataTypeNameFromAdoCode = "Float"		
				Case 202
					DataTypeNameFromAdoCode = "String"								
				Case 135
					DataTypeNameFromAdoCode = "DateTime"	
				Case 7
					DataTypeNameFromAdoCode = "DateTime"												
				Case 131
					DataTypeNameFromAdoCode = "Decimal"					
				Case 16
					DataTypeNameFromAdoCode = "Int32"	
				Case 2
					DataTypeNameFromAdoCode = "Int32"	
				Case 20
					DataTypeNameFromAdoCode = "Int64"																							
				Case Else
					DataTypeNameFromAdoCode = "String"
			End Select
		Else
			DataTypeNameFromAdoCode = "String"
		End If
		
	End Function
	
	Public Sub TraceRecordset(inputRs)
		
		On Error Resume Next
				
		If m_Trace = True Then
					
				Dim i, j, colspan 
				colspan = inputRs.Fields.Count
				Response.Write "<TABLE style=""border: 2px solid #17b;	margin: 0.3em 0.2em;""><TR>"
				Response.Write "<tr><TD colspan=""" & colspan & """ style=""color: white; background-color: #0066FF""><B>Dump of Recordset</B></TD><tr>"
				For i = 0 To inputRs.Fields.Count - 1
					Response.Write "<TD style=""color: black; background-color: #ddd""><B>" & inputRs.Fields(i).Name & "</B><font color=""grey"">&lt;" & DataTypeNameFromAdoCode(inputRs.Fields(i).Type) & "&gt;</font></TD>"
				Next
				Response.Write "</TR>"
				i = 1
				Do While Not inputRs.EOF 
					Response.Write "<TR"
					If i mod 2 = 0 Then
					 Response.Write "BGCOLOR=""#d3d3d3"""
					End If 
					Response.Write ">"
			
					For j = 0 To inputRs.Fields.Count-1
						If inputRs.Fields(i).Type <> 205 Then
						
							Dim val : val = "&nbsp;"
							If Not IsNull(inputRs.Fields(j).Value) Then
								val = Trim(Cstr(inputRs.Fields(j).Value))
								If Len(val) = 0 Then
									val = "&nbsp;&nbsp;"
								End If
							Else
								val = "<font color=""green""><i>null</i></font>"
							End If
						
							Response.Write "<TD style=""vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"">" & val & "</TD>"
						Else
							Response.Write "<TD style=""vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"">&nbsp;&nbsp;</TD>"
						End If
			
					Next
					Response.Write "</TR>"
					inputRs.MoveNext
					i=i+1
				Loop
				
				If i > 1 Then
					inputRs.MoveFirst 
				End If
				
				Response.Write "</TABLE>"
			End If
			
		On Error Goto 0
	
	End Sub
	
	Public Property Let Trace(enable)
		m_Trace = enable
	End Property
	
	Public Property Let ConnectionString(strIn)
		m_ConnectionString = strIn
	End Property
	
	Public LastSqlStatement
	
	Public Function Open()
	
	 If IsObject(m_Connection) = False Then
		Set m_Connection = Server.CreateObject("ADODB.Connection")
	 End If
	
	 If m_Connection.State <> 1 Then
			m_Connection.Open m_ConnectionString
			m_Connection.CursorLocation=3
	 End If
	
	End Function
	
	Public Function SanitizeString(ByVal input)
		If input = "" or IsNull(input) Then Exit Function
		SanitizeString = Replace(input, "'", "''")
	End Function 
	
	Public Function Close()
	
	 If m_Connection.State = 1 Then
			m_Connection.Close
	 End If
	
	End Function
	
	Public Function Execute(Byval sql)
	
	    LastSqlStatement = sql
	    Execute = Exec(sql)
	
	End Function
	
	Public Function Exec(ByVal sql)
	
	    LastSqlStatement = sql
	    
		TraceSql(LastSqlStatement)
		
		Dim returnValue : returnValue = False
		
		Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
		cmd.ActiveConnection = m_ConnectionString
		cmd.ActiveConnection.BeginTrans
		cmd.CommandText = sql
		cmd.CommandType = 1
		
		cmd.Execute True, , 128
		cmd.ActiveConnection.CommitTrans
		
		If Err <> 0 Then
		    returnValue = False
			cmd.ActiveConnection.RollBackTrans
			Err.Raise  Err.Number, "SqlConnection: Exec", Err.Description
		Else
			returnValue = True
		End If
	
		Exec = returnValue
	
	End Function
	
	Function Insert(object,table,primaryKey)
		
		Dim sql: sql = InsertSql(object, table, primaryKey)
	    LastSqlStatement = sql
	    
		TraceSql(LastSqlStatement)
				
		Dim cmd : Set cmd = Server.CreateObject("adodb.Command")

		Dim recAffected : recAffected = 0

		cmd.ActiveConnection  = m_ConnectionString
		cmd.ActiveConnection.BeginTrans
		cmd.CommandText = sql
		cmd.CommandType = 1

		cmd.Execute true, recaffected, 128
		cmd.ActiveConnection.CommitTrans

		If Err <> 0 Then
			cmd.ActiveConnection.RollBackTrans
			Err.Raise  Err.Number, "SqlConnection: Insert", Err.Description
		End If

		Update = recAffected
		
	End Function
	
	Function Update(object,table,primaryKey)
		
		Dim sql: sql = UpdateSql(object, table, primaryKey)
	    LastSqlStatement = sql
	    
		TraceSql(LastSqlStatement)
				
		Dim cmd : Set cmd = Server.CreateObject("adodb.Command")

		Dim recAffected : recAffected = 0

		cmd.ActiveConnection  = m_ConnectionString
		cmd.ActiveConnection.BeginTrans
		cmd.CommandText = sql
		cmd.CommandType = 1

		cmd.Execute true, recaffected, 128
		cmd.ActiveConnection.CommitTrans

		If Err <> 0 Then
			cmd.ActiveConnection.RollBackTrans
			Err.Raise  Err.Number, "SqlConnection: Update", Err.Description
		End If

		Update = recAffected
		
	End Function
	
	Function Delete(Byval table, Byval primaryKey, Byval key)
	
	End Function
	
    Function ExecReturnRowsAffected(ByVal sql)
	
	    LastSqlStatement = sql
		
		TraceSql(LastSqlStatement)
	    
		Dim cmd : Set cmd = Server.CreateObject("adodb.Command")

		Dim recAffected : recAffected = 0

		cmd.ActiveConnection  = m_ConnectionString
		cmd.ActiveConnection.BeginTrans
		cmd.CommandText = sql
		cmd.CommandType = 1

		cmd.Execute true, recaffected, 128
		cmd.ActiveConnection.CommitTrans

		If Err <> 0 then
			cmd.ActiveConnection.RollBackTrans
			Err.Raise  Err.Number, "SqlConnection: ExecReturnRowsAffected", Err.Description
		End If

		ExecReturnRowsAffected = recAffected
		
	End Function
	
	Public Function BinaryStream(sql) 
	
	    Dim stream
        Set stream = Server.CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Position = 0
        
        LastSqlStatement = sql
		TraceSql(LastSqlStatement)
	    		
		Set conn=Server.CreateObject("ADODB.Connection")
		conn.Open dbConnectionString
				
		Dim rs : Set rs=conn.Execute(sql)
				
		If Not IsNull(rs(0)) Then
		    stream.Write rs.Fields(0).Value
		End If
        
        stream.Position = 0
        stream.Close
        
        BinaryStream = stream

	End Function
	
	Public Function RecordSet(sql)
		
		LastSqlStatement = sql
		TraceSql(LastSqlStatement)
		    
		Dim rs, cmd
	
		Set rs = Server.CreateObject("ADODB.Recordset")
		Set cmd = Server.CreateObject("ADODB.Command")
		
		cmd.ActiveConnection = m_ConnectionString
		cmd.CommandText = sql
		cmd.CommandType = 1
		cmd.Prepared = True
		
		rs.CursorLocation = 3
		rs.Open cmd, , 0, 1
		
		If Err <> 0 Then
			Err.Raise  Err.Number, "SqlConnection: RecordSet", Err.Description
		End If
		
		TraceRecordset(rs)
		
		Set RecordSet = rs
							
	End Function

	Public Function ExecStoredProcedure(Byval name,Byval params)
		
		Dim sql : sql = StoredProcedureSql(name,params)
		
		Set rs = RecordSet(sql)
		
		Dim list
		Set list = RecordSetToList(rs)
		
		Set ExecStoredProcedure = list

	End Function
	
	Function ExecStoredProcedureIdentity(Byval name,Byval params)
	
		Dim sql : sql = StoredProcedureSql(name,params)
		LastSqlStatement = sql
		TraceSql(LastSqlStatement)
		
		Dim returnValue : returnValue = 0
		
		Set conn=Server.CreateObject("ADODB.Connection")
		conn.Open m_ConnectionString
		
		sql = sql & ";SELECT @@IDENTITY AS NewID" 
		
		Dim rs : Set rs=conn.Execute(sql)
		
		Dim rs2 : Set rs2 = rs.NextRecordSet()
		
		If Not IsNull(rs2(0)) Then
			returnValue = CDbl(rs2(0).value) 
		Else
			returnValue = 0
		End If
		
		conn.Close
										
		ExecStoredProcedureIdentity = returnValue

	End Function
	
	Public Function ExecStoredProcedureToInt(Byval name,Byval params)
		
		Dim sql : sql = StoredProcedureSql(name,params)
		LastSqlStatement = sql
		TraceSql(LastSqlStatement)
		
		Set rs = RecordSet(sql)
				
		ExecStoredProcedureToInt = RecordSetToInt(rs)
		
	End Function
	
	Public Function ExecStoredProcedureSingle(Byval name,Byval params)
		
		Dim sql : sql = StoredProcedureSql(name,params)
        
		Dim rs, cmd
			
		If Len(sql) > 1 Then
		
			Set rs = RecordSet(sql)
					
			Dim list
			Set list = RecordSetToSingle(rs)
							
			If Err <> 0 Then
				Err.Raise  Err.Number, "SqlConnection: ExecStoredProcedureSingle", Err.Description
			End If
									
			Set ExecStoredProcedureSingle = list
			
		End If
		
	End Function
			  
	Public Function QueryToList(sql)

		Set rs = RecordSet(sql)
								
		Dim list
		Set list = RecordSetToList(rs)
		
		Set QueryToList = list
		
	End Function
	
	Public Function Query(sql)
		
		Set rs = RecordSet(sql)
				
		Dim list
		Set list = RecordSetToArray(rs)
		
		Set Query = list
		
	End Function
	
	Public Function QueryToEnumerable(sql)
		
		Set rs = RecordSet(sql)
		
		Dim list
		Set list = Enumerable.From(rs).Select()
		
		Set QueryToEnumerable = list
		
	End Function	
	
	Public Function QuerySingle(sql)
		
		Set rs = RecordSet(sql)
		
		Dim obj
		Set obj = RecordSetToSingle(rs)
				
		Set QuerySingle = obj
		
	End Function
	
	Public Function QueryToString(sql)
	
	    Set rs = RecordSet(sql)

		QueryToString = RecordSetToString(rs)
	
	End Function
	
	Public Function QueryDateTime(sql)
	
	    Set rs = RecordSet(sql)

		QueryDateTime = RecordSetToDateTime(rs)
	
	End Function
	
	Public Function QueryBoolean(sql)
		
		Set rs = RecordSet(sql)
		
		QueryBoolean = RecordSetToBoolean(rs)
		
	End Function
	
	Public Function QueryInt(sql)
		
		Set rs = RecordSet(sql)

		QueryInt = RecordSetToInt(rs)
		
	End Function
	
	Public Function QueryToArray(sql)
		
		Set rs = RecordSet(sql)
		
		Dim list
		Set list = RecordSetToArray(rs)
		
		Set QueryToArray = list
		
	End Function
	
	Public Function QueryList(sql)
	
	    Set rs = RecordSet(sql)
								
		Dim list
		Set list = RecordSetToList(rs)
		
		Set QueryToList = list
		
	End Function
	
	Function ExecuteReturnIdentity(Byval sql)
	
	    LastSqlStatement = sql
		TraceSql(LastSqlStatement)
	    
		Dim returnValue : returnValue = 0
		
		Set conn=Server.CreateObject("ADODB.Connection")
		conn.Open m_ConnectionString
		
		sql = sql & ";SELECT @@IDENTITY AS NewID" 
		
		Dim rs : Set rs=conn.Execute(sql)
		
		Dim rs2 : Set rs2 = rs.NextRecordSet()
		
		If Not IsNull(rs2(0)) Then
			returnValue = Cint(rs2(0).value) 
		Else
			returnValue = 0
		End If
						
		ExecuteReturnIdentity = returnValue

	End Function
		
	' Helper functions
	Function IsValidGuid(byval strGUID)
		If IsNull(strGUID) Then
		isGUID = false
		Exit Function
		End If
		Dim regEx
		Set regEx = New RegExp
		regEx.Pattern = "[0-9A-Fa-f-]+"
		IsValidGuid = regEx.Test(strGUID)
		Set RegEx = Nothing
	End Function
	
	Function GetCleanGuid() 
		GetCleanGuid = Replace(Replace(Left(CStr(GetGuid()), 38),"{",""),"}","")
	End Function 

	Function GetNewGuid() 
        Set TypeLib = CreateObject("Scriptlet.TypeLib") 
        GetNewGuid = Left(CStr(TypeLib.Guid), 38)
        Set TypeLib = Nothing 
	End Function 
	
	Public Function ToFloat(value, alternative)
		val = value & ""
		ToFloat = alternative
		if isNumeric(val) and val <> "" then ToFloat = cdbl(val)
	End Function
	
	Public Function ToInt(value, alternative)
			val = value & ""
			ToInt = alternative
			if isNumeric(val) and val <> "" then ToInt = cInt(val)
	End Function
			
	Function CleanString(Byval s)

		BlackList = Array("--", ";", "/*", "*/", "@@", "@",_
						  "char", "nchar", "varchar", "nvarchar",_
						  "alter", "begin", "cast", "create", "insert",_
						  "declare", "delete", "drop", "end", "exec",_
						  "table", "update")
	
		For Each i in BlackList
		  s = Replace(s,i,"")  
		Next
	
		s = Replace(s,"'","''")
	
		CleanString = s

	End Function
	
	Function BoolToYesNo(byval bool)
		
		If IsNull(bool) Then
			BoolToYesNo = "No"
		End If
		If Cbool(bool) = True Then
			BoolToYesNo = "Yes"
		Else
			BoolToYesNo = "No"
		End If
	
	End Function
	
	Function IsEmailAddress(myEmail)
		Dim isValidE
		Dim regEx
		isValidE = True
		Set regEx = New RegExp
		regEx.IgnoreCase = False
		regEx.Pattern = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
		isValidE = regEx.Test(myEmail)
		IsEmailAddress = isValidE
	End Function
  
End Class

Function RandomNumberString() 
	intHighNumber = 100
	intLowNumber = 1
	numberString = ""
	
	For i = 1 to 4
		Randomize
		intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
		numberString = numberString + CStr(intNumber)
	Next
	RandomNumberString = numberString 
End Function

Function WriteFormValueOrDefault(Byval formValue, Byval defaultValue)

    Response.Write(FormValueOrDefault(formValue,defaultValue))

End Function

Function WriteQueryStringValueOrDefault(Byval formValue, Byval defaultValue)

    Response.Write(QueryStringValueOrDefault(formValue,defaultValue))

End Function

Function QueryStringValueOrDefault(Byval formValue, Byval defaultValue)

    Dim returnValue

    If (Len(Request.QueryString(formValue))> 0) Then
        ' The form value has something
       returnValue = (Request.QueryString(formValue))
    Else
        ' The form value does not have any data and thus the default will be used
        returnValue = (defaultValue)
    End If
    ' For no reason return the form value.
    QueryStringValueOrDefault = returnValue

End Function

Function FormValueOrDefault(Byval formValue, Byval defaultValue)

    Dim returnValue

    If (Len(Request.Form(formValue))> 0) Then
        ' The form value has something
       returnValue = (Request.Form(formValue))
    Else
        ' The form value does not have any data and thus the default will be used
        returnValue = (defaultValue)
    End If
    ' For no reason return the form value.
    FormValueOrDefault = returnValue

End Function

Function VbScriptTypeName(object)
	VbScriptTypeName = TypeName(object)
End Function

Function ToSQLDateTime(value)
	ToSQLDateTime = "CONVERT(DATETIME,'" & FormatDateTime(value,vbGeneralDate) & "'," & SqlDateTimeFormatCode & ") "
End Function

Public Function IsVbDate(value) 
	IsVbDate = IsDate(value)
End Function

Public Function ToVbDateTime(value) 
	ToVbDateTime = CDate(value)
End Function

Public Function ToVbBoolean(value)
		ToVbBoolean = Cbool(value)
End Function

Public Function IsVbNumeric(value) 
	IsVbNumeric = IsNumeric(value)
End Function

Public Function IsVbBoolean(value)
	IsVbBoolean = (LCase(value) = "true" Or LCase(value) = "false")
End Function
%>
<script runat="SERVER" language="JSCRIPT">

    // Load from form/querystring into object properties

    function toTitleCase(toTransform) {
        return toTransform.replace(/\b([a-z])/g, function(_, initial) {
            return initial.toUpperCase();
        });
    }

    function LoadFromFormVariables(a) {


        var propertiesOfT = GetProperties(a).join();

        for (f = new Enumerator(Request.Form()); !f.atEnd(); f.moveNext()) {
            var key = toTitleCase(f.item());
            var val = Request.Form(f.item());

            if (propertiesOfT.indexOf(key, 0) > -1) {
				
                try {
                    if (String(val).charAt(0) == '0') {
                        eval("a." + key + " = \"" + val + "\";");
                    } else if (IsVbNumeric(val)) {
                        eval("a." + key + " = " + val + " ;");
                    } else if (IsVbDate(val)) {
                        eval("a." + key + " = ToVbDateTime(\"" + val + "\") ;");
                    } else if (IsVbBoolean(val)) {
                        eval("a." + key + " = ToVbBoolean(\"" + val + "\") ;");
                    } else {
                        eval("a." + key + " = \"" + val + "\" ;");
                    }
                } catch (ex) {
                }
            }   
			if (propertiesOfT.indexOf(f.item(), 0) > -1) {
				
                try {
                    if (String(val).charAt(0) == '0') {
                        eval("a." + f.item() + " = \"" + val + "\";");
                    } else if (IsVbNumeric(val)) {
                        eval("a." + f.item() + " = " + val + " ;");
                    } else if (IsVbDate(val)) {
                        eval("a." + f.item() + " = ToVbDateTime(\"" + val + "\") ;");
                    } else if (IsVbBoolean(val)) {
                        eval("a." + f.item() + " = ToVbBoolean(\"" + val + "\") ;");
                    } else {
                        eval("a." + f.item() + " = \"" + val + "\" ;");
                    }
                } catch (ex) {
                }
            }   			
        }

    }
    function LoadFromQueryStringVariables(a) {

        try {
            formItem = new Array();
            formItemIndex = 0;

            for (e = new Enumerator(Request.QueryString); !e.atEnd(); e.moveNext()) {

                formItem[formItemIndex] = Request.QueryString(e.item());

                if (formItem[formItemIndex] != "") {
                    // attempt to map to a possible property of a 	
                    var propertyExists = eval('typeof a.' + e.item() + ';');
                    if (propertyExists != null) {
                        try {

                            var val = formItem[formItemIndex];

                            if (String(val).charAt(0) == '0') {
                                eval("a." + e.item() + " = '" + val + "' ;");
                            } else if (IsVbNumeric(val)) {
                                eval("a." + e.item() + " = " + val + " ;");
                            } else if (IsVbDate(val)) {
                                eval("a." + e.item() + " = ToVbDateTime('" + val + "') ;");
                            } else if (IsVbBoolean(val)) {
                                eval("a." + e.item() + " = ToVbBoolean('" + val + "') ;");
                            } else {
                                eval("a." + e.item() + " = '" + val + "' ;");
                            }
                        } catch (ex) {
                        }
                    }
					
                }
                formItemIndex++;
            }
        } catch (ex) { }

    }

    // Reflection
    function GetMethods(a) { var b = new Array; for (var c in a) { if (typeof a[c] == "function") { b[b.length] = c } } return b }
    function GetProperties(a) { var b = new Array; for (var c in a) { if (typeof a[c] != "function") { b[b.length] = c } } return b }

    function ExpandoObject(a) {
        var b = new Object; var c = (new VBArray(a.Keys())).toArray(); var d = (new VBArray(a.Items())).toArray(); for (i in d) {
            switch (d[i]) {
                case "Date": b[c[i]] = (new Date).getVarDate(); break; case "String": b[c[i]] = ""; break; case "Integer": b[c[i]] = 0; break; case "Boolean": b[c[i]] = false; break; default: b[c[i]] = ""; break
            } 
        } return b
    }

    function FlattenObject(a) { var b = ""; for (property in a) { b += property + ": " + a[property] + ";" + typeof a[property] + " <br /> " } return b }

    function DictionaryFlatten(a) {
        var b = (new VBArray(a.Keys())).toArray(); var c = (new VBArray(a.Items())).toArray(); for (key in b) { Response.Write(b[key]); Response.Write(c[key]) }
    }

    // Javascript Fluent SQL Builder (SQUEL)
    (function() {
        var DefaultInsertBuilderOptions, DefaultUpdateBuilderOptions, Delete, Expression, ExpressionClassName, Insert, Select, Update, WhereOrderLimit, formatValue, getObjectClassName, sanitizeAlias, sanitizeCondition, sanitizeField, sanitizeLimitOffset, sanitizeName, sanitizeTable, sanitizeValue, _export, _extend, __slice = Array.prototype.slice, __hasProp = Object.prototype.hasOwnProperty, __bind = function(a, b) { return function() { return a.apply(b, arguments) } }, __extends = function(a, b) { function d() { this.constructor = a } for (var c in b) { if (__hasProp.call(b, c)) a[c] = b[c] } d.prototype = b.prototype; a.prototype = new d; a.__super__ = b.prototype; return a }; _extend = function() { var a, b, c, d, e, f, g; a = arguments[0], c = 2 <= arguments.length ? __slice.call(arguments, 1) : []; if (c) { for (f = 0, g = c.length; f < g; f++) { d = c[f]; if (d) { for (b in d) { if (!__hasProp.call(d, b)) continue; e = d[b]; a[b] = e } } } } return a }; Expression = function() { function b() { this.toString = __bind(this.toString, this); this.or = __bind(this.or, this); this.and = __bind(this.and, this); this.end = __bind(this.end, this); this.or_begin = __bind(this.or_begin, this); this.and_begin = __bind(this.and_begin, this); var a = this; this.tree = { parent: null, nodes: [] }; this.current = this.tree; this._begin = function(b) { var c; c = { type: b, parent: a.current, nodes: [] }; a.current.nodes.push(c); a.current = a.current.nodes[a.current.nodes.length - 1]; return a } } var a; b.prototype.tree = null; b.prototype.current = null; b.prototype.and_begin = function() { return this._begin("AND") }; b.prototype.or_begin = function() { return this._begin("OR") }; b.prototype.end = function() { if (!this.current.parent) throw new Error("begin() needs to be called"); this.current = this.current.parent; return this }; b.prototype.and = function(a) { if (!a || "string" !== typeof a) { throw new Error("expr must be a string") } this.current.nodes.push({ type: "AND", expr: a }); return this }; b.prototype.or = function(a) { if (!a || "string" !== typeof a) { throw new Error("expr must be a string") } this.current.nodes.push({ type: "OR", expr: a }); return this }; b.prototype.toString = function() { if (null !== this.current.parent) { throw new Error("end() needs to be called") } return a(this.tree) }; a = function(b) { var c, d, e, f, g, h; e = ""; h = b.nodes; for (f = 0, g = h.length; f < g; f++) { c = h[f]; if (c.expr != null) { d = c.expr } else { d = a(c); if ("" !== d) d = "(" + d + ")" } if ("" !== d) { if ("" !== e) e += " " + c.type + " "; e += d } } return e }; return b } (); DefaultInsertBuilderOptions = DefaultUpdateBuilderOptions = { usingValuePlaceholders: false }; getObjectClassName = function(a) { var b; if (a && a.constructor && a.constructor.toString) { b = a.constructor.toString().match(/function\s*(\w+)/); if (b && b.length === 2) return b[1] } }; ExpressionClassName = getObjectClassName(new Expression); sanitizeCondition = function(a) { var b; b = typeof a; if (ExpressionClassName !== getObjectClassName(a) && "string" !== b) { throw new Error("condition must be a string or Expression instance") } if ("Expression" === b) a = a.toString(); return a }; sanitizeName = function(a, b) { if ("string" !== typeof a) { throw new Error("" + b + " must be a string") } return a }; sanitizeField = function(a) { return sanitizeName(a, "field name") }; sanitizeTable = function(a) { return sanitizeName(a, "table name") }; sanitizeAlias = function(a) { return sanitizeName(a, "alias") }; sanitizeLimitOffset = function(a) { a = parseInt(a); if (0 > a) throw new Error("limit/offset must be >=0"); return a }; sanitizeValue = function(a) { var b; b = typeof a; if (null !== a && "string" !== b && "number" !== b && "boolean" !== b) { throw new Error("field value must be a string, number, boolean or null") } return a }; formatValue = function(a, b) { if (null === a) { a = "NULL" } else if ("boolean" === typeof a) { a = a ? "TRUE" : "FALSE" } else if ("number" !== typeof a) { if (false === b.usingValuePlaceholders) a = '"' + a + '"' } return a }; WhereOrderLimit = function() { function a() { this.limitString = __bind(this.limitString, this); this.orderString = __bind(this.orderString, this); this.whereString = __bind(this.whereString, this); this.limit = __bind(this.limit, this); this.order = __bind(this.order, this); this.where = __bind(this.where, this); this.wheres = []; this.orders = [] } a.prototype.wheres = null; a.prototype.orders = null; a.prototype.limits = null; a.prototype.where = function(a) { a = sanitizeCondition(a); if ("" !== a) this.wheres.push(a); return this }; a.prototype.order = function(a, b) { if (b == null) b = true; a = sanitizeField(a); this.orders.push({ field: a, dir: b ? "ASC" : "DESC" }); return this }; a.prototype.limit = function(a) { a = sanitizeLimitOffset(a); this.limits = a; return this }; a.prototype.whereString = function() { if (0 < this.wheres.length) { return " WHERE (" + this.wheres.join(") AND (") + ")" } else { return "" } }; a.prototype.orderString = function() { var a, b, c, d, e; if (0 < this.orders.length) { b = ""; e = this.orders; for (c = 0, d = e.length; c < d; c++) { a = e[c]; if ("" !== b) b += ", "; b += "" + a.field + " " + a.dir } return " ORDER BY " + b } else { return "" } }; a.prototype.limitString = function() { if (this.limits) { return " LIMIT " + this.limits } else { return "" } }; return a } (); Select = function(a) { function b() { this.toString = __bind(this.toString, this); this.offset = __bind(this.offset, this); this.group = __bind(this.group, this); this.outer_join = __bind(this.outer_join, this); this.right_join = __bind(this.right_join, this); this.left_join = __bind(this.left_join, this); this.join = __bind(this.join, this); this.field = __bind(this.field, this); this.from = __bind(this.from, this); this.distinct = __bind(this.distinct, this); var a = this; b.__super__.constructor.apply(this, arguments); this.froms = []; this.fields = []; this.joins = []; this.groups = []; this._join = function(b, c, d, e) { c = sanitizeTable(c); if (d) d = sanitizeAlias(d); if (e) e = sanitizeCondition(e); a.joins.push({ type: b, table: c, alias: d, condition: e }); return a } } __extends(b, a); b.prototype.froms = null; b.prototype.fields = null; b.prototype.joins = null; b.prototype.groups = null; b.prototype.offsets = null; b.prototype.useDistinct = false; b.prototype.distinct = function() { this.useDistinct = true; return this }; b.prototype.from = function(a, b) { if (b == null) b = null; a = sanitizeTable(a); if (b) b = sanitizeAlias(b); this.froms.push({ name: a, alias: b }); return this }; b.prototype.field = function(a, b) { if (b == null) b = null; a = sanitizeField(a); if (b) b = sanitizeAlias(b); this.fields.push({ field: a, alias: b }); return this }; b.prototype.join = function(a, b, c) { if (b == null) b = null; if (c == null) c = null; return this._join("INNER", a, b, c) }; b.prototype.left_join = function(a, b, c) { if (b == null) b = null; if (c == null) c = null; return this._join("LEFT", a, b, c) }; b.prototype.right_join = function(a, b, c) { if (b == null) b = null; if (c == null) c = null; return this._join("RIGHT", a, b, c) }; b.prototype.outer_join = function(a, b, c) { if (b == null) b = null; if (c == null) c = null; return this._join("OUTER", a, b, c) }; b.prototype.group = function(a) { a = sanitizeField(a); this.groups.push(a); return this }; b.prototype.offset = function(a) { a = sanitizeLimitOffset(a); this.offsets = a; return this }; b.prototype.toString = function() { var a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u; if (0 >= this.froms.length) throw new Error("from() needs to be called"); g = "SELECT "; if (this.useDistinct) g += "DISTINCT "; c = ""; r = this.fields; for (j = 0, n = r.length; j < n; j++) { b = r[j]; if ("" !== c) c += ", "; c += b.field; if (b.alias) c += ' AS "' + b.alias + '"' } g += "" === c ? "*" : c; i = ""; s = this.froms; for (k = 0, o = s.length; k < o; k++) { h = s[k]; if ("" !== i) i += ", "; i += h.name; if (h.alias) i += " `" + h.alias + "`" } g += " FROM " + i; f = ""; t = this.joins; for (l = 0, p = t.length; l < p; l++) { e = t[l]; f += " " + e.type + " JOIN " + e.table; if (e.alias) f += " `" + e.alias + "`"; if (e.condition) f += " ON (" + e.condition + ")" } g += f; g += this.whereString(); if (0 < this.groups.length) { d = ""; u = this.groups; for (m = 0, q = u.length; m < q; m++) { a = u[m]; if ("" !== d) d += ", "; d += a } g += " GROUP BY " + d } g += this.orderString(); g += this.limitString(); if (this.offsets) g += " OFFSET " + this.offsets; return g }; return b } (WhereOrderLimit); Update = function(a) { function b(a) { this.toString = __bind(this.toString, this); this.set = __bind(this.set, this); this.table = __bind(this.table, this); b.__super__.constructor.apply(this, arguments); this.tables = []; this.fields = {}; this.options = _extend({}, DefaultUpdateBuilderOptions, a) } __extends(b, a); b.prototype.tables = null; b.prototype.fields = null; b.prototype.options = null; b.prototype.table = function(a, b) { if (b == null) b = null; a = sanitizeTable(a); if (b) b = sanitizeAlias(b); this.tables.push({ name: a, alias: b }); return this }; b.prototype.set = function(a, b) { a = sanitizeField(a); b = sanitizeValue(b); this.fields[a] = b; return this }; b.prototype.toString = function() { var a, b, c, d, e, f, g, h, i, j, k; if (0 >= this.tables.length) throw new Error("table() needs to be called"); b = function() { var b, c; b = this.fields; c = []; for (a in b) { if (!__hasProp.call(b, a)) continue; c.push(a) } return c } .call(this); if (0 >= b.length) throw new Error("set() needs to be called"); d = "UPDATE "; f = ""; k = this.tables; for (g = 0, i = k.length; g < i; g++) { e = k[g]; if ("" !== f) f += ", "; f += e.name; if (e.alias) f += " AS `" + e.alias + "`" } d += f; c = ""; for (h = 0, j = b.length; h < j; h++) { a = b[h]; if ("" !== c) c += ", "; c += "" + a + " = " + formatValue(this.fields[a], this.options) } d += " SET " + c; d += this.whereString(); d += this.orderString(); d += this.limitString(); return d }; return b } (WhereOrderLimit); Delete = function(a) { function b() { this.toString = __bind(this.toString, this); this.from = __bind(this.from, this); b.__super__.constructor.apply(this, arguments) } __extends(b, a); b.prototype.table = null; b.prototype.from = function(a) { a = sanitizeTable(a); this.table = a; return this }; b.prototype.toString = function() { var a; if (!this.table) throw new Error("from() needs to be called"); a = "DELETE FROM " + this.table; a += this.whereString(); a += this.orderString(); a += this.limitString(); return a }; return b } (WhereOrderLimit); Insert = function() { function a(a) { this.toString = __bind(this.toString, this); this.set = __bind(this.set, this); this.into = __bind(this.into, this); this.fields = {}; this.options = _extend({}, DefaultInsertBuilderOptions, a) } a.prototype.table = null; a.prototype.fields = null; a.prototype.options = null; a.prototype.into = function(a) { a = sanitizeTable(a); this.table = a; return this }; a.prototype.set = function(a, b) { a = sanitizeField(a); b = sanitizeValue(b); this.fields[a] = b; return this }; a.prototype.toString = function() { var a, b, c, d, e, f, g; if (!this.table) throw new Error("into() needs to be called"); b = function() { var a, b; a = this.fields; b = []; for (d in a) { if (!__hasProp.call(a, d)) continue; b.push(d) } return b } .call(this); if (0 >= b.length) throw new Error("set() needs to be called"); c = ""; e = ""; for (f = 0, g = b.length; f < g; f++) { a = b[f]; if ("" !== c) c += ", "; c += a; if ("" !== e) e += ", "; e += formatValue(this.fields[a], this.options) } return "INSERT INTO " + this.table + " (" + c + ") VALUES (" + e + ")" }; return a } (); _export = { expr: function() { return new Expression }, select: function() { return new Select }, update: function(a) { return new Update(a) }, insert: function(a) { return new Insert(a) }, "delete": function() { return new Delete } }; sqlBuilder = _export;
    }).call(this);

    // Javascript Linq Implemenation
    Enumerable = function() { var m = "Single:sequence contains more than one element.", e = true, b = null, a = false, c = function(a) { this.GetEnumerator = a }; c.Choice = function() { var a = arguments[0] instanceof Array ? arguments[0] : arguments; return new c(function() { return new f(g.Blank, function() { return this.Yield(a[Math.floor(Math.random() * a.length)]) }, g.Blank) }) }; c.Cycle = function() { var a = arguments[0] instanceof Array ? arguments[0] : arguments; return new c(function() { var b = 0; return new f(g.Blank, function() { if (b >= a.length) b = 0; return this.Yield(a[b++]) }, g.Blank) }) }; c.Empty = function() { return new c(function() { return new f(g.Blank, function() { return a }, g.Blank) }) }; c.From = function(j) { if (j == b) return c.Empty(); if (j instanceof c) return j; if (typeof j == i.Number || typeof j == i.Boolean) return c.Repeat(j, 1); if (typeof j == i.String) return new c(function() { var b = 0; return new f(g.Blank, function() { return b < j.length ? this.Yield(j.charAt(b++)) : a }, g.Blank) }); if (typeof j != i.Function) { if (typeof j.length == i.Number) return new h(j); if (!(j instanceof Object) && d.IsIEnumerable(j)) return new c(function() { var c = e, b; return new f(function() { b = new Enumerator(j) }, function() { if (c) c = a; else b.moveNext(); return b.atEnd() ? a : this.Yield(b.item()) }, g.Blank) }) } return new c(function() { var b = [], c = 0; return new f(function() { for (var a in j) !(j[a] instanceof Function) && b.push({ Key: a, Value: j[a] }) }, function() { return c < b.length ? this.Yield(b[c++]) : a }, g.Blank) }) }, c.Return = function(a) { return c.Repeat(a, 1) }; c.Matches = function(h, e, d) { if (d == b) d = ""; if (e instanceof RegExp) { d += e.ignoreCase ? "i" : ""; d += e.multiline ? "m" : ""; e = e.source } if (d.indexOf("g") === -1) d += "g"; return new c(function() { var b; return new f(function() { b = new RegExp(e, d) }, function() { var c = b.exec(h); return c ? this.Yield(c) : a }, g.Blank) }) }; c.Range = function(e, d, a) { if (a == b) a = 1; return c.ToInfinity(e, a).Take(d) }; c.RangeDown = function(e, d, a) { if (a == b) a = 1; return c.ToNegativeInfinity(e, a).Take(d) }; c.RangeTo = function(d, e, a) { if (a == b) a = 1; return d < e ? c.ToInfinity(d, a).TakeWhile(function(a) { return a <= e }) : c.ToNegativeInfinity(d, a).TakeWhile(function(a) { return a >= e }) }; c.Repeat = function(d, a) { return a != b ? c.Repeat(d).Take(a) : new c(function() { return new f(g.Blank, function() { return this.Yield(d) }, g.Blank) }) }; c.RepeatWithFinalize = function(a, e) { a = d.CreateLambda(a); e = d.CreateLambda(e); return new c(function() { var c; return new f(function() { c = a() }, function() { return this.Yield(c) }, function() { if (c != b) { e(c); c = b } }) }) }; c.Generate = function(a, e) { if (e != b) return c.Generate(a).Take(e); a = d.CreateLambda(a); return new c(function() { return new f(g.Blank, function() { return this.Yield(a()) }, g.Blank) }) }; c.ToInfinity = function(d, a) { if (d == b) d = 0; if (a == b) a = 1; return new c(function() { var b; return new f(function() { b = d - a }, function() { return this.Yield(b += a) }, g.Blank) }) }; c.ToNegativeInfinity = function(d, a) { if (d == b) d = 0; if (a == b) a = 1; return new c(function() { var b; return new f(function() { b = d + a }, function() { return this.Yield(b -= a) }, g.Blank) }) }; c.Unfold = function(h, b) { b = d.CreateLambda(b); return new c(function() { var d = e, c; return new f(g.Blank, function() { if (d) { d = a; c = h; return this.Yield(c) } c = b(c); return this.Yield(c) }, g.Blank) }) }; c.prototype = { CascadeBreadthFirst: function(g, b) { var h = this; g = d.CreateLambda(g); b = d.CreateLambda(b); return new c(function() { var i, k = 0, j = []; return new f(function() { i = h.GetEnumerator() }, function() { while (e) { if (i.MoveNext()) { j.push(i.Current()); return this.Yield(b(i.Current(), k)) } var f = c.From(j).SelectMany(function(a) { return g(a) }); if (!f.Any()) return a; else { k++; j = []; d.Dispose(i); i = f.GetEnumerator() } } }, function() { d.Dispose(i) }) }) }, CascadeDepthFirst: function(g, b) { var h = this; g = d.CreateLambda(g); b = d.CreateLambda(b); return new c(function() { var j = [], i; return new f(function() { i = h.GetEnumerator() }, function() { while (e) { if (i.MoveNext()) { var f = b(i.Current(), j.length); j.push(i); i = c.From(g(i.Current())).GetEnumerator(); return this.Yield(f) } if (j.length <= 0) return a; d.Dispose(i); i = j.pop() } }, function() { try { d.Dispose(i) } finally { c.From(j).ForEach(function(a) { a.Dispose() }) } }) }) }, Flatten: function() { var h = this; return new c(function() { var j, i = b; return new f(function() { j = h.GetEnumerator() }, function() { while (e) { if (i != b) if (i.MoveNext()) return this.Yield(i.Current()); else i = b; if (j.MoveNext()) if (j.Current() instanceof Array) { d.Dispose(i); i = c.From(j.Current()).SelectMany(g.Identity).Flatten().GetEnumerator(); continue } else return this.Yield(j.Current()); return a } }, function() { try { d.Dispose(j) } finally { d.Dispose(i) } }) }) }, Pairwise: function(b) { var e = this; b = d.CreateLambda(b); return new c(function() { var c; return new f(function() { c = e.GetEnumerator(); c.MoveNext() }, function() { var d = c.Current(); return c.MoveNext() ? this.Yield(b(d, c.Current())) : a }, function() { d.Dispose(c) }) }) }, Scan: function(i, g, j) { if (j != b) return this.Scan(i, g).Select(j); var h; if (g == b) { g = d.CreateLambda(i); h = a } else { g = d.CreateLambda(g); h = e } var k = this; return new c(function() { var b, c, j = e; return new f(function() { b = k.GetEnumerator() }, function() { if (j) { j = a; if (!h) { if (b.MoveNext()) return this.Yield(c = b.Current()) } else return this.Yield(c = i) } return b.MoveNext() ? this.Yield(c = g(c, b.Current())) : a }, function() { d.Dispose(b) }) }) }, Select: function(b) { var e = this; b = d.CreateLambda(b); return new c(function() { var c, g = 0; return new f(function() { c = e.GetEnumerator() }, function() { return c.MoveNext() ? this.Yield(b(c.Current(), g++)) : a }, function() { d.Dispose(c) }) }) }, SelectMany: function(g, e) { var h = this; g = d.CreateLambda(g); if (e == b) e = function(b, a) { return a }; e = d.CreateLambda(e); return new c(function() { var j, i = undefined, k = 0; return new f(function() { j = h.GetEnumerator() }, function() { if (i === undefined) if (!j.MoveNext()) return a; do { if (i == b) { var f = g(j.Current(), k++); i = c.From(f).GetEnumerator() } if (i.MoveNext()) return this.Yield(e(j.Current(), i.Current())); d.Dispose(i); i = b } while (j.MoveNext()); return a }, function() { try { d.Dispose(j) } finally { d.Dispose(i) } }) }) }, Where: function(b) { b = d.CreateLambda(b); var e = this; return new c(function() { var c, g = 0; return new f(function() { c = e.GetEnumerator() }, function() { while (c.MoveNext()) if (b(c.Current(), g++)) return this.Yield(c.Current()); return a }, function() { d.Dispose(c) }) }) }, OfType: function(c) { var a; switch (c) { case Number: a = i.Number; break; case String: a = i.String; break; case Boolean: a = i.Boolean; break; case Function: a = i.Function; break; default: a = b } return a === b ? this.Where(function(a) { return a instanceof c }) : this.Where(function(b) { return typeof b === a }) }, Zip: function(e, b) { b = d.CreateLambda(b); var g = this; return new c(function() { var i, h, j = 0; return new f(function() { i = g.GetEnumerator(); h = c.From(e).GetEnumerator() }, function() { return i.MoveNext() && h.MoveNext() ? this.Yield(b(i.Current(), h.Current(), j++)) : a }, function() { try { d.Dispose(i) } finally { d.Dispose(h) } }) }) }, Join: function(m, i, h, k, j) { i = d.CreateLambda(i); h = d.CreateLambda(h); k = d.CreateLambda(k); j = d.CreateLambda(j); var l = this; return new c(function() { var n, q, o = b, p = 0; return new f(function() { n = l.GetEnumerator(); q = c.From(m).ToLookup(h, g.Identity, j) }, function() { while (e) { if (o != b) { var c = o[p++]; if (c !== undefined) return this.Yield(k(n.Current(), c)); c = b; p = 0 } if (n.MoveNext()) { var d = i(n.Current()); o = q.Get(d).ToArray() } else return a } }, function() { d.Dispose(n) }) }) }, GroupJoin: function(l, h, e, j, i) { h = d.CreateLambda(h); e = d.CreateLambda(e); j = d.CreateLambda(j); i = d.CreateLambda(i); var k = this; return new c(function() { var m = k.GetEnumerator(), n = b; return new f(function() { m = k.GetEnumerator(); n = c.From(l).ToLookup(e, g.Identity, i) }, function() { if (m.MoveNext()) { var b = n.Get(h(m.Current())); return this.Yield(j(m.Current(), b)) } return a }, function() { d.Dispose(m) }) }) }, All: function(b) { b = d.CreateLambda(b); var c = e; this.ForEach(function(d) { if (!b(d)) { c = a; return a } }); return c }, Any: function(c) { c = d.CreateLambda(c); var b = this.GetEnumerator(); try { if (arguments.length == 0) return b.MoveNext(); while (b.MoveNext()) if (c(b.Current())) return e; return a } finally { d.Dispose(b) } }, Concat: function(e) { var g = this; return new c(function() { var i, h; return new f(function() { i = g.GetEnumerator() }, function() { if (h == b) { if (i.MoveNext()) return this.Yield(i.Current()); h = c.From(e).GetEnumerator() } return h.MoveNext() ? this.Yield(h.Current()) : a }, function() { try { d.Dispose(i) } finally { d.Dispose(h) } }) }) }, Insert: function(h, b) { var g = this; return new c(function() { var j, i, l = 0, k = a; return new f(function() { j = g.GetEnumerator(); i = c.From(b).GetEnumerator() }, function() { if (l == h && i.MoveNext()) { k = e; return this.Yield(i.Current()) } if (j.MoveNext()) { l++; return this.Yield(j.Current()) } return !k && i.MoveNext() ? this.Yield(i.Current()) : a }, function() { try { d.Dispose(j) } finally { d.Dispose(i) } }) }) }, Alternate: function(a) { a = c.Return(a); return this.SelectMany(function(b) { return c.Return(b).Concat(a) }).TakeExceptLast() }, Contains: function(f, b) { b = d.CreateLambda(b); var c = this.GetEnumerator(); try { while (c.MoveNext()) if (b(c.Current()) === f) return e; return a } finally { d.Dispose(c) } }, DefaultIfEmpty: function(b) { var g = this; return new c(function() { var c, h = e; return new f(function() { c = g.GetEnumerator() }, function() { if (c.MoveNext()) { h = a; return this.Yield(c.Current()) } else if (h) { h = a; return this.Yield(b) } return a }, function() { d.Dispose(c) }) }) }, Distinct: function(a) { return this.Except(c.Empty(), a) }, Except: function(e, b) { b = d.CreateLambda(b); var g = this; return new c(function() { var h, i; return new f(function() { h = g.GetEnumerator(); i = new n(b); c.From(e).ForEach(function(a) { i.Add(a) }) }, function() { while (h.MoveNext()) { var b = h.Current(); if (!i.Contains(b)) { i.Add(b); return this.Yield(b) } } return a }, function() { d.Dispose(h) }) }) }, Intersect: function(e, b) { b = d.CreateLambda(b); var g = this; return new c(function() { var h, i, j; return new f(function() { h = g.GetEnumerator(); i = new n(b); c.From(e).ForEach(function(a) { i.Add(a) }); j = new n(b) }, function() { while (h.MoveNext()) { var b = h.Current(); if (!j.Contains(b) && i.Contains(b)) { j.Add(b); return this.Yield(b) } } return a }, function() { d.Dispose(h) }) }) }, SequenceEqual: function(h, f) { f = d.CreateLambda(f); var g = this.GetEnumerator(); try { var b = c.From(h).GetEnumerator(); try { while (g.MoveNext()) if (!b.MoveNext() || f(g.Current()) !== f(b.Current())) return a; return b.MoveNext() ? a : e } finally { d.Dispose(b) } } finally { d.Dispose(g) } }, Union: function(e, b) { b = d.CreateLambda(b); var g = this; return new c(function() { var j, h, i; return new f(function() { j = g.GetEnumerator(); i = new n(b) }, function() { var b; if (h === undefined) { while (j.MoveNext()) { b = j.Current(); if (!i.Contains(b)) { i.Add(b); return this.Yield(b) } } h = c.From(e).GetEnumerator() } while (h.MoveNext()) { b = h.Current(); if (!i.Contains(b)) { i.Add(b); return this.Yield(b) } } return a }, function() { try { d.Dispose(j) } finally { d.Dispose(h) } }) }) }, OrderBy: function(b) { return new j(this, b, a) }, OrderByDescending: function(a) { return new j(this, a, e) }, Reverse: function() { var b = this; return new c(function() { var c, d; return new f(function() { c = b.ToArray(); d = c.length }, function() { return d > 0 ? this.Yield(c[--d]) : a }, g.Blank) }) }, Shuffle: function() { var b = this; return new c(function() { var c; return new f(function() { c = b.ToArray() }, function() { if (c.length > 0) { var b = Math.floor(Math.random() * c.length); return this.Yield(c.splice(b, 1)[0]) } return a }, g.Blank) }) }, GroupBy: function(i, h, e, g) { var j = this; i = d.CreateLambda(i); h = d.CreateLambda(h); if (e != b) e = d.CreateLambda(e); g = d.CreateLambda(g); return new c(function() { var c; return new f(function() { c = j.ToLookup(i, h, g).ToEnumerable().GetEnumerator() }, function() { while (c.MoveNext()) return e == b ? this.Yield(c.Current()) : this.Yield(e(c.Current().Key(), c.Current())); return a }, function() { d.Dispose(c) }) }) }, PartitionBy: function(j, i, g, h) { var l = this; j = d.CreateLambda(j); i = d.CreateLambda(i); h = d.CreateLambda(h); var k; if (g == b) { k = a; g = function(b, a) { return new o(b, a) } } else { k = e; g = d.CreateLambda(g) } return new c(function() { var b, n, o, m = []; return new f(function() { b = l.GetEnumerator(); if (b.MoveNext()) { n = j(b.Current()); o = h(n); m.push(i(b.Current())) } }, function() { var d; while ((d = b.MoveNext()) == e) if (o === h(j(b.Current()))) m.push(i(b.Current())); else break; if (m.length > 0) { var f = k ? g(n, c.From(m)) : g(n, m); if (d) { n = j(b.Current()); o = h(n); m = [i(b.Current())] } else m = []; return this.Yield(f) } return a }, function() { d.Dispose(b) }) }) }, BufferWithCount: function(e) { var b = this; return new c(function() { var c; return new f(function() { c = b.GetEnumerator() }, function() { var b = [], d = 0; while (c.MoveNext()) { b.push(c.Current()); if (++d >= e) return this.Yield(b) } return b.length > 0 ? this.Yield(b) : a }, function() { d.Dispose(c) }) }) }, Aggregate: function(c, b, a) { return this.Scan(c, b, a).Last() }, Average: function(a) { a = d.CreateLambda(a); var c = 0, b = 0; this.ForEach(function(d) { c += a(d); ++b }); return c / b }, Count: function(a) { a = a == b ? g.True : d.CreateLambda(a); var c = 0; this.ForEach(function(d, b) { if (a(d, b)) ++c }); return c }, Max: function(a) { if (a == b) a = g.Identity; return this.Select(a).Aggregate(function(a, b) { return a > b ? a : b }) }, Min: function(a) { if (a == b) a = g.Identity; return this.Select(a).Aggregate(function(a, b) { return a < b ? a : b }) }, MaxBy: function(a) { a = d.CreateLambda(a); return this.Aggregate(function(b, c) { return a(b) > a(c) ? b : c }) }, MinBy: function(a) { a = d.CreateLambda(a); return this.Aggregate(function(b, c) { return a(b) < a(c) ? b : c }) }, Sum: function(a) { if (a == b) a = g.Identity; return this.Select(a).Aggregate(0, function(a, b) { return a + b }) }, ElementAt: function(d) { var c, b = a; this.ForEach(function(g, f) { if (f == d) { c = g; b = e; return a } }); if (!b) throw new Error("index is less than 0 or greater than or equal to the number of elements in source."); return c }, ElementAtOrDefault: function(f, d) { var c, b = a; this.ForEach(function(g, d) { if (d == f) { c = g; b = e; return a } }); return !b ? d : c }, First: function(c) { if (c != b) return this.Where(c).First(); var f, d = a; this.ForEach(function(b) { f = b; d = e; return a }); if (!d) throw new Error("First:No element satisfies the condition."); return f }, FirstOrDefault: function(c, d) { if (d != b) return this.Where(d).FirstOrDefault(c); var g, f = a; this.ForEach(function(b) { g = b; f = e; return a }); return !f ? c : g }, Last: function(c) { if (c != b) return this.Where(c).Last(); var f, d = a; this.ForEach(function(a) { d = e; f = a }); if (!d) throw new Error("Last:No element satisfies the condition."); return f }, LastOrDefault: function(c, d) { if (d != b) return this.Where(d).LastOrDefault(c); var g, f = a; this.ForEach(function(a) { f = e; g = a }); return !f ? c : g }, Single: function(d) { if (d != b) return this.Where(d).Single(); var f, c = a; this.ForEach(function(a) { if (!c) { c = e; f = a } else throw new Error(m); }); if (!c) throw new Error("Single:No element satisfies the condition."); return f }, SingleOrDefault: function(d, f) { if (f != b) return this.Where(f).SingleOrDefault(d); var g, c = a; this.ForEach(function(a) { if (!c) { c = e; g = a } else throw new Error(m); }); return !c ? d : g }, Skip: function(e) { var b = this; return new c(function() { var c, g = 0; return new f(function() { c = b.GetEnumerator(); while (g++ < e && c.MoveNext()); }, function() { return c.MoveNext() ? this.Yield(c.Current()) : a }, function() { d.Dispose(c) }) }) }, SkipWhile: function(b) { b = d.CreateLambda(b); var g = this; return new c(function() { var c, i = 0, h = a; return new f(function() { c = g.GetEnumerator() }, function() { while (!h) if (c.MoveNext()) { if (!b(c.Current(), i++)) { h = e; return this.Yield(c.Current()) } continue } else return a; return c.MoveNext() ? this.Yield(c.Current()) : a }, function() { d.Dispose(c) }) }) }, Take: function(e) { var b = this; return new c(function() { var c, g = 0; return new f(function() { c = b.GetEnumerator() }, function() { return g++ < e && c.MoveNext() ? this.Yield(c.Current()) : a }, function() { d.Dispose(c) }) }) }, TakeWhile: function(b) { b = d.CreateLambda(b); var e = this; return new c(function() { var c, g = 0; return new f(function() { c = e.GetEnumerator() }, function() { return c.MoveNext() && b(c.Current(), g++) ? this.Yield(c.Current()) : a }, function() { d.Dispose(c) }) }) }, TakeExceptLast: function(e) { if (e == b) e = 1; var g = this; return new c(function() { if (e <= 0) return g.GetEnumerator(); var b, c = []; return new f(function() { b = g.GetEnumerator() }, function() { while (b.MoveNext()) { if (c.length == e) { c.push(b.Current()); return this.Yield(c.shift()) } c.push(b.Current()) } return a }, function() { d.Dispose(b) }) }) }, TakeFromLast: function(e) { if (e <= 0 || e == b) return c.Empty(); var g = this; return new c(function() { var j, h, i = []; return new f(function() { j = g.GetEnumerator() }, function() { while (j.MoveNext()) { i.length == e && i.shift(); i.push(j.Current()) } if (h == b) h = c.From(i).GetEnumerator(); return h.MoveNext() ? this.Yield(h.Current()) : a }, function() { d.Dispose(h) }) }) }, IndexOf: function(c) { var a = b; this.ForEach(function(d, b) { if (d === c) { a = b; return e } }); return a !== b ? a : -1 }, LastIndexOf: function(b) { var a = -1; this.ForEach(function(d, c) { if (d === b) a = c }); return a }, ToArray: function() { var a = []; this.ForEach(function(b) { a.push(b) }); return a }, ToLookup: function(c, b, a) { c = d.CreateLambda(c); b = d.CreateLambda(b); a = d.CreateLambda(a); var e = new n(a); this.ForEach(function(g) { var f = c(g), a = b(g), d = e.Get(f); if (d !== undefined) d.push(a); else e.Add(f, [a]) }); return new q(e) }, ToObject: function(b, a) { b = d.CreateLambda(b); a = d.CreateLambda(a); var c = {}; this.ForEach(function(d) { c[b(d)] = a(d) }); return c }, ToDictionary: function(c, b, a) { c = d.CreateLambda(c); b = d.CreateLambda(b); a = d.CreateLambda(a); var e = new n(a); this.ForEach(function(a) { e.Add(c(a), b(a)) }); return e }, ToJSON: function(a, b) { return JSON.stringify(this.ToArray(), a, b) }, ToString: function(a, c) { if (a == b) a = ""; if (c == b) c = g.Identity; return this.Select(c).ToArray().join(a) }, Do: function(b) { var e = this; b = d.CreateLambda(b); return new c(function() { var c, g = 0; return new f(function() { c = e.GetEnumerator() }, function() { if (c.MoveNext()) { b(c.Current(), g++); return this.Yield(c.Current()) } return a }, function() { d.Dispose(c) }) }) }, ForEach: function(c) { c = d.CreateLambda(c); var e = 0, b = this.GetEnumerator(); try { while (b.MoveNext()) if (c(b.Current(), e++) === a) break } finally { d.Dispose(b) } }, Write: function(c, f) { if (c == b) c = ""; f = d.CreateLambda(f); var g = e; this.ForEach(function(b) { if (g) g = a; else document.write(c); document.write(f(b)) }) }, WriteLine: function(a) { a = d.CreateLambda(a); this.ForEach(function(b) { document.write(a(b)); document.write("<br />") }) }, Force: function() { var a = this.GetEnumerator(); try { while (a.MoveNext()); } finally { d.Dispose(a) } }, Let: function(b) { b = d.CreateLambda(b); var e = this; return new c(function() { var g; return new f(function() { g = c.From(b(e)).GetEnumerator() }, function() { return g.MoveNext() ? this.Yield(g.Current()) : a }, function() { d.Dispose(g) }) }) }, Share: function() { var e = this, d; return new c(function() { return new f(function() { if (d == b) d = e.GetEnumerator() }, function() { return d.MoveNext() ? this.Yield(d.Current()) : a }, g.Blank) }) }, MemoizeAll: function() { var h = this, e, d; return new c(function() { var c = -1; return new f(function() { if (d == b) { d = h.GetEnumerator(); e = [] } }, function() { c++; return e.length <= c ? d.MoveNext() ? this.Yield(e[c] = d.Current()) : a : this.Yield(e[c]) }, g.Blank) }) }, Catch: function(b) { b = d.CreateLambda(b); var e = this; return new c(function() { var c; return new f(function() { c = e.GetEnumerator() }, function() { try { return c.MoveNext() ? this.Yield(c.Current()) : a } catch (d) { b(d); return a } }, function() { d.Dispose(c) }) }) }, Finally: function(b) { b = d.CreateLambda(b); var e = this; return new c(function() { var c; return new f(function() { c = e.GetEnumerator() }, function() { return c.MoveNext() ? this.Yield(c.Current()) : a }, function() { try { d.Dispose(c) } finally { b() } }) }) }, Trace: function(c, a) { if (c == b) c = "Trace"; a = d.CreateLambda(a); return this.Do(function(b) { console.log(c, ":", a(b)) }) } }; var g = { Identity: function(a) { return a }, True: function() { return e }, Blank: function() { } }, i = { Boolean: typeof e, Number: typeof 0, String: typeof "", Object: typeof {}, Undefined: typeof undefined, Function: typeof function() { } }, d = { CreateLambda: function(a) { if (a == b) return g.Identity; if (typeof a == i.String) if (a == "") return g.Identity; else if (a.indexOf("=>") == -1) return new Function("$,$$,$$$,$$$$", "return " + a); else { var c = a.match(/^[(\s]*([^()]*?)[)\s]*=>(.*)/); return new Function(c[1], "return " + c[2]) } return a }, IsIEnumerable: function(b) { if (typeof Enumerator != i.Undefined) try { new Enumerator(b); return e } catch (c) { } return a }, Compare: function(a, b) { return a === b ? 0 : a > b ? 1 : -1 }, Dispose: function(a) { a != b && a.Dispose() } }, k = { Before: 0, Running: 1, After: 2 }, f = function(d, f, g) { var c = new p, b = k.Before; this.Current = c.Current; this.MoveNext = function() { try { switch (b) { case k.Before: b = k.Running; d(); case k.Running: if (f.apply(c)) return e; else { this.Dispose(); return a } case k.After: return a } } catch (g) { this.Dispose(); throw g; } }; this.Dispose = function() { if (b != k.Running) return; try { g() } finally { b = k.After } } }, p = function() { var a = b; this.Current = function() { return a }; this.Yield = function(b) { a = b; return e } }, j = function(f, b, c, e) { var a = this; a.source = f; a.keySelector = d.CreateLambda(b); a.descending = c; a.parent = e }; j.prototype = new c; j.prototype.CreateOrderedEnumerable = function(a, b) { return new j(this.source, a, b, this) }; j.prototype.ThenBy = function(b) { return this.CreateOrderedEnumerable(b, a) }; j.prototype.ThenByDescending = function(a) { return this.CreateOrderedEnumerable(a, e) }; j.prototype.GetEnumerator = function() { var h = this, d, c, e = 0; return new f(function() { d = []; c = []; h.source.ForEach(function(b, a) { d.push(b); c.push(a) }); var a = l.Create(h, b); a.GenerateKeys(d); c.sort(function(b, c) { return a.Compare(b, c) }) }, function() { return e < c.length ? this.Yield(d[c[e++]]) : a }, g.Blank) }; var l = function(c, d, e) { var a = this; a.keySelector = c; a.descending = d; a.child = e; a.keys = b }; l.Create = function(a, d) { var c = new l(a.keySelector, a.descending, d); return a.parent != b ? l.Create(a.parent, c) : c }; l.prototype.GenerateKeys = function(d) { var a = this; for (var f = d.length, g = a.keySelector, e = new Array(f), c = 0; c < f; c++) e[c] = g(d[c]); a.keys = e; a.child != b && a.child.GenerateKeys(d) }; l.prototype.Compare = function(e, f) { var a = this, c = d.Compare(a.keys[e], a.keys[f]); if (c == 0) { if (a.child != b) return a.child.Compare(e, f); c = d.Compare(e, f) } return a.descending ? -c : c }; var h = function(a) { this.source = a }; h.prototype = new c; h.prototype.Any = function(a) { return a == b ? this.source.length > 0 : c.prototype.Any.apply(this, arguments) }; h.prototype.Count = function(a) { return a == b ? this.source.length : c.prototype.Count.apply(this, arguments) }; h.prototype.ElementAt = function(a) { return 0 <= a && a < this.source.length ? this.source[a] : c.prototype.ElementAt.apply(this, arguments) }; h.prototype.ElementAtOrDefault = function(a, b) { return 0 <= a && a < this.source.length ? this.source[a] : b }; h.prototype.First = function(a) { return a == b && this.source.length > 0 ? this.source[0] : c.prototype.First.apply(this, arguments) }; h.prototype.FirstOrDefault = function(a, d) { return d != b ? c.prototype.FirstOrDefault.apply(this, arguments) : this.source.length > 0 ? this.source[0] : a }; h.prototype.Last = function(d) { var a = this; return d == b && a.source.length > 0 ? a.source[a.source.length - 1] : c.prototype.Last.apply(a, arguments) }; h.prototype.LastOrDefault = function(d, e) { var a = this; return e != b ? c.prototype.LastOrDefault.apply(a, arguments) : a.source.length > 0 ? a.source[a.source.length - 1] : d }; h.prototype.Skip = function(d) { var b = this.source; return new c(function() { var c; return new f(function() { c = d < 0 ? 0 : d }, function() { return c < b.length ? this.Yield(b[c++]) : a }, g.Blank) }) }; h.prototype.TakeExceptLast = function(a) { if (a == b) a = 1; return this.Take(this.source.length - a) }; h.prototype.TakeFromLast = function(a) { return this.Skip(this.source.length - a) }; h.prototype.Reverse = function() { var b = this.source; return new c(function() { var c; return new f(function() { c = b.length }, function() { return c > 0 ? this.Yield(b[--c]) : a }, g.Blank) }) }; h.prototype.SequenceEqual = function(d, e) { return (d instanceof h || d instanceof Array) && e == b && c.From(d).Count() != this.Count() ? a : c.prototype.SequenceEqual.apply(this, arguments) }; h.prototype.ToString = function(a, d) { if (d != b || !(this.source instanceof Array)) return c.prototype.ToString.apply(this, arguments); if (a == b) a = ""; return this.source.join(a) }; h.prototype.GetEnumerator = function() { var b = this.source, c = 0; return new f(g.Blank, function() { return c < b.length ? this.Yield(b[c++]) : a }, g.Blank) }; var n = function() { var h = function(a, b) { return Object.prototype.hasOwnProperty.call(a, b) }, d = function(a) { return a === b ? "null" : a === undefined ? "undefined" : typeof a.toString === i.Function ? a.toString() : Object.prototype.toString.call(a) }, l = function(d, c) { var a = this; a.Key = d; a.Value = c; a.Prev = b; a.Next = b }, j = function() { this.First = b; this.Last = b }; j.prototype = { AddLast: function(c) { var a = this; if (a.Last != b) { a.Last.Next = c; c.Prev = a.Last; a.Last = c } else a.First = a.Last = c }, Replace: function(c, a) { if (c.Prev != b) { c.Prev.Next = a; a.Prev = c.Prev } else this.First = a; if (c.Next != b) { c.Next.Prev = a; a.Next = c.Next } else this.Last = a }, Remove: function(a) { if (a.Prev != b) a.Prev.Next = a.Next; else this.First = a.Next; if (a.Next != b) a.Next.Prev = a.Prev; else this.Last = a.Prev } }; var k = function(c) { var a = this; a.count = 0; a.entryList = new j; a.buckets = {}; a.compareSelector = c == b ? g.Identity : c }; k.prototype = { Add: function(i, j) { var a = this, g = a.compareSelector(i), f = d(g), c = new l(i, j); if (h(a.buckets, f)) { for (var b = a.buckets[f], e = 0; e < b.length; e++) if (a.compareSelector(b[e].Key) === g) { a.entryList.Replace(b[e], c); b[e] = c; return } b.push(c) } else a.buckets[f] = [c]; a.count++; a.entryList.AddLast(c) }, Get: function(i) { var a = this, c = a.compareSelector(i), g = d(c); if (!h(a.buckets, g)) return undefined; for (var e = a.buckets[g], b = 0; b < e.length; b++) { var f = e[b]; if (a.compareSelector(f.Key) === c) return f.Value } return undefined }, Set: function(k, m) { var b = this, g = b.compareSelector(k), j = d(g); if (h(b.buckets, j)) for (var f = b.buckets[j], c = 0; c < f.length; c++) if (b.compareSelector(f[c].Key) === g) { var i = new l(k, m); b.entryList.Replace(f[c], i); f[c] = i; return e } return a }, Contains: function(j) { var b = this, f = b.compareSelector(j), i = d(f); if (!h(b.buckets, i)) return a; for (var g = b.buckets[i], c = 0; c < g.length; c++) if (b.compareSelector(g[c].Key) === f) return e; return a }, Clear: function() { this.count = 0; this.buckets = {}; this.entryList = new j }, Remove: function(g) { var a = this, f = a.compareSelector(g), e = d(f); if (!h(a.buckets, e)) return; for (var b = a.buckets[e], c = 0; c < b.length; c++) if (a.compareSelector(b[c].Key) === f) { a.entryList.Remove(b[c]); b.splice(c, 1); if (b.length == 0) delete a.buckets[e]; a.count--; return } }, Count: function() { return this.count }, ToEnumerable: function() { var d = this; return new c(function() { var c; return new f(function() { c = d.entryList.First }, function() { if (c != b) { var d = { Key: c.Key, Value: c.Value }; c = c.Next; return this.Yield(d) } return a }, g.Blank) }) } }; return k } (), q = function(a) { var b = this; b.Count = function() { return a.Count() }; b.Get = function(b) { return c.From(a.Get(b)) }; b.Contains = function(b) { return a.Contains(b) }; b.ToEnumerable = function() { return a.ToEnumerable().Select(function(a) { return new o(a.Key, a.Value) }) } }, o = function(b, a) { this.Key = function() { return b }; h.call(this, a) }; o.prototype = new h; return c } ()

    // Javascript fluent HTML
    // version 0.2
    var Htmls = function() { function d(a, b) { this.name = a; this.isSelfClose = b; this.attrs = {}; this.children = [] } var a = ["id", "class", "style", "title", "lang", "dir", "accesskey", "tabindex", "charset", "type", "name", "href", "hreflang", "rel", "rev", "shape", "coords", "src", "alt", "longdesc", "height", "width", "usemap", "ismap", "action", "method", "enctype", "accept", "accept-charset", "for", "value", "checked", "disabled", "readonly", "size", "maxlength", "summary", "border", "frame", "rules", "cellspacing", "cellpadding", "align", "char", "charoff", "valign", "abbr", "axis", "headers", "scope", "rowspan", "colspan"]; var b = ["html", "head", "title", "base", "meta", "link", "style", "script", "noscript", "body", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "dl", "dt", "dd", "address", "hr", "pre", "blockquote", "ins", "del", "a", "span", "bdo", "br", "em", "strong", "dfn", "code", "samp", "kbd", "var", "cite", "abbr", "acronym", "q", "sub", "sup", "tt", "i", "b", "big", "small", "object", "param", "img", "map", "area", "form", "label", "input", "select", "optgroup", "option", "textarea", "fieldset", "legend", "button", "table", "caption", "thead", "tfoot", "tbody", "colgroup", "col", "tr", "th", "td"]; var c = { base: "", meta: "", link: "", hr: "", br: "", img: "", area: "", input: "" }; for (var e = 0; e < a.length; e++) { var f = a[e]; if (f != null || f != undefined || f.length >= 1) { methodName = String(f); d.prototype[methodName] = function(a) { return function(b) { return this.Attr(a, b) } } (f) } } d.prototype.Attr = function(a, b) { if (a.constructor == Object) { for (var c in a) { if (a.hasOwnProperty(c)) { this.attrs[c] = a[c] } } } else { this.attrs[a] = b } return this }; d.prototype.$ = function() { for (var a = 0; a < arguments.length; a++) { var b = arguments[a]; if (b == undefined) continue; if (b.constructor == Array) { for (var c = 0; c < b.length; c++) { this.children.push(b[c]) } } else if (b.constructor == Object) { this.Attr(b) } else { this.children.push(b) } } return this }; d.prototype.render = function(a) { a.push("<" + this.name); for (var b in this.attrs) { if (this.attrs.hasOwnProperty(b)) { a.push(" " + b + '="' + this.attrs[b] + '"') } } if (this.isSelfClose) { a.push("/>") } else { a.push(">"); for (var c = 0; c < this.children.length; c++) { var e = this.children[c]; if (e instanceof d) { e.render(a) } else { a.push(e.toString()) } } a.push("</" + this.name + ">") } return a }; d.prototype.toString = function() { return this.render([]).join("") }; var g = {}; for (var h = 0; h < b.length; h++) { var i = b[h]; g[i] = function(a) { return function() { var b = new d(a, a in c); return b.$.apply(b, arguments) } } (i) } return g } ()

    // Add slashes to strings to ensure they are escaped
    String.prototype.addSlashes = function() {
        return this.replace(/([\\"'])/g, "\\$1").replace(/\0/g, "\\0");
    }
    // RTrim similar to SQL Server function
    String.prototype.rtrim = function() { return this.replace(/\s+$/, ''); }

    function StoredProcedureSql(name, values) {
        var output = '';
        var parameters = [];

        var notAllowedColumnNames = ['get', 'set', 'delete', 'keys'];

        var keys = values.Keys().toArray();
        var items = values.Items().toArray();

        for (var i = 0; i < keys.length; i++) {
            if ((String(keys[i]) in notAllowedColumnNames) == false && keys[i] != null) {
                var argument = '';
                if (items[i] == null || items[i] == undefined) {
                    argument = ('@' + String(keys[i]) + ' = NULL ');
                } else {
                    if (typeof (items[i]) == 'date') {
                        output += ('DECLARE @' + String(keys[i]) + 'Param DATETIME;');
                        output += ('SET @' + String(keys[i]) + 'Param = ' + ValueToSQLValue(items[i]) + ';');
                        argument = ('@' + String(keys[i]) + '= @' + String(keys[i]) + 'Param');
                    } else {
                        argument = ('@' + String(keys[i]) + ' = ' + ValueToSQLValue(items[i]));
                    }
                }
                parameters.push(argument);
            }
        }
        output += 'EXEC ' + name + ' ';
        output += parameters.join(', ') + '';
        return output + ';';

    }
    
    function DeleteSql(tableName, primaryKey, val) {
        var output = '';
        output += 'DELETE FROM ' + tableName + ' WHERE ' + primaryKey + ' = ' + val + ';' ;
        return output;
    }

    function UpdateSql(object, tableName, primaryKey) {
        var primaryKeyValue = 0;
        var output = '';
        output += 'UPDATE ' + tableName + ' SET ';
        var values = [];

        for (property in object) {
            if (primaryKey == property) {
				
                primaryKeyValue = String(ValueToSQLValue(object[property]));
            } else {
                values.push(String(property + '=' + ValueToSQLValue(object[property])));
            }
        }
        output += values.join(',') + '';
        output += ' WHERE ' + primaryKey + '=' + primaryKeyValue;
        return output + ';';
    }

    function InsertSql(object, tableName) {
        var output = '';
        output += 'INSERT INTO ' + tableName + ' (';
        var columns = [];
        for (property in object) {
            columns.push(String(property));
        }
        output += columns.join(',') + ') VALUES (';

        var values = [];
        for (property in object) {
            values.push(String(ValueToSQLValue(object[property])));
        }
        output += values.join(',') + ');';
        return output;
    }

    function ValueToSQLValue(value) {
        if (String(value) == 'null') {
            return "NULL";
        }
        switch (typeof value) {
            case 'string':
                return '\'' + String(value).rtrim().replace("'", "''") + '\'';
            case 'boolean':
                //return (value)?1:0;
                if (value == true) { return 1 } else { return 0 };
            case 'number':
                return String(value);
            case 'date':
                return ToSQLDateTime(value);
            case 'object':
                return '\'' + String(value).rtrim() + '\'';
            default:
                return '\'' + String(value).rtrim() + '\'';
        }
    }

    function ToNumberObject(a) { var b = new Object; b = parseInt(a); return b }

    function RecordSetToInt(rs) {

        var count = 0;

        var itemObject = new Object();
        itemObject = parseInt("0");

        if (rs.State != 1) {
            return parseInt(0);
        }

        while (rs.EOF != true) {

            if (Response.IsClientConnected && count <= 1) {
                if (rs.Fields(0).Value == null) {
                    itemObject = parseInt("0");
                } else {
                    itemObject = parseInt(rs.Fields(0).Value);
                }
                count = count + 1
            } // Client Connected

            rs.MoveNext();
        }

        return itemObject;
    }

    function RecordSetToBoolean(rs) {

        var count = 0;
        var itemObject = new Object();
        itemObject = false;

        while (rs.EOF != true) {

            if (Response.IsClientConnected && count <= 1) {

                if (rs.Fields(0).Value == null) {
                    itemObject = false;
                } else {
                    itemObject = new Boolean(rs.Fields(0).Value);
                }
                count = count + 1
            } // Client Connected

            rs.MoveNext();
        }

        return itemObject;
    }

    function RecordSetToString(rs) {

        var count = 0;

        var itemObject = new Object();
        itemObject = new String('');

        while (rs.EOF != true) {

            if (Response.IsClientConnected && count <= 1) {

                if (rs.Fields(0).Value == null) {
                    itemObject = new String('');
                } else {
                    itemObject = new String(rs.Fields(0).Value);
                }
                count = count + 1
            } // Client Connected

            rs.MoveNext();
        }

        return itemObject;
    }

    function RecordSetToDateTime(rs) {

        var count = 0;

        var itemObject = new Object();
        itemObject = new Date();

        while (rs.EOF != true) {

            if (Response.IsClientConnected && count <= 1) {

                if (rs.Fields(0).Value == null) {
                    itemObject = null;
                } else {
                    itemObject = new Date(rs.Fields(0).Value).getVarDate();
                }
                count = count + 1
            } // Client Connected

            rs.MoveNext();
        }

        return itemObject;
    }

    // Convert a AdoDb RecordSet to a single dynamic object (Much like ExpandoObject in C#)
    function RecordSetToSingle(rs) {

        var count = 0;
        var itemObject = new Object();

        try {

            while (rs.EOF != true) {

                if (Response.IsClientConnected && count <= 0) {

                    createObject(rs, itemObject);
                } // Client Connected

                count = count + 1;
                rs.MoveNext();
            }

            if (count == 0) {
                eval("itemObject.IsNotNull = false;");
            } else {
                eval("itemObject.IsNotNull = true;");
            }
        }
        catch (ex) {
            Response.Write(ex.message + ',' + ex.number);
        }

        return itemObject;

    }

    // Convert a AdoDb RecordSet to a dynamic array
    function RecordSetToArray(rs) {

        var returnList = new Array();

        while (rs.EOF != true) {

            if (Response.IsClientConnected) {

                var itemObject = new Object();

                createObject(rs, itemObject);

                returnList.push(itemObject);

            } // Client Connected

            rs.MoveNext();
        }
        return returnList;
    }

    // Convert a AdoDb RecordSet to a dynamic object ArrayList
    function RecordSetToList(rs) {

        var returnList = Server.CreateObject("System.Collections.ArrayList");

        while (rs.EOF != true) {

            if (Response.IsClientConnected) {

                var itemObject = new Object();

                createObject(rs, itemObject);

                returnList.Add(itemObject);

            } // Client Connected

            rs.MoveNext();
        }
        return returnList;
	}

    //Moved the code that coverted a rs into a dynamic object to a function to consolidate replicated code
    function createObject(rs, itemObject){
        for (var i = 0; i < rs.Fields.Count; i++) {
            var dataType = rs.Fields(i).Type;
            //Response.Write(dataType + " at:" + rs.Fields(i).Name + "<br>");
            //Response.Write(typeof(rs.Fields(i).Value)==null);
            if (rs.Fields(i).Value == null) {
                eval("itemObject." + rs.Fields(i).Name + " = null;");
            } else {
                switch (dataType) {
                    case 205: // Binary
                        eval("itemObject." + rs.Fields(i).Name + " = null;");
                        break;
                    case 11: // Bool
                        eval("itemObject." + rs.Fields(i).Name + " = new Boolean(" + rs.Fields(i).Value + ");");
                        break;
                    case 3: case 5:case 6: case 14: // Int/BigInt/Numeric
                        eval("itemObject." + rs.Fields(i).Name + " = " + rs.Fields(i).Value + ";");
                        break;
                    case 16: case 2: case 20 : case 131: // Int/BigInt/Numeric
                        eval("itemObject." + rs.Fields(i).Name + " = " + rs.Fields(i).Value + ";");
                        break;
                    case 135: case 7: // DateTime
                        eval("itemObject." + rs.Fields(i).Name + " = new Date('" + rs.Fields(i).Value + "').getVarDate();");
                        break;
                    default:
                        try {
                            var value = (rs.Fields(i).Value).replace(/([\\"'])/g, "\\$1").replace(/\0/g, "\\0");
                            eval("itemObject." + rs.Fields(i).Name + " = (new String('" + value + "')).rtrim();");
                        }
                        catch (e) {
                            // Some line breaks create unterminated string constants
                            try {
                                /*
                                    Used "value" from the first try block instead of rs.Fields(i).Value so we keep the
                                    already escaped \ " ' and null strings as we escape \r and \n
                                */
                                 eval('itemObject.' + rs.Fields(i).Name + ' = \"' + value.replace(/[\r\n]+/gm,"\\ ") + '\";');
                            } catch(e) {
                                //Response.Write(e.message);
                                value = null;
                                eval("itemObject." + rs.Fields(i).Name + " = null;");
                            }
                        }
                }
            }
        }
        return itemObject;
    }
</script>
