<h2>Clapper : Asp Classic Sql Mapper</h2>
<p>Copyright 2012 Jeremy Child |<a href="mailto:jeremychild@gmail.com">Support</a></p>
<p>Released under <a href="http://en.wikipedia.org/wiki/MIT_License">MIT Licence</a></p>
<h3>Connecting</h3>
<p>You can connect with any provider that is supported with ADODB.Connection. You can check providers using ODBC in Windows or you can connect to a dsn.</p>
<p>Connecting in easy. Create a new instance of the SqlServerConnection class and pass it your connection string.</p>
<pre><code class="vbscript">
Dim sqlConnection : Set sqlConnection = New SqlServerConnection<br />
sqlConnection.ConnectionString = "SomeConnString"
</code></pre>
<p>You can enable tracing which provides the debug output you see in the below examples, detailing the sql generated and a dump of a Recordset if there is one returned. To enable it set the Trace property to True.</p>
<pre><code class="vbscript">
sqlConnection.Trace = True
</code></pre>
<h3>QueryToList</h3>
<p>
You can execute a  generic sql query that returns columns (with names or alias) and have them presented as objects with properties dynamically created at execution. These objects are just plain old objects. The type of the property is infered from the type given from the sql provider, or is returned as a string. You can use IsNull as per usual.</p>
</p>
<p>
Because you are getting an Object back that is not a value type you must use the 'Set' syntax on results back that are not value types.
</p>
<pre><code class="vbscript">
Set employeeList = sqlConnection.QueryToList("Select * From Employee")<br />
For Each employee in employeeList<br />
Response.Write(employeee.FirstName) ' Yes thats right you can just type in the field name!<br />
Response.Write(employeee.EmployeeId) ' OMG its already an integer!<br />
Response.Write(employeee.HireDtm) ' OMG its already an DateTime!<br />
Next
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select * From Employee</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>EmployeeId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>FirstName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>LastName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>CompanyId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>HireDtm</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Active</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Number</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Address</B><font color="grey">&lt;String&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jeremy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Child</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">465443534534534</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">2</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Peter</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Mason</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">435345435334411</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">3</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Leeroy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jenkins</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2004</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">False</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">342342322216678</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Sometown</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">4</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jack</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jackson</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2008</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">23432119876</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD></TR></TABLE>
<h3>QueryToList Fluent Sql</h3>
<pre><code class="vbscript">
Set sql = SqlBuilder.Select().From("Employee").Where("CompanyId = 1").Where("EmployeeId Between 1 And 300")<br />
Set employeeList = sqlConnection.QueryToList(sql)<br />
For Each employee in employeeList<br />
Response.Write(employeee.FirstName) ' Yes thats right you can just type in the field name!<br />
Next
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">SELECT * FROM Employee WHERE (CompanyId = 1) AND (EmployeeId Between 1 And 300)</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>EmployeeId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>FirstName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>LastName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>CompanyId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>HireDtm</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Active</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Number</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Address</B><font color="grey">&lt;String&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jeremy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Child</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">465443534534534</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">2</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Peter</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Mason</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">435345435334411</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">3</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Leeroy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jenkins</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2004</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">False</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">342342322216678</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Sometown</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">4</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jack</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jackson</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2008</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">23432119876</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD></TR></TABLE>
<hr><h3>QueryToArray</h3>
<pre><code class="vbscript">
Set employeeArray = sqlConnection.QueryToArray("Select * From Employee Where Active = 1")
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select * From Employee Where Active = 1</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>EmployeeId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>FirstName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>LastName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>CompanyId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>HireDtm</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Active</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Number</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Address</B><font color="grey">&lt;String&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jeremy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Child</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">465443534534534</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">2</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Peter</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Mason</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">435345435334411</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">4</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jack</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jackson</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2008</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">23432119876</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD></TR></TABLE>
<hr><h3>QuerySingle</h3>
<pre><code class="vbscript">
Set singleEmployee = sqlConnection.QuerySingle("Select * From Employee Where EmployeeId = 1")
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select * From Employee Where EmployeeId = 1</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>EmployeeId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>FirstName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>LastName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>CompanyId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>HireDtm</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Active</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Number</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Address</B><font color="grey">&lt;String&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jeremy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Child</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">465443534534534</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR></TABLE>
<hr><h3>QueryInteger</h3>
<pre><code class="vbscript">
employeeCount = sqlConnection.QueryInt("Select COUNT(1) From Employee")
Response.Write(employeeCount) ' Integer no need to use Set
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select COUNT(1) From Employee</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="1" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B></B><font color="grey">&lt;Int32&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">4</TD></TR></TABLE>
<hr><h3>QueryDateTime</h3>
<pre><code class="vbscript">
serverDate = sqlConnection.QueryDateTime("Select GETDATE() AS 'Today'")<br />
'serverDate = sqlConnection.QueryDateTime("Select CAST(NULL As DateTime) AS 'Today'")<br />
Response.Write(serverDate) ' DateTime no need to use Set
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select GETDATE() AS 'Today'</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="1" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>Today</B><font color="grey">&lt;DateTime&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">8/08/2012 4:03:18 PM</TD></TR></TABLE>
<hr><h3>IsNull Checking</h3>
<pre><code class="vbscript">
serverDate = sqlConnection.QueryDateTime("Select CAST(NULL As DateTime) AS 'Today'")<br />
Response.Write(IsNull(serverDate)) ' Should return True
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select CAST(NULL As DateTime) AS 'SomeNullDate'</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="1" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>SomeNullDate</B><font color="grey">&lt;DateTime&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD></TR></TABLE>
<hr><h3>Inserting</h3>
<pre><code class="vbscript">
sql = "INSERT INTO [Test].[dbo].[Employee] ([FirstName],[LastName],[CompanyId],[HireDtm],[Active],[Number],[Address])" & _<br />
"VALUES ('Testing','Person',1,GETDATE(),1,1234,'Sometown')"<br />
insertedEmployeeId = sqlConnection.ExecuteReturnIdentity(sql)<br />
Response.Write(insertedEmployeeId) ' Should return Integer
</code></pre>
<hr><h3>Updating</h3>
<pre><code class="vbscript">
Set employeeList = sqlConnection.QueryToList("Select * From Employee")<br />
For Each e in employeeList<br />
   e.Active = True ' Yes thats right you can just type in the field name!<br />
   sqlConnection.Update e, "Employee", "EmployeeId"<br />
Next
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select * From Employee WHERE Address = 'Someville'</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>EmployeeId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>FirstName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>LastName</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>CompanyId</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>HireDtm</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Active</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Number</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Address</B><font color="grey">&lt;String&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Jeremy</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Child</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">465443534534534</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR><TRBGCOLOR="#d3d3d3"><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">2</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Peter</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Mason</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1/01/2000</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">True</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">435345435334411</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">Someville</TD></TR></TABLE><div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">UPDATE Employee SET FirstName='Jeremy',LastName='Child',CompanyId=1,HireDtm=CONVERT(DATETIME,'1/01/2000',103) ,Active=1,Number=465443534534534,Address='Someville' WHERE EmployeeId=1;</div><div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">UPDATE Employee SET FirstName='Peter',LastName='Mason',CompanyId=1,HireDtm=CONVERT(DATETIME,'1/01/2000',103) ,Active=1,Number=435345435334411,Address='Someville' WHERE EmployeeId=2;</div>
<hr><h3>Updating/Inserting using ExpandoObject</h3>
<p>ExpandoObject is an object created from a Dictionary specification. This object is created at runtime to have the properties specified in the dictionary. Types are infered from the typename of the property value.</p>
<pre><code class="vbscript">
Dim inspectionDefinition: Set inspectionDefinition = Server.CreateObject("Scripting.Dictionary")<br />
inspectionDefinition.Add "InspectionTypeId", "Integer"<br />
inspectionDefinition.Add "Question", "String"<br />
inspectionDefinition.Add "ResponseTypeId", "Integer"<br />
inspectionDefinition.Add "LastEdited", "DateTime"<br />
inspectionDefinition.Add "LastEditedBy", "Integer"<br />
inspectionDefinition.Add "Weight", "Integer"       <br /><br />
' Use a class if you want more control / poco behaviour<br />
Dim newInspection : Set newInspection = ExpandoObject(inspectionDefinition)<br /><br />

sqlConnection.Update newInspection, "Inspection", "InspectionId"<br />
'sqlConnection.Insert newInspection, "Inspection", "InspectionId"

</code></pre>
<hr><h3>Deleting</h3>
<pre><code class="vbscript">
Set employeeList = sqlConnection.QueryToList("Select * From Employee")<br />
For Each e in employeeList<br />
   sqlConnection.Delete e, "Employee", "EmployeeId"<br />
Next
</code></pre>

<hr><h3>Insert via Stored Procedure</h3>
<pre><code class="vbscript">
Dim addEmployeeParams: Set addEmployeeParams = CreateObject("Scripting.Dictionary")<br />
    addEmployeeParams.Add "employeeId", 1234<br />
    addEmployeeParams.Add "someOtherId",4321<br />
    addEmployeeParams.Add "someDate",Now()<br />
    addEmployeeParams.Add "someString","Sometown"<br />
    
Dim addEmployee : addEmployee = sqlConnection.ExecStoredProcedureIdentity("[InsertNewEmployee]",addEmployeeParams)<br />
newEmployeeId = addEmployee
</code></pre>

<hr><h3>Execute</h3>
<pre><code class="vbscript">
sqlConnection.Execute("UPDATE WHERE .... ")
</code></pre>

<hr><h3>Execute Rows Affected</h3>
<pre><code class="vbscript">
Dim affected : affected = sqlConnection.ExecReturnRowsAffected("UPDATE WHERE .... ")
</code></pre>

<hr><h3>Data Types</h3>
<pre><code class="vbscript">
Set dataTypeTest = sqlConnection.QuerySingle("Select 123 As Int, 1.3 As Decimal, GETDATE() As Date, 'SomeString' As VarChar, Cast(0 As BIT) As Bit, CAST(NULL As BIT) As SomeBoolNull,CAST(NULL As Bigint) As SomeInt64Null, 5155474835647 As BigInt")
</code></pre>
<div style="background-color: #FCF8E3; color: #C09853; padding: 3px; border: 1px solid #FBEED5;">Select 123 As Int, 1.3 As Decimal, GETDATE() As Date, 'SomeString' As VarChar, Cast(0 As BIT) As Bit, CAST(NULL As BIT) As SomeBoolNull,CAST(NULL As Bigint) As SomeInt64Null, 5155474835647 As BigInt</div><TABLE width="850" style="border: 2px solid #17b;	margin: 0.3em 0.2em;"><TR><tr><TD colspan="8" style="color: white; background-color: #0066FF"><B>Dump of Recordset</B></TD><tr><TD style="color: black; background-color: #ddd"><B>Int</B><font color="grey">&lt;Int32&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Decimal</B><font color="grey">&lt;Decimal&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Date</B><font color="grey">&lt;DateTime&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>VarChar</B><font color="grey">&lt;String&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>Bit</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>SomeBoolNull</B><font color="grey">&lt;Boolean&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>SomeInt64Null</B><font color="grey">&lt;Int64&gt;</font></TD><TD style="color: black; background-color: #ddd"><B>BigInt</B><font color="grey">&lt;Decimal&gt;</font></TD></TR><TR><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">123</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">1.3</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">8/08/2012 4:03:18 PM</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">SomeString</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">False</TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;"><font color="green"><i>null</i></font></TD><TD style="vertical-align: top;border: 1px solid #aaa;	padding: 0.1em 0.2em;margin: 0;">5155474835647</TD></TR></TABLE>
<h2>Method List</h2>
<table width="850" height="484" border="1" cellpadding="0" cellspacing="0">
  <tr>
    <td width="298"><strong>Method Name</strong></td>
    <td width="372"><strong>Description</strong></td>
    <td width="172"><strong>Return Type</strong></td>
  </tr>
  <tr>
    <td><strong>RecordSet(string sql)</strong></td>
    <td>Returns an ADODB.Recordset object from the sql text provided</td>
    <td>ADODB.Recordset</td>
  </tr>
  <tr>
    <td><strong>BinaryStream(string sql)</strong></td>
    <td>Return sn ADODB.Stream object from the sql text provided</td>
    <td>ADODB.Stream</td>
  </tr>
  <tr>
    <td><strong>ExecReturnRowsAffected(string sql)</strong></td>
    <td>Returns the number of rows affected by the executed sql</td>
    <td>Integer</td>
  </tr>
  <tr>
    <td><strong>Execute(string sql)</strong></td>
    <td>Returns boolean if the executed sql completed without error</td>
    <td>Boolean</td>
  </tr>
  <tr>
    <td><strong>Open()</strong></td>
    <td>Opens the connection</td>
    <td>Void</td>
  </tr>
  <tr>
    <td><strong>SanitizeString(string input)</strong></td>
    <td>Returns a sanitized sql formatted string</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>Close()</strong></td>
    <td>Closes the connection</td>
    <td>Void</td>
  </tr>
  <tr>
    <td><strong>TraceRecordset(recordset rs)</strong></td>
    <td>Dumps the recordset into a HTML table</td>
    <td>Void</td>
  </tr>
  <tr>
    <td><strong>DataTypeNameFromAdoCode(int32 code)</strong></td>
    <td>Returns a friendly name for an ADO data type code</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>QueryToList(string sql)</strong></td>
    <td>Returns a list of dynamic objects based on the sql results</td>
    <td>System.Collections.ArrayList</td>
  </tr>
  <tr>
    <td><strong>Query(string sql)</strong></td>
    <td>Returns an array of dynamic objects based on the sql results</td>
    <td>Array</td>
  </tr>
  <tr>
    <td><strong>QuerySingle(string sql)</strong></td>
    <td>Returns a single dynamic object based on the sql results</td>
    <td>T is dynamic</td>
  </tr>
  <tr>
    <td><strong>QueryToString(string sql)</strong></td>
    <td>Retuns a string for the first column returned from the first row returned based on the sql results</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>QueryInt(string sql)</strong></td>
    <td>Retuns an integer for the first column returned from the first row returned based on the sql results</td>
    <td>Integer</td>
  </tr>
  <tr>
    <td height="45"><strong>QueryBoolean(string sql)</strong></td>
    <td>Retuns a boolean for the first column returned from the first row returned based on the sql results</td>
    <td>Boolean</td>
  </tr>
  <tr>
    <td><strong>QueryDateTime(string sql)</strong></td>
    <td>Return a DateTime for the  first column returned from the first row returned based on the sql results</td>
    <td>DateTime</td>
  </tr>
  <tr>
    <td><strong>QueryToArray(string sql)</strong></td>
    <td>Returns an array of dynamic objects based on the sql results (see Query)</td>
    <td>Array</td>
  </tr>
  <tr>
    <td><strong>QueryList(string sql)</strong></td>
    <td>Returns a list of dynamic objects based on the sql results (see QueryToList)</td>
    <td>System.Collections.ArrayList</td>
  </tr>
  <tr>
    <td><strong>ExecuteReturnIdentity(string sql)</strong></td>
    <td>Returns the @@IDENTITY from the executed sql</td>
    <td>Integer</td>
  </tr>
  <tr>
    <td><strong>GetNewGuid()</strong></td>
    <td>Returns a new GUID</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>GetCleanGuid()</strong></td>
    <td>Returns a new guid with only alpha numeric characters</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>CleanString(string input)</strong></td>
    <td>Returns a reasonably safe sql string</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>BoolToYesNo(object input)</strong></td>
    <td>Returns a &quot;Yes&quot; or &quot;No&quot; string based on a boolean or string value</td>
    <td>String</td>
  </tr>
  <tr>
    <td><strong>IsEmailAddress(string input)</strong></td>
    <td>Returns true or false if the email is valid based on the in built regex pattern</td>
    <td>Boolean</td>
  </tr>
</table>

<br />
<h2>Licence</h2>
<p>Copyright (C) 2012 Jeremy Child</p>
<p>Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:</p>
<p>The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.</p>
<p>THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.</p>