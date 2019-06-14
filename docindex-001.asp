<html>
	<head>
		<link rel="shortcut icon" href="favicon.ico" type="image/x-icon">
		<title>Document Control</title>
		<link href="https://fonts.googleapis.com/css?family=Merriweather+Sans" rel="stylesheet">
		<!--Bootstrap-->
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css">
		<!--jQuery-->
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
		<!-- Popper.JS -->
		<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js"></script>
		<!-- Bootstrap JS -->
		<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js"></script>
	</head>

	<body>
		<div class="container-fluid">
			<!--header-->
			<div class="header">
				<table width="100%">
				  <tr>
						<td><img src="images/co-branded-logo_300x69.jpg" width="300" height="69"></td>
						<td align="right"><img src="images/doccontrol.jpg" width="350" height="45"></td>
				  </tr>
				</table>
			</div>
			<%
				set deptDict = Server.CreateObject("Scripting.Dictionary")
				deptDict.add "Executive", 0
				deptDict.add "Sales & Marketing", 1
				deptDict.add "Engineering", 2
				deptDict.add "Material Control", 3
				deptDict.add "Manufacturing", 4
				deptDict.add "Welding", 420
				deptDict.add "Quality Assurance", 5
				deptDict.add "Administrative Support", 6
				deptDict.add "UHP Technology", 7
				deptDict.add "Customer Service", 8
				deptDict.add "Information & Communication", 9

				set catDict = Server.CreateObject("Scripting.Dictionary")
				catDict.add "Q", "Quality Policy"
				catDict.add "P", "Quality Procedure"
				catDict.add "N", "Quality Plan"
				catDict.add "S", "Specification"
				catDict.add "W", "Work Instruction"
				catDict.add "T", "Technical Bulletin"
				catDict.add "M", "Product Manual"
				catDict.add "F", "Form"
				catDict.add "A", "Addendum"
				catDict.add "R", "Reference"

				set dirDict = Server.CreateObject("Scripting.Dictionary")
				dirDict.add "Q", "K:\Department\doc_con\QSD_Policies\Released"
				dirDict.add "P", "K:\Department\doc_con\QSD_Procedures\Released"
				dirDict.add "N", "K:\Department\doc_con\QSD_Plans\Released"
				dirDict.add "S", "K:\Department\doc_con\QSD_Specifications\Released"
				dirDict.add "W", "K:\Department\doc_con\QSD_WorkInstructions\Released"
				dirDict.add "T", "K:\Department\doc_con\QSD_TechnicalBulletins\Released"
				dirDict.add "M", "K:\Department\doc_con\QSD_Manuals\Released"
				dirDict.add "F", "K:\Department\doc_con\QSD_Forms\Released"
				dirDict.add "A", "K:\Department\doc_con\QSD_Addendums\Released"
				dirDict.add "R", "K:\Department\doc_con\QSD_Reference\Released"

				category = UCase(Request("category"))
				department = Request("department")
				title = UCase(Request("title"))
				docnumber = UCase(Request("docnumber"))
				keyword = UCase(Request("keyword"))
				fulltext = UCASE(Request("fulltext"))				
			%>
			<form method="GET" action="docindex-001.asp">
			  <table bgcolor="#8BA8B7" width="100%">
					<tr>
					  <td><a href="index.html">&#x21E0; Go Back </a></td>
					  <td colspan="3" align="center"><h5>Search Quality System Documentation</h5></td>
					</tr>
					<tr>
					  <td valign="bottom">
							Search for<BR>
							<select size="1" name="Category" tabindex="1">
								<option <%if category = "" or UCase(category) = "ALL" then%>selected<%end if%>>All</option>
								<%for each cat in catDict%>
								<option <%if cat = category then%>selected<%end if%> value="<%=cat%>"><%=catDict(cat)%></option>
								<%Next%>
							</select>
					  </td>
						<td valign="bottom">By (Optional)<BR>
							<select size="1" name="Department" tabindex="2">
								<option <%if department = "ALL" then%>selected<%end if%>>All</option>
								<%	for each dept in deptDict %>
								<option <%if UCase(department) = UCase(dept) then%>selected<%end if%>><%=dept%></option>
								<%Next%>
							</select>
					  </td>
					  <td valign="bottom">
							Title Keyword<BR>
							<input type="text" name="title" value="<%=title%>" tabindex="3">
					  </td>
					  <td align="middle" rowspan="2" width="100%">Sort By:&nbsp;
							<select name="sort" tabindex="5">
								<option value="System.FileName" <%if lcase(request("sort")) = "System.FileName" then%>selected<%end if%>>Filename</option>
								<option value="System.Title" <%if lcase(request("sort")) = "System.Title" then%>selected<%end if%>>Title</option>
							</select>
							<input type="submit" value="Show Results" name="B1" style="font-weight: bold" tabindex="6">
					  </td>
					  <td rowspan="2" align="center" valign="middle">
							<div id="hitCount"><%if category <> "" then%>
								<p>Searching...<%end if%></p>
							</div>
					  </td>
					</tr>
					<tr>
					  <td>&nbsp;</td>
					  <td>Document Number / File Name<BR><input type="text" name="docnumber" size="20" value="<%=docnumber%>" tabindex="4"></td>
					  <!--<td><font size="-1">and/or full text<BR><input type="text" name="fulltext" size="20" value="<%=fulltext%>"></font></td>-->
					</tr>
			  </table>
			</form>
			<%
					if Request("Category") <> "" then
					Set Conn = Server.CreateObject("ADODB.Connection")
					Conn.ConnectionString = "Provider=Search.CollatorDSO.1;Extended Properties='Application=Windows';"
					Conn.Open

					sql = "SELECT System.FileName, System.Title, System.ItemPathDisplay, System.Keywords FROM SystemIndex "

					if category = "ALL" then
						sql = sql & " WHERE (DIRECTORY = 'K:\Department\doc_con\QSD_Policies\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Procedures\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Plans\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Specifications\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_WorkInstructions\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_TechnicalBulletins\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Manuals\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Forms\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Addendums\Released' OR DIRECTORY = 'K:\Department\doc_con\QSD_Reference\Released')"
					else
						sql = sql & " WHERE DIRECTORY = '" & dirDict(category) & "' "
					end if

					if department <> "ALL" then
						sql = sql & " AND System.FileName LIKE '_" & deptDict(department) & "%' "
					end if

					if title <> "" then
						sql = sql & " AND System.Title LIKE '%" & title & "%' "
					end if

					if docnumber <> "" then
						sql = sql & " AND System.FileName LIKE '%" & docnumber & "%'"
					end if

					sql = sql & " AND System.FileName NOT LIKE 'Thumbs.db'  AND System.FileName NOT LIKE '~%' AND System.FileName NOT LIKE '%Word%' AND System.FileName NOT LIKE '%Files%' AND System.FileName NOT LIKE '%Visio%' AND System.FileName NOT LIKE '%Design%' AND System.FileName NOT LIKE '%Fonts%' AND System.FileName NOT LIKE '%Links%' AND System.FileName NOT LIKE '%TTF%' AND System.FileName NOT LIKE '%bak%' AND System.FileName NOT LIKE '%zip%' AND System.FileName NOT LIKE '%txt%' AND System.FileName NOT LIKE '%indd%' ORDER BY " & Request("sort")
					Response.Write "<" & "!--" & sql & "--" & ">" & vbCrLf1
					set myRS = Conn.Execute(sql)
			%>
						
			<div align="center">			
				
				<table>
				  <tr>
						<th width="40" style="padding-left: 5; padding-right: 5"></th>
						<th style="padding-left: 5; padding-right: 5" bgcolor="#8BA8B7" align="left">Category</th>
						<th style="padding-left: 5; padding-right: 5" bgcolor="#8BA8B7" align="left">Department</th>
						<th style="padding-left: 5; padding-right: 5" bgcolor="#8BA8B7" align="left">FileName</th>
						<th style="padding-left: 5; padding-right: 5" bgcolor="#8BA8B7" align="left">Title</th>
				  </tr>
				  <%
						hitcount = 0
						do while not myRS.eof
							hitcount = hitcount + 1
							deptTemp = ""
							for each dept in deptDict
								if mid(myRS("System.FileName"), 2, 1) >= "0" and mid(myRS("System.FileName"), 2, 1) <= "9" then
									if deptDict(dept) = CInt(mid(myRS("System.FileName"), 2, 1)) then
										deptTemp = dept
										exit for
									end if
								else
									deptTemp = ""
									exit for
								end if
							Next
				  %>
				  
				  <tr onMouseOver="this.bgColor='#D9D6CF'" onMouseOut="this.bgColor=''">
						<td width="40" style="padding-left: 5; padding-right: 5"></td>
						<td style="padding-left: 5; padding-right: 5" nowrap><%=catDict(UCase(Left(myRS("System.FileName"), 1)))%></td>
						<td style="padding-left: 5; padding-right: 5" nowrap><%=deptTemp%></td>
						<td style="padding-left: 5; padding-right: 5" align="left" nowrap><a href="<%=mid(myRS("System.ItemPathDisplay"), 3)%>" target="_blank"><%=myRS("System.FileName")%></a></td>
						<td style="padding-left: 5; padding-right: 5" align="left" nowrap><%=left(myRS("System.Title"), 50)%></td>
				  </tr>
					<%
						myRS.MoveNext
							Loop
							myRS.Close
							set myRS = Nothing
							Conn.Close
							Set Conn = Nothing
						end if

						set dirDict = Nothing
						set deptDict = Nothing
						set catDict = Nothing
					%>
				</table>
			</div>
	</body>
</html>
<% if category <> "" then %>
<script Language="Javascript1.2" DEFER>
	hitCount.innerHTML = "<%=hitcount & " Matches"%>"
</script>
<% end if %>
