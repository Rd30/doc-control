<!DOCTYPE html>
<!--Document Control page-->
<html  lang="en-us">
	<head>
		<!-- #include file = "../gp-slo/common/gp-sloHead.html" -->
	</head>
	<body>
		<!-- Dark overlay element -->
		<div class="overlay" id="overlay"></div>

		<!--NavBar/Header-->
		<div class="all-gp-sloHeader" id="docControlIndexHeader"><!-- #include file = "../gp-slo/common/gp-sloHeader.html" --></div>			
		

		<!--SideBar-->
		<!-- #include file = "../gp-slo/common/gp-sloSidebar.html" -->

		<div class="gp-slo-container container-fluid" id="docCtrlAspContainer">
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
			<form method="GET" action="gp-slo_docindex.asp" class="docCtrlForm">
				<div class="sel-options-div" id="aspSelOptionsDiv">
					<div>
						<h6>Search Quality System Documentation</h6>
						<a href="gp-slo_index.php" style="color: darkblue;">&#x21E0; Go Back </a>
					</div>					
					<div class="row row-container">
						<div class="doc-ctrl-category col-md-3 col-sm-3">
							Search for
							<br>
							<select size="1" name="Category">
								<option <%if category = "" or UCase(category) = "ALL" then%>selected<%end if%>>All</option>
								<%for each cat in catDict%>
								<option <%if cat = category then%>selected<%end if%> value="<%=cat%>"><%=catDict(cat)%></option>
								<%Next%>
							</select>
						</div>
						<div class="doc-ctrl-department col-md-3 col-sm-3">
							By (optional)
							<br>
							<select size="-1" name="Department">
								<option <%if department = "ALL" then%>selected<%end if%>>All</option>
								<%	for each dept in deptDict %>
								<option <%if UCase(department) = UCase(dept) then%>selected<%end if%>><%=dept%></option>
								<%Next%>
							</select>
							<br>
							Document Number / File Name
							<br>
							<input type="text" name="docnumber" value="<%=docnumber%>">
						</div>
						<div class="doc-ctrl-title col-md-3 col-sm-3">
							Title keyword
							<br>
							<input type="text" name="title" value="<%=title%>">
						</div>
						<div class="doc-ctrl-sort col-md-3 col-sm-3">
							Sort By
							<br>
							<select name="sort">
									<option value="System.FileName" <%if lcase(request("sort")) = "System.FileName" then%>selected<%end if%>>Filename</option>
									<option value="System.Title" <%if lcase(request("sort")) = "System.Title" then%>selected<%end if%>>Title</option>
							</select>
							<input type="submit" value="Show Results" name="B1" style="font-weight: bold">
							<br>
							<br>
							<div id="hitCount" style="color: firebrick;"><%if category <> "" then%>
								Searching...<%end if%>
							</div>
						</div>
					</div>
				</div>
			</form>
			<br>
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
		<div class="container search-results" id="searchResults">
			<table class="table table-sm table-hover">
				<thead class="thead-light">
			    <tr>
			      <th scope="col">Category</th>
			      <th scope="col">Department</th>
			      <th scope="col">File Name</th>
			      <th scope="col">Title</th>
			    </tr>
  			</thead>
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
				<tbody>
					<tr>
						<td><%=catDict(UCase(Left(myRS("System.FileName"), 1)))%></td>
						<td><%=deptTemp%></td>
						<td><a style="color: darkblue;" href="<%=mid(myRS("System.ItemPathDisplay"), 3)%>" target="_blank"><%=myRS("System.FileName")%></a></td>
						<td><%=left(myRS("System.Title"), 50)%></td>
					</tr>
				</tbody>
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
		</div>

		<script type="text/javascript" src="http://nd-wind.entegris.com/gp-slo/gp-slo.js"></script>
		<script type="text/javascript">
	      $(document).ready(function () {
	        $('#pageTitleDiv').html("");
	        $('#pageTitleDiv').html("<h5>Document Control</h5>");
			$('#shortPageTitleDiv').html("");
			$('#shortPageTitleDiv').html("<h5>Doc Ctrl.</h5>");
	      })
	  </script>

  </body>
</html>
<% if category <> "" then %>
<script Language="Javascript1.2" DEFER>
	hitCount.innerHTML = "<%=hitcount & " Matches"%>"
</script>
<% end if %>
