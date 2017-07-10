<%@ Language = VBScript %>
<%server.ScriptTimeout = 1800%>
<!--#include file="../Common/mainHeader.asp" -->
<!--#include file="../Shared/SharedPost.asp" -->
<%
Call CheckSession()

If Request.QueryString("tabIndex") <> "" Then
		Session("tabIndex") = Request.QueryString("tabIndex")
End If
'Response.Write(Session("tabIndex"))
firmcode=Session("FirmCode")
'Response.Write(firmcode)
Loginid = Session("LoginID")
'Response.Write(Loginid)
Dim Flag
Flag = Request.QueryString("flg")
%>
<html>
<head>
	<title>PTFM Foreign Associates</title>
	<meta charset="urf-8">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../intel/styles.css" rel="stylesheet" type="text/css">
	<link href="../Styles/styles.css" rel="stylesheet" type="text/css">
	<script language="Javascript" src="../Script/jquery-1.11.3.min.js"></script> 
	<link href="../Script/jquery-ui.css" rel="stylesheet" type="text/css">
    <script type="text/javascript" src="../Script/jquery-ui.js"></script> 
	<script language="JavaScript" src="../script/preloadImages.js"> </script>
	<script language="JavaScript" src="../Script/mm_menu.js"></script> 
	<script language="JavaScript" src="../Script/params.js"></script> 
	<script language="JavaScript" src="../script/navBarRenderer.js"></script>
	<script type="text/javascript">
		    	function SubmitForm() {
		    		var imgpath = document.getElementById("fileUpload").value;
		    		if(imgpath=="")
		    		{
		    			//alert("Upload Your File..");
		    			//document.form.fileUpload.focus();
		    			document.getElementById("lblError").innerHTML = "No file was chosen before clicking on Upload button. Please chose a file first.";
		    			return;
		    		}

			        var allowedFiles = [".txt"];
			        var fileUpload = document.getElementById("fileUpload");
			        var lblError = document.getElementById("lblError");
			        var regex = new RegExp("([a-zA-Z0-9\s_\\.\-:])+(" + allowedFiles.join('|') + ")$");
			        if (!regex.test(fileUpload.value.toLowerCase())) {
			            lblError.innerHTML = "Please upload files having extensions: <b>" + allowedFiles.join(', ') + "</b> only.";
			            return;
			        }
					/*if (fileUpload.files[0].size > 10485760){
						lblError.innerHTML = "File size is more than 10 MB.";
						return;
					} Commented temporarily as showing error on IE*/
			        lblError.innerHTML = "";
			        //return;

			        lblError.innerHTML = "File Upload in Progress.......";
			        document.form.action = "ImportChaseData.asp";
			        document.form.submit();
					
		    	}
				function OpenPopUp() {
				var url="help.html";
				my_window = window.open(url,'',"scrollbars=yes,location=0,menubar=no,toolbar=no,titlebar=no,width=600,height=600,resizable=no");
				}
	</script>
	<script>
//	    $(document).ready(function() {
//	        $("#fileUpload").click(function() {
//	            $("#lblError").hide();
//	        });
//	    });
	</script>
	<script>
	    $(function() {
	        $(document).tooltip();
	    });
    </script>
</head>
<body style="padding-top: 0px;" leftmargin="0" topmargin="0">
 <% 	  	   
	  if UCase(trim(Session("firmCode"))) = "WGSB" then
			strReturnPath = "dashboard.asp"
				call OrgHeader()						 
	  End if		
	 %> 
	<script language="JavaScript">
	var navText = new Array();
	var navUrl = new Array();

	navText[0] = "AMEX/MC Reconciliation";
	navText[1] = "File Transaformation";
	navUrl[0] = "";	
	navUrl[1] = "";	

	renderNavigationBar(navText, navUrl,"Invoice Processing");
	</script>
	<div class="container">
		<div><!--Heading-->
			<table width="60%" border="0" align="center" cellpadding="0" cellspacing="1">
				<div align="center" class="clsdv">
			        <%If(Flag = "1") Then%>			   
			        <p class="normaltextwelcome1">Your File Has Been Sucessfully Uploaded.</p>
			        <% End If %>
			    </div>
				<tr>
    				<td height="20" colspan="4" class="normaltext">&nbsp;</td>
  				</tr>
				<tr>
					<td height="20" colspan="4" bgcolor="6E6E6B" class="labledesidebar">&nbsp;<img src="../images/arrow.gif" width="13" height="12" align="absmiddle"> &nbsp;Upload Chase File</td>
				</tr>
				<tr>
    				<td colspan="4" bgcolor="#C9C9C9"><img src="../images/transpacer.gif" width="8" height="2"></td>
  				</tr>
  				<tr>
    				<!--<td colspan="4" class="normaltext"><font color=maroon>The file size should be less than 10 MB.</td>-->
  				</tr>
  				<tr>
    				<td colspan="4" class="normaltext">
    					<div id="lblError" style="color: red;">&nbsp;</div>
    				</td>
  				</tr>
  				<tr>
	  				<td colspan="4" class="BlackLiner"></td>
  				</tr> 
			</table>
		</div><!--Heading End-->
		<form name="form" action="upload.asp" method="POST" enctype="multipart/form-data" role="form">

			<table width="60%" border="0" align="center" cellpadding="0" cellspacing="1">
				<tr>
    				<td height="22" colspan="2" class="Ntablerow1bold">&nbsp;Upload Chase File</td>
  				</tr>
				<tr>
					<td width="20%" class="NcontextlabelboldDark">&nbsp;File:</td>
					<td class="NcontextlabelboldNormal"><input type="file" id="fileUpload" name="file" size="64" maxlength=1000></td>
				</tr>
				<tr>
	  				<td colspan="2" class="BlackLiner"></td>
	  			</tr>
					
				<!--<tr>
					<td colspan="3" class="normaltext">
						<span id="lblError" style="color: red;"></span>
					</td>
				</tr>-->
				<tr>
					<td colspan="3" align="right">
						<a href="JavaScript:SubmitForm();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','../images/upload_over.gif',1)">
							<img src="../images/upload.gif" name="Image2" width="75" height="22" border="0"></a>
						<a href="../dashboard.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','../images/cancel_try_over.gif',1)">
							<img src="../images/cancel_try.gif" name="Image17" width="75" height="22" border="0"></a>
							<!-- <a href="Javascript:OpenPopUp()"><img src="../images/Help.jpg" border="0" alt="Click to see instruction" title="Click to see instruction"></a>&nbsp;&nbsp;&nbsp; -->
					</td>
				</tr>


				<!--<tr>
					<td colspan="2" align="right">
						<button type="submit" class="btn btn-success" onclick="return ValidateExtension()">Upload</button>
						&nbsp;
						<button type="reset" class="btn btn-danger">Reset</button>
					</td>
				</tr>-->
			</table>
		</form>
	</div>
</body>
</html>