<%
	'=================================================================
	'SET objFIL	=Server.CreateObject("Filgifts.mba_conn")
	'=================================================================
%>
<script language="javascript">
	function opencontribute()
	{
		window.open ("/ffp/contribute.html","","left=200,top=300,width=500,height=250");
	}
</script>

<table width="100%" cellspacing="2" cellpadding="0" border="0">
  <tr> <% xflag=0 %> 
    <td width="98%"><!--#include virtual="/gapi_inc/greetings.asp"--></td>
<!--BODY CONTENT-->
    <td width="221" align="right" valign="top"> 
      <table width="171" border="0" cellspacing="0" cellpadding="0">
        <tr> <%
			                if session("_xlogid")="" then
								Response.Write "<td width='97'><a href='/ssl/default.asp'><img src='/images/homepage/mother02/mothloginhere.gif' border='0'></a></td><td width='96'><a href='https://www.filgifts.com/ssl/account.asp?xnew=1'><img src='/images/homepage/mother02/mothjoinnow.gif' border='0'></a></td>"
								
			                else
								Response.Write "<td>&nbsp</td><td>&nbsp</td>"
			                end if
			                %> 
			<td colspan="2" align="right" valign="top"><!--<img src="/images/homepage/curve.gif" width="28" height="24">--></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<!--PRODUCT OF THE SEASON--> 
<table width="459" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
    <td width="245" align="center" valign="center">
<center><img src="/images/homepage/xmas02/xmasgreeting.gif" width="449" height="25" border="0"></center>
 </td>
  </tr>
</table>
<table width="459" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
 <td width="245" align="center" valign="center" height="10">
</td>
  </tr>
  <tr> 
    <td width="220" align="center" valign="center"><img src="/images/homepage/xmas02/santacolor3.jpg" width="220" height="172"></td>
    <td height="230" align="center" rowspan="2" valign="top" width="229"> <%
								'set xcmd=objFIL.gerbel_sp("js_GetRandomProductoftheSeason")
								set xcmd=gerbel_sp(xconn,"js_GetRandomProductoftheSeason")
								set xrst=xcmd.execute
							
								Response.Write "<a href=view.asp?xitem=" & xrst("gapi") & "><img border=0 width=230 height=230 src='\images\product\small\" & xrst("kbpimg") & ".jpg' alt=" & xrst("kbptitl") & "></a>"
								xdestroy
							%> </td>
  </tr>
  <!--	<tr>
		<td height="60" align="center"><a href="http://202.163.221.226:8181//ss?click&Fil001&3d6dd228"><img src="/images/homepage/july02/julyfindout.gif" border="0" alt="Click here!" width="112" height="20" valign="middle"></a> 
    	</td>
	</tr>--> <!--	<tr>
    <td colspan="2"><marquee><font face="verdana, helvetica, arial" size="1"><b> 
      <font color="Red">September 8</font> is GRANDPARENTS' DAY!</B></font>
	  </marquee>
	</td>
	</tr>--> 
</table>
<br>
<!-- FILGIFTS FEATURE-->
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
    	
    <td align="right" height="50"> <img src="/images/homepage/xmas02/xmastreats.gif" align="left"></td>
	    </tr>
	<tr>
		
</table>
<table width="410" border="0" cellspacing="2" cellpadding="2" align="center">
		<tr>
			<td width="140" valign="top"><center>
        	<a href="http://202.163.221.226:8181/ss?click&Fil001&3dd071af">
			<img src="/images/homepage/xmas02/chairmanoftheboard.jpg" border="0" alt="Chairman of the Board" width="104" height="104"></a> 
      		</center>
			</td>
			<td width="143" valign="top"><center>
        <a href="http://202.163.221.226:8181/ss?click&Fil001&3da24832"> <img src="/images/homepage/oct02/foodforthegods.jpg" width="104" height="104" border="0" alt="Food for the Gods"></a> 
      </center>
			</td>
			<td width="140" valign="top"><center>
        	<a href="http://202.163.221.226:8181/ss?click&Fil001&3dc61512">
			<img src="/images/homepage/oct02/rosesarered.jpg" border="0" width="104" height="104" alt="Farm Fresh Assorted Roses"></a> 
      		</center>
			</td>
		</tr>
		<tr>
			<td width="140" valign="top"><center>
        	<font face="verdana" size="1">
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3dd071af" style="text-decoration:none">
			Chairman of the Board</a><br>
        	<b><font color="red">$84.04</font><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3dd07adf">More Gift Baskets</a></b>
			</font>
	      	</center>
			</td>
			<td width="143" valign="top"><center>
        	<font face="verdana" size="1">
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3da24832" style="text-decoration:none">
			Food for the Gods</a><br>
        	<b><font color="red">$15.33</font><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3d6dd118">More Pastries and Cakes</a></b>
			</font>
			</center>
			</td>
			<td width="140" valign="top"><center>
        	<font face="verdana" size="1">
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3dc61512" style="text-decoration:none">
			Roses are Red</a><br>
        	<b><font color="red">$22.77</font></b> <br>
        	<a href="http://202.163.221.226:8181/ss?click&Fil001&3d6dd077"><b>More Flowers</b></a> 
        	</font> 
      		</center>
			</td>
		</tr>
		<tr>
			<td width="140" valign="top"><center><br>
        	<a href="http://202.163.221.226:8181/ss?click&Fil001&3da2486a">
			<img src="/images/homepage/oct02/DN0108.jpg" width="104" height="104" border="0" alt="Little Italy"></a> 
      		</center>
			</td>
		    <td width="143" valign="top"><center><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3d9a9163">
			<img src="/images/homepage/oct02/fortantwithchardonnay.jpg" border="0" width="104" height="104" alt="Cabernet Sauvignon and Chardonnay"></a>
			</center>
			</td>
		    <td width="140" valign="top"><center><br>
        	<a href="http://202.163.221.226:8181/ss?click&Fil001&3dd071d9"> <img src="/images/homepage/xmas02/nokia8855.jpg" border="0" width="104" height="104" alt="Nokia 3610"></a> 
      		</center>
			</td>
		</tr>
		<tr>
			<td valign="top" width="140"><center>
        	<font face="verdana" size="1">
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3da2486a" style="text-decoration:none">
			Daisy Drop Pearl Necklace</a><br>
        	<b><font color="red">$123.66</font><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3d75b98f">More Jewelry</a></b>
			</font> 
      		</center>
			</td>
			<td width="143" valign="top"><center>
        	<font face="verdana" size="1"> <a href="http://202.163.221.226:8181/ss?click&Fil001&3d9a9163" style="text-decoration:none"> 
        	French Wines</a><br>
        	<b><font color="red">$9.29</font><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3d6de391">More Fine Wines</a></b>
			</font> 
		    </center>
			</td>
			<td width="140" valign="top"><center>
        <font face="verdana" size="1"> <a href="http://202.163.221.226:8181/ss?click&Fil001&3dd071d9" style="text-decoration:none"> 
        Nokia 8855</a><br>
        <b><font color="red">$445.54</font><br>
			<a href="http://202.163.221.226:8181/ss?click&Fil001&3d6de44f">More Mobile Phones</a></b>
			</font>
	        </center>
			</td>
		</tr>
	</table>
<br>
<!-- COMING SOON 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
	<td align="right"><img src="/images/homepage/mother02/gradcomin
<table width="400" border="0" cellspacing="2" cellpadding="2" align="center">gsoon.gif" width="175" height="67" alt="Coming Soon"></td>
	<td width="99%" background="/images/homepage/line_orange.jpg">&nbsp;</td>
  </tr>
</table>
	<tr>
    	<td><font face="Verdana, Arial, Helvetica, sans-serif" size="-2" color="#000000"> 
		Watch out for luscious <b>Mrs. Fields goodies, Black and Decker do-it-yourself Gadgets, 
		silver jewelry and more restaurant Gift Certificates!  <font color="red">Coming very soon!!!</font></b></font> 
    </td>
	</tr>
</table>
<br>-->
<!-- FRESH FROM PINAS --> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
    <td width="60%" align="center"  height="50"><img src="/images/homepage/xmas02/ffp.gif"></td>
<td width="40%"></td>
</tr>
</table>
<table width="400" border="0" cellspacing="2" cellpadding="2" align="center">
	<tr>
    <td width="140" align=center><a href="http://202.163.221.226:8181/ss?click&Fil001&3dcf16d3"><img src="/images/homepage/xmas02/starsparol.jpg" border=0 alt="Stars within Stars" width="135" height="109"></a>
	<br><font face="Verdana, Arial, Helvetica, sans-serif" size="-2" align="left" color="#000000">Image (c) <a href="http://sim.soe.umich.edu/parol/" target="_blank">Denise Nacu</a></font>
    </td>
    <td width="260" valign="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="-2" color="#000000">It's 
      Christmas time in Manila! Find out what makes Manila's Christmas so colorful 
      and bright as we count down the traditions that make Christmas even more 
      meaningful. </font><font face="Verdana" size="1" color="black"> <br>
      <br>
      <a href="http://202.163.221.226:8181/ss?click&Fil001&3dcf16d3">More</a><br>
      <br>
      </font> </td>
	</tr>
	<tr>
	<td colspan="2">
	<font face="verdana" size="1">
	<a href="/ffp/ffparchives.asp">More on Fresh From Pinas</a>
	</font></td>
	</tr>
</table>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
    <td width="65%" align="center" height="25"><img src="/images/homepage/xmas02/balikbayanch.gif"> 
      </td>
	<td width="40%"></td>
</tr>
</table>
<table width="400" border="0" cellspacing="2" cellpadding="2" align="center">
	<tr>
    <td valign="center"><font face="Verdana" size="2" color="black">Care to share 
      on your trip back home? Post it on our <a href="/ffp/chronicles/read.asp"> 
      <b>Balikbayan Chronicles!</b></a><br>
      <br>
	  <font face="verdana" size="1" color="black">
		<!--#include virtual="/ffp/chronicles/firstmsg_hp.asp"-->
		<p align="right"><a href="/ffp/chronicles/read.asp">Read on</a></p>
	  </font>
      </font>
	</td>
	</tr>
</table>
