<% @CODEPAGE="65001" language="VBScript" %>
<!-- #include virtual="/zincludethema/serverfn.asp"-->
<!-- #include virtual="/zincludethema/serverchk.asp"-->
<%

Session.CodePage = 65001

Response.CharSet = "UTF-8"

os=request("os")

Session("os")=os

call isPc3()

iday = sf_strdt(now,"yyyymmdd")

idayif=Mid(iday,5,2) & "월" & Mid(iday,7,2) & "일"

Set objconn = Server.createObject("ADODB.Connection")
objconn.open Application("objconn_ConnectionString_Main")

Set objconnNeo=Server.createObject("ADODB.Connection")
objconnNeo.open Application("objconn_ConnectionString")

Set objRec = Server.CreateObject ("ADODB.Recordset")
Set objCount = Server.CreateObject ("ADODB.Recordset")
Set objRecCode = Server.CreateObject ("ADODB.Recordset")
Set objconnCode = Server.createObject("ADODB.Connection")

Dim thema_Name(3)
Dim thema_Rank_Index(6) 
Dim thema_Rank_Flow(6)%>
<!DOCTYPE html>
<html lang="ko">
<head>
	<meta charset="utf-8">
	<meta name="format-detection" content="telephone=no">
	<meta name="theme-color" content="#121213">
	<meta content="width=device-width,initial-scale=1.0,minimum-scale=1.0,maximum-scale=1.0,user-scalable=no" name="viewport">
	<meta name="subject" content="">
	<meta name="description" content="">
	<meta name="keywords" content="">
	<meta property="og:url"	content="">
	<meta property="og:type" content="website">
	<meta property="og:title" content="">
	<meta property="og:description"	content="">
	<meta property="og:image" content="../images/sns_link.png">
	<title>시그널노트</title>
	<link rel="shortcut icon" href="../images/favicon.png" type="image/x-icon">
	<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap" rel="stylesheet">
	<link rel="stylesheet" type="text/css" href="../newcss/swiper-bundle.min.css">
	<link rel="stylesheet" type="text/css" href="../newcss/jquery.modal.css">
	<link rel="stylesheet" type="text/css" href="../newcss/style.css">
	<script type="text/javascript" src="../newjs/jquery-1.11.1.min.js"></script>
	<script type="text/javascript" src="../newjs/swiper-bundle.min.js"></script>
	<script type="text/javascript" src="../newjs/jquery.modal.min.js"></script>
	<script type="text/javascript" src="../newjs/script.js"></script>
</head>
<body class="front-background">
<div id="wrapper">

	<div class="content">	
		<div class="tab">
			<div class="swiper-container">
				<ul class="swiper-wrapper">
					<li class="swiper-slide active"><a href="../goodtiming/home.asp">HOME</a></li>
					<li class="swiper-slide"><a href="../vvip/themastock5.asp">5%매매</a></li>
					<li class="swiper-slide"><a href="../vvip/closeStock.asp">종가매매</a></li>
					<li class="swiper-slide"><a href="../vvip/themastock20.asp">20%매매</a></li>
					<li class="swiper-slide"><a href="../powerPriceSearch/powerSearch.asp">세력가 검색</a></li>
					<li class="swiper-slide"><a href="../vvip/morninginfo.asp">7AM</a></li>
					<li class="swiper-slide"><a href="../about/about.asp">테마지수소개</a></li>
					<li class="swiper-slide"><a href="../about/qa.asp"">Q&A</a></li>
				</ul>
			</div>
		</div>	
		<script>
		var tabIndex = $(".tab li.active").index();
		var tabswiper = new Swiper('.tab .swiper-container', {
			slidesPerView: 'auto',
			slideToClickedSlide: true,
			preventClicksPropagation: false,
			initialSlide: tabIndex
		});
		</script>
		<div class="panel">
		<%strsql="select count(*) from stockweight"
		Set Rs=objconn.Execute(strsql)
		totNum=Rs(0)
		endNum=totNum-weightdayCnt
		startNum=Endnum-40
		strsql="select top 1 regdt from stockweight order by sn desc"
		Set Rs=objconn.Execute(strsql)
		weightRegdt=Rs(0)
		Set Rs=Nothing
		cnt=0
		nowstate=0        ' 현재 변곡
		continueState=0   ' 현재 벽곡이 몇일째인지.
		calcState=0		  ' 전일대비 비중이 얼마나 빠졌는지 늘었는지.
		lsatWeight=0	  ' 전일 비중	
		nowWeight=0       ' 현재 비중 
		strState=""		  ' 변곡 표시 문자열
		viewCnt=0		  ' 데이타를 총 40개 읽어서 최근 20개만 표시하기 위한 변수
		strsql="select top 40 * from stockweight where sn<=" & endnum & " and sn>" & startNum & " order by sn"
	'	Response.write strsql
		objRec.Open strsql, objconn%>
			<div class="heading">
				<h2>주식비중 <a href="#info" rel="modal:open" class="icon-info"><img src="../images/icon_info.png" alt="주식비중이란?"></a> <small>Updated <%=Mid(weightRegdt,6,2) & "." & Mid(weightRegdt,9,2) & " " & Mid(weightRegdt,12,12)%></small></h2>
			</div>
			<div class="stock-weather">
				<div class="swiper-container">
					<div class="swiper-wrapper">
						<%Do Until objRec.EOF
							viewCnt=viewCnt+1
							lastWeight=nowWeight
							nowWeight=objrec("highRiskrlt2")
							calcState=nowWeight-lastWeight
							
							If calcState>0 Then ' 전일대비 상승 
								If cnt<0 Then 
									cnt=0
								End If
						'		Response.write "상"
								cnt=cnt+1
						'		If continueState>0 Then 									
									continueState=continueState+1
						'		Else
						'			continueState=1
						'		End If 
							ElseIf calcState<0 Then '전일대비 하락
								If cnt>0 Then 
									cnt=0
								End If 
						'		Response.write "하"
								cnt=cnt-1
						'		If continueState<0 Then 
									continueState=continueState-1
						'		Else
						'			continueState=-1
						'		End If 
							Else 
								If nowstate>0 Then 
									continueState=continueState+1
								ElseIf nowstate<0 then
									continueState=continueState-1
								End If 
							End If 	
					'		Response.write nowstate & " " & objrec("highRiskrlt2") & " " & cnt  
							If (calcState>5 And continueState<-1) Or (calcState>10 And continueState=-1) Or (nowstate=-1 And cnt>=2)  Then 
								strState="매수 변곡"
								nowstate=1
								continueState=0
							ElseIf ((calcState<-5 And continueState>1) Or (calcState<-10 And continueState=1) Or (nowstate=1 and cnt<=-2)) And objrec("highRiskrlt2")<90  Then
								strState="매도 변곡"
								continueState=0
								nowstate=-1
							Else
								strState=""
							End If 
							
						
						If viewCnt>20 then%>
						<div class="swiper-slide">
							<div class="box">
								<%daystr=Mid(objrec("regdt"),6,2) & "월" & Mid(objrec("regdt"),9,2) & "일"
								If daystr=idayif then%>
									<span class="today">오늘 <%=strState%></span>
								<%else%>
									<%If Len(strState)>2 Then %>
										<span class="today"><%=strState%></span>
									<%End if%>
								<% End if
								weekstr=sf_strweekday(Mid(objrec("regdt"),1,10))%>
								<div class="date"><%=daystr & "(" & weekstr & ")"%></div>
								<div class="icon">
									<%If objrec("highRiskrlt2")>=80 Then%>
										<img src="../images/wether_1.png" alt="">
									<%ElseIf objrec("highRiskrlt2")>=60 And objrec("highRiskrlt2")<80 then%>
										<img src="../images/wether_2.png" alt="">
									<%ElseIf objrec("highRiskrlt2")>=40 And objrec("highRiskrlt2")<60 then%>
										<img src="../images/wether_3.png" alt="">
									<%ElseIf objrec("highRiskrlt2")>=20 And objrec("highRiskrlt2")<40 then%>
										<img src="../images/wether_4.png" alt="">
									<%else%>
										<img src="../images/wether_5.png" alt="">
									<%End if%>
								</div>
								<div class="per">
									<span class="text-danger"><%=objrec("highRiskrlt2")%>%</span>
								</div>
							</div>
						</div>
						<%End If 
						objrec.movenext						
						loop%>
					</div>					
				</div>
				<div class="swiper-pagination"></div>
			</div>
			<script>
			var slideNum = $(".stock-weather .swiper-slide").length;
			var weatherswiper = new Swiper('.stock-weather .swiper-container', {
				slidesPerView: 'auto',
				centeredSlides: true,
				spaceBetween: 7,
				grabCursor: true,
				//freeMode: true,
				initialSlide: slideNum,
				pagination: {
					el: '.stock-weather .swiper-pagination',
					clickable: true
				}
			});
			</script>
		<%objrec.close%>
		</div><!-- end panel -->
		<%strsql="select top 3 * from themaindex_total order by rank"
		objRec.Open strsql, objconn%>
		<div class="panel bg">
			<div class="heading clearfix">
				<h2 class="fl">실시간 HOT 테마</h2>				
				<button type="button" class="fr btn-sort" title="정렬"></button>
			</div>
			<div class="stock-hot-note">
				<div class="swiper-container">
					<div class="swiper-wrapper">
					<%jjj=1
					Do Until objrec.EOF%>
						<div class="swiper-slide">
							<a href="../stock/themainfo.asp?themagroup=<%=objrec("themagroup")%>" class="box">
								<div class="tit"><%=objrec("themagroup")%></div>
								<%thema_Name(jjj)=objrec("themagroup")
								ratio=objrec("ratio")
								ratio=int(ratio*100)/100
								If ratio>=0 Then%>
									<div class="text-danger"><span class="pt"><%=objrec("C")%>pt</span> 
										<span class="per">▲ <%=ratio%>%</span>
									</div>
								<%Else
									ratio=Abs(ratio)%>
									<div class="text-primary"><span class="pt"><%=objrec("C")%>pt</span> 
										<span class="per">▼ <%=ratio%>%</span>
									</div>
								<%End if%>
								<div class="tt">테마 강도</div>
								<%thema_flow=objrec("rankflow")
								j=1
								m=1
								For i=0 To Len(thema_flow)
									If Mid(thema_flow,i+1,1)="(" Then 
										thema_rank_index(j)=Mid(thema_flow,i+2,3)
										j=j+1
									End If
									If Mid(thema_flow,i+1,1)=")" Then 
										thema_Rank_Flow(m)=Mid(thema_flow,i+2,1)
										m=m+1
									End If 			 
								Next %>
								<div class="theme-strength">
									<ul>
										<%For i=1 To 6
										rank=CInt(thema_rank_index(i))
										select Case thema_Rank_Flow(i)								
										Case "↑":%>
											<li>
												<div class="item up full">
													<img src="../images/arrow_up_w.png" alt=""><%=rank%>위
												</div>
											</li>
										<%Case "↗":%>
											<li>
												<div class="item up">
													<img src="../images/arrow_up.png" alt=""><%=rank%>위
												</div>
											</li>
										<%Case "↔":%>
											<li>
												<div class="item up">
													<img src="../images/arrow_same.png" alt=""><%=rank%>위
												</div>
											</li>
										<%Case "↘":%>
											<li>
												<div class="item down">
													<img src="../images/arrow_down.png" alt=""><%=rank%>위
												</div>
											</li>
										<%Case "↓":%>
											<li>
												<div class="item down full">
													<img src="../images/arrow_down_w.png" alt=""><%=rank%>위
												</div>
											</li>
										<%End select%>
										<%next%>
									</ul>
								</div>
								<div class="tt">매수 강도</div>
								<%tot=int(objrec("totalscore")/100)
								e=9
								If tot>=9 Then
									tot=9
									e=0
									harfcnt=0
								Else 
									harfcnt=objrec("totalscore")-tot*100
								End If%>
								<div class="buy-strength">
									<span class="dots">
									<%For i=1 To tot
										e=e-1%>
										<span class="dot full"></span>
									<%Next
									If harfcnt>=50 Then 
										e=e-1%>
										<span class="dot half"></span>
									<%End if%>
									<%For i=1 To e%>
										<span class="dot"></span>
									<%next%>
										</span>
									</span>
									<span class="num"><%=objrec("totalscore")%></span>
								</div>
							</a>
						</div>
						<%jjj=jjj+1
						objrec.movenext
						loop%>
					</div>
				</div>
			</div>
			<%objrec.close%>
			<div class="stock-comps">
				<div class="swiper-container">
					<div class="swiper-wrapper">
					<%For jjj=1 To 3%>
					<!-- #include virtual="/zincludethema/stockChoice.asp"-->
					<%Set objconnCode = Server.createObject("ADODB.Connection")
					connstr="objconn_Connectionstring_Num" & themagroup
					objconnCode.open Application(connstr)
					If thema_Name(jjj)="항공" Then 
						strsql="select top 3 * from themacode_marketprice where themagroup like '%" & thema_Name(jjj) & ";%' and not	themagroup like '%우주항공;%' order by ratio desc"
					else
						strsql="select  top 3 * from themacode_marketprice where themagroup like '%" & thema_Name(jjj) & ";%' order by ratio	desc"
					End If 
					objRec.Open strsql, objconnCode%>
						<div class="swiper-slide">
							<ul>
							<%Do Until objrec.EOF%>
								<li>
									<a href="../stock/stockdetail.asp?themagroup=<%=thema_Name(jjj)%>&codeName=<%=objrec("codeName")%>" class="stock-comp circle">
										<div class="in">
											<div class="name"><%=objrec("codeName")%></div>
											<%ratio=objrec("ratio")
											If ratio>=0 then%>
												<div class="pt text-danger"><%=mid(formatcurrency(objrec("nowprice")),2,10)%></div>
												<div class="per text-danger"><small>▲</small> <%=ratio%>%</div>
											<%else%>
												<div class="pt text-primary"><%=mid(formatcurrency(objrec("nowprice")),2,10)%></div>
												<div class="per text-primary"><small>▼</small> <%=ratio%>%</div>
											<%End if%>
										</div>
									</a>
								</li>
								<%objrec.movenext
								loop%>
							</ul>
						</div>
					<%objrec.close
					objconnCode.close
		'			Set objconnCode=nothing
					next%>
					</div>
				</div>
			</div>
			<script>
			var hotnoteswiper = new Swiper('.stock-hot-note .swiper-container', {
				slidesPerView: 'auto',
				centeredSlides: true,
				spaceBetween: 22,
				grabCursor: true
			});	
			var hotnoteswiper2 = new Swiper('.stock-comps .swiper-container');
			hotnoteswiper.on('slideChange', function (swiper) {
				hotnoteswiper2.slideTo(swiper.activeIndex)
			});
			hotnoteswiper2.on('slideChange', function (swiper) {
				hotnoteswiper.slideTo(swiper.activeIndex)
			});					
			</script>
		</div><!-- end panel -->	
		<div class="panel bg">
			<div class="heading">
				<h2>5%매매 현황</h2>				
			</div>
			<div class="trading-status">
				<div class="tab-wrapper">
					<div class="nav">
					<%iyear=Mid(iday,1,4)
					iyearcnt=iyear-2016%>
						<ul class="tabs">
							<%For i=1 To iyearcnt
								tabstr="tab1-" & i%>
								<li <%If i=1 then%>class="active" <%End if%>><a href="#<%=tabstr%>"><%=iyear-i+1%>년</a></li>
							<%next%>
						</ul>
					</div>		
					<%nowyear=iyear+1
					For i=1 To iyearcnt
						tabstr="tab1-" & i
						nowyear=nowyear-1

						strsql="Select sum(profit)*0.2 from stockgroup0 Where endday1>='"  & nowyear & "-01-01' and endday1<='" & nowyear & "-12-31'"
				'		Response.write strsql & "<br>"
						Set Rs=objconnNeo.Execute(strsql)
						profit=Rs(0)
						
						If profit<=0 Then 
							bar=0
							height=0
						else
							bar=CInt(profit/400*100)
							height=100
						End If
						If bar<20 Then
							height=bar
						End If %>						
					<div class="tab-content" id="<%=tabstr%>" <%If i=1 Then%> style="display:block;" <%else%> style="display:none;" <%End if%>>
						<div class="trading-status-wrap">
							<div class="yield">
								<div class="tt">수익률</div>
								<div class="stick"><span class="stick-bar" style="height:<%=height%>%"><span class="stick-in" style="height:<%=bar%>%"></span></span></div>
								<%If profit>=1000 Then 
									profit=FormatNumber(int(profit*100)/100,0)
								ElseIf profit>=100 And profit<1000 Then 
									profit=FormatNumber(int(profit*100)/100,1)
								Else
									profit=FormatNumber(int(profit*100)/100,2)
								end If 
								If profit>=0 then%>
									<div class="per"><%=profit%>%</div>
								<%Else
									profit=Abs(profit)%>
									<div class="per"><font color=#2d9cdb><%=profit%>%</font></div>
								<%End if%>
							</div>
							<div class="list">
								<%strsql="select top 3 * from stockgroup0 Where endday1>='"  & nowyear & "-01-01' and endday1<='" & nowyear & "-12-31' and profit>5  order by sn desc"
									objRec.Open strsql, objconnNeo%>									
								<ul>
								<%Do Until objrec.EOF%>
									<li>
										<div class="date"><%=Mid(objrec("endday1"),3,8)%></div>
										<div class="box">
											<div class="name"><%=objrec("codeName")%></div>
											<div class="won"><img src="../images/won.png" alt="won"> <%=objrec("loscutPrice")%></div>
											<div class="per"><%=objrec("profit")%>%</div>
										</div>
									</li>
								<%objrec.movenext
								Loop
								objrec.close%>
								</ul>
							</div>
						</div>
					</div>
					<%next%>
				</div>
			</div>
		</div><!-- end panel -->
	</div><!-- end content -->
</div><!-- end wrapper -->

<!-- modal -->
<div id="info" class="modal modal-sm modal-dark">
	<div class="modal-header">
		<h2>주식비중</h2>
	</div>
	<div class="modal-content">
		<p>최근 장의 흐름을 분석하여 30분 단위로<br>
		주식 비중을 표시합니다.<br>
		매수, 매도변곡이 나올때가 통상<br>
		바닥또는 꼭지일 경우가 많습니다.<br>
		약 한달간 주식비중을 확인 할 수 있습니다.<br>
		또한 통상적으로 오전보다 오후가 비중이 높게 나타납니다.</p>
	</div>
</div>
</body>
<%dbCount_Save objconn,"유진홈"


objconn.close
objconnNeo.close

Set objconnNeo=nothing
Set objRec = nothing
Set objCount = nothing
Set objRecCode = nothing
Set objconnCode = nothing%>
</html>

