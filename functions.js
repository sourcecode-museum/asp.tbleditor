/*	==============================================================
	 TableEditoR 0.81 Beta
	 http://www.2enetworx.com/dev/projects/tableeditor.asp
	--------------------------------------------------------------
	 File: functions.js
	 Description: Javascript functions for dynamic operations
	 Initiated By Rami Kattan on Apr 18, 2002
	--------------------------------------------------------------
	 Copyright (c) 2002, 2eNetWorX/dev.
	
	 TableEditoR is distributed with General Public License.
	 Any derivatives of this software must remain OpenSource and
	 must be distributed at no charge.
	 (See license.txt for additional information)
	
	 See Credits.txt for the list of contributors.
	
	 Change Log:
	--------------------------------------------------------------
	' # Jun 3, 2002 by Rami Kattan
	' Fix javascript for Combo Tables
	==============================================================
*/

function GetObject(obj){
	isNS4 = (document.layers) ? true : false;
	isIE4 = (document.all && !document.getElementById) ? true : false;
	isIE5 = (document.all && document.getElementById) ? true : false;
	isNS6 = (!document.all && document.getElementById) ? true : false;

	if (isNS4){
	   elem = document.layers[obj];
	}
	else if (isIE4) {
	   elem = document.all[obj];
	}
	else if (isIE5 || isNS6) {
	   elem = document.getElementById(obj);
	}

	return elem;
}

function SetCurrentPage(page){
	GetObject("Current_Page").value = page;
}

function GetCurrentPage(){
	return GetObject("Current_Page").value;
}

function GetQueryString(){
	return 	GetObject("db_string").value;
}

function GetPerPage(){
	return GetObject("URLSelect").options[GetObject("URLSelect").selectedIndex].value;
}

function GetTotalPages(){
	if (GetPerPage() != '0')
		retValue = Math.ceil(GetTotalRecs()/GetPerPage());	
	else
		retValue = 1;
	if (retValue == 0) retValue = 1;
	return retValue;
}

function GetTotalRecs(){
	return GetObject("totalRecords").value;
}

function GetlConnID(){
	return GetObject("db_lConnID").value;
}

function GetsTableName(){
	return GetObject("db_sTableName").value;
}

function SetPageIndicator(Current_Page){
	if (GetObject("pg").type != undefined)
		GetObject("pg").value = Current_Page;
	else
		GetObject("pg").innerText = Current_Page;
}

function GetPageIndicator(){
	var Current_Page
	if (GetObject("pg").type != undefined)
		Current_Page = parseInt(GetObject("pg").value);
	else
		Current_Page = parseInt(GetObject("pg").innerText);

	if (Current_Page < 1){
		alert("Page number cannot be less than 1.")
		SetPageIndicator(1);
		Current_Page = 1;
	}

	return Current_Page;
}

function gotoNext(){
	page_num = parseInt(GetCurrentPage()) + 1;
	per_page = GetPerPage();
	totPages = GetTotalPages();
	OrderType = GetObject("db_Ordering").value;

	DynamicData.DataURL='te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + page_num + '&cPerPage=' + per_page + OrderType + GetQueryString();
	DynamicData.Reset();""

	SetCurrentPage(page_num);
	SetPageIndicator(page_num);

	GetObject("btnFirst").disabled = false;
	GetObject("btnPrev").disabled = false;
	if (page_num == totPages) {
		GetObject("btnNext").disabled = true;
		GetObject("btnLast").disabled = true;
	}
}

function gotoPrev(){
	page_num = parseInt(GetCurrentPage()) - 1;
	per_page = GetPerPage();
	totPages = GetTotalPages();
	OrderType = GetObject("db_Ordering").value;

	DynamicData.DataURL='te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + page_num + '&cPerPage=' + per_page + OrderType + GetQueryString();
	DynamicData.Reset();""

	SetCurrentPage(page_num);
	SetPageIndicator(page_num);

	GetObject("btnLast").disabled = false;
	GetObject("btnNext").disabled = false;
	if (page_num == 1) {
		GetObject("btnFirst").disabled = true;
		GetObject("btnPrev").disabled = true;
	}
}

function gotoFirst(){
	per_page = GetPerPage();
	totPages = GetTotalPages();
	page_num = 1;
	OrderType = GetObject("db_Ordering").value;

	DynamicData.DataURL='te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + page_num + '&cPerPage=' + per_page + OrderType + GetQueryString();
	DynamicData.Reset();""

	SetPageIndicator(1);
	SetCurrentPage(1);

	GetObject("btnFirst").disabled = true;
	GetObject("btnPrev").disabled = true;
	if (totPages == 1) {
		GetObject("btnNext").disabled = true;
		GetObject("btnLast").disabled = true;
	}	else	{
		GetObject("btnNext").disabled = false;
		GetObject("btnLast").disabled = false;
	}
}

function gotoLast(){
	per_page = GetPerPage();
	totPages = GetTotalPages();
	page_num = totPages;
	OrderType = GetObject("db_Ordering").value;

	DynamicData.DataURL='te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + page_num + '&cPerPage=' + per_page + OrderType + GetQueryString();
	DynamicData.Reset();""

	SetPageIndicator(page_num);
	SetCurrentPage(page_num);

	GetObject("btnNext").disabled = true;
	GetObject("btnLast").disabled = true;
	if (totPages != 1) {
		GetObject("btnFirst").disabled = false;
		GetObject("btnPrev").disabled = false;
	} else {
		GetObject("btnFirst").disabled = true;
		GetObject("btnPrev").disabled = true;
	}
}

function check_paging(){
	page_num = GetCurrentPage();
	per_page = GetPerPage();
	totPages = GetTotalPages();
	if (page_num != 1){
		GetObject("btnFirst").disabled = false;
	} else {
		GetObject("btnFirst").disabled = true;
	}

	if (page_num > 1){
		GetObject("btnPrev").disabled = false;
	} else {
		GetObject("btnPrev").disabled = true;
	}

	if (page_num < totPages){
		GetObject("btnNext").disabled = false;
	} else {
		GetObject("btnNext").disabled = true;
	}

	if (page_num != totPages){
		GetObject("btnLast").disabled = false;
	} else {
		GetObject("btnLast").disabled = true;
	}
}

function update_view(){
	SetCurrentPage(GetPageIndicator());
	OrderType = GetObject("db_Ordering").value;
	per_page = GetPerPage();
	if (per_page == '0')
	{
		per_page = GetTotalRecs();
	}

	if (parseInt(GetCurrentPage()) > parseInt(GetTotalPages()))
	{
		SetCurrentPage(GetTotalPages());
		SetPageIndicator(GetTotalPages());
	}
	page_num = GetCurrentPage();

	GetObject("pagecount").innerText = GetTotalPages();

	RecordCounter.DataURL='te_readDB_counter.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + GetQueryString();
	RecordCounter.Reset();

	DynamicData.DataURL='te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + page_num + '&cPerPage=' + per_page + OrderType + GetQueryString();
	DynamicData.Reset();
	
	var ops_html = "";
	if (GetObject("pg").type == "select-one" && GetObject("pg").options[GetObject("pg").options.length-1].value != parseInt(GetTotalPages()) )
	{
		for (m = GetObject("pg").options.length - 1; m > 0 ; m--)
			GetObject("pg").options[m]=null;

		for (i=0; i < GetTotalPages(); i++)
			GetObject("pg").options[i] = new Option(i+1,i+1);
		SetPageIndicator(GetCurrentPage());
	}
	check_paging();
}

function openWindow(url) {
  popupWin = window.open(url,'new_page','width=540,height=450,top=80,left=40,scrollbars=yes,resizable=yes')
}

function ShowTransientMessage() {
	window.status="Please wait While loading data...";
//	DisplayMessageBox.innerText=sMessage;
	DisplayMessageBox.style.display='';
	DisplayMessageBox.style.pixelTop=(document.body.clientHeight/2)-(DisplayMessageBox.offsetHeight/2)+(document.body.scrollTop);
	DisplayMessageBox.style.pixelLeft=(document.body.clientWidth/2)-(DisplayMessageBox.offsetWidth/2)+(document.body.scrollLeft);
}

function HideTransientMessage() {
	window.status="";
	DisplayMessageBox.style.display='none';
}

//Check all radio/check buttons script- by javascriptkit.com
//Visit JavaScript Kit (http://javascriptkit.com) for script
//Credit must stay intact for use  

function checkall(thestate){
	var el_collection = eval("document.forms.frmAddDelete.chkDel")
	for (c=0; c < el_collection.length; c++)
		el_collection[c].checked = thestate
}

function toggleFilter(){
	if (GetObject("filtering").style.display  == "inline")
		GetObject("filtering").style.display  = "none";
	else
		GetObject("filtering").style.display  = "inline";
}

function FilterData(){
	var FilterField = GetObject("FilterField").value;
	var FilterWord = GetObject("FilterWord").value;

	DynamicData.Filter = FilterField + "=*" + FilterWord + "*";
	DynamicData.Reset();

	toggleFilter();
}

function changeOrderingAsc(Sortfield){
	GetObject("db_Ordering").value = "&orderby=" + Sortfield + "&dir=asc";
	GetObject("excel_ordering").value = Sortfield;
	GetObject("excel_ordering_dir").value = "ASC";
	SetPageIndicator(1);
	update_view();
}
function changeOrderingDesc(Sortfield){
	GetObject("db_Ordering").value = "&orderby=" + Sortfield + "&dir=desc";
	GetObject("excel_ordering").value = Sortfield;
	GetObject("excel_ordering_dir").value = "DESC";
	SetPageIndicator(1);
	update_view();
}

function MainFormAction(action){
	GetObject("frmAddDelete").action = "te_" + action + ".asp" + GetObject("mainFrmExt").value;
	GetObject("frmAddDelete").submit();
}

function cboChangeDB() {
	location.href = "te_listtables.asp?cid=" + GetObject("allDBs").options[GetObject("allDBs").selectedIndex].value
}

function debug(){
	alert('te_readDB.asp?cid=' + GetlConnID() + '&tablename=' + GetsTableName() + '&ipage=' + GetCurrentPage() + '&cPerPage=' + GetPerPage() + GetQueryString());
}