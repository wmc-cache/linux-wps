/*
 *Created by chenyuxin
 */
var office = null;

function verify(obj) {
    console.info(obj);
}
/*2019年5月改了初始化方式*/
function Init(tagID) {
    if (office != undefined)
        office.Application.Quit();
    var iframe;
    iframe = document.getElementById(tagID);
    var codes = [];
    codes.push('<object name="rpcet" id="rpcet_id" type="application/x-et" wpsshieldbutton="false" width="100%" height="100%">');
    codes.push('<param name="quality" value="high" />');
    codes.push('<param name="bgcolor" value="#ffffff" />');
    codes.push('<param name="Enabled" value="1" />');
    codes.push('<param name="allowFullScreen" value="true" />');
    codes.push('</object>');
    iframe.innerHTML = codes.join("");
    office = document.getElementById("rpcet_id");
    window.onbeforeunload = function () {
        office.Application.Quit();
    };
	return office.Application;
}

// 初始化插件及启动ET
function InitEt() {
    var app = null;
    app = Init("et");
    if (!app) {
        // 为了兼容
        office.setAttribute('data', './newfile.et');
        var Interval_control = setInterval(
            function () {
                if (app) {
                    clearInterval(Interval_control);
                    AddFile();
                }
            }, 500);
    }
     else {
		setTimeout(function (){
			AddFile();
		},200);
    //RegisterBeforeCloseEvent();
	}
}
var _InitEt = new CreateFunction("初始化ET", InitEt, []);

function GetName()
{
	var app = office.Application;
	verify(app);
	alert(app.Name);
}

var _GetName = new CreateFunction("获取APP名", GetName, []);

function GetUserName()
{
	var app = office.Application;
	verify(app);
	alert(app.UserName);
}

var _GetUserName = new CreateFunction("获取用户名", GetUserName, []);

function AddFile()
{
	var app = office.Application;
	var workbook = app.Workbooks.Add();
	verify(workbook);
	alert(workbook.Name);
}

var _AddFile = new CreateFunction("新建文档", AddFile, []);


function wpsWorkbookBeforeCloseCallback()
{
	alert("wpsWorkbookBeforeCloseCallback");
}

function RegisterBeforeCloseEvent()
{
	var app = office.Application; 
	var ret = app.registerEvent("DIID_AppEvents","WorkbookBeforeClose","wpsWorkbookBeforeCloseCallback"); 
	var ret = app.registerEvent("DIID_AppEvents","WorkbookBeforeClose","wpsWorkbookBeforeCloseCallback"); 
	alert("WorkbookBeforeClose--"+ret); 
}
var _RegisterBeforeCloseEvent = new CreateFunction("注册关闭事件", RegisterBeforeCloseEvent, []);

function WorkbookBeforeSaveCallBack(Doc, SaveAsUI, Cancel) {
	if (SaveAsUI) {
		alert("另存");
		// 代表当前会弹出保存窗口，即另存为操作。
	} else {
		alert("保存");
		// 代表当前不会弹出保存窗口，即保存操作。
	}

	// 如果该值设置为true，则表示取消当前保存操作，无论如何不会弹出另存为对话框。
	//Cancel.SetValue(true);
	//console.log("Ctrl+S快捷键监控");
	// 如果该值设置为false，则表示正常保存，不影响对话框的显示与否。
     Cancel.SetValue(true);
	 alert("保存监控");
}

function RegisterWorkbookBeforeBrforeSaveEvent()
{
	var app = office.Application;
	var ret = app.registerEvent("DIID_AppEvents", "WorkbookBeforeSave", "WorkbookBeforeSaveCallBack");
	alert("监控 " + ret);		
}
var _RegisterWorkbookBeforeBrforeSaveEvent = new CreateFunction("注册保存事件", RegisterWorkbookBeforeBrforeSaveEvent, []);


function unRegisterDocumentBrforeSaveEvent() 
{
	var app = office.Application;
	var ret = app.unRegisterEvent("DIID_AppEvents", "WorkbookBeforeSave", "WorkbookBeforeSaveCallBack");
	alert("取消监控 " + ret);		
}
var _unRegisterDocumentBrforeSaveEvent = new CreateFunction("取消注册保存事件", unRegisterDocumentBrforeSaveEvent, []);


var office_ctrl;
function AddCommandBarButton()
{
	var app = office.Application;
	var cmdbar = app.CommandBars.Add("ETHAN");
	verify(cmdbar);
	alert(cmdbar.Name);
	var ctrl = cmdbar.Controls.Add(1);
	verify(ctrl);
	ctrl.Caption = "TEST";
	office_ctrl = ctrl;
}
var _AddCommandBarButton = new CreateFunction("CommandBar测试", AddCommandBarButton, []);

function InsertPic()
{
	var app = office.Application;
	var workbook = app.ActiveWorkbook;
	var sheet = workbook.ActiveSheet;
	sheet.Shapes.AddPicture("/home/wpsgch/桌面/1.png", 10,10,20,20,150,50);
}
var _InsertPic = new CreateFunction("插入图片", InsertPic, []);

function shapesitem()
{
	var app = office.Application;
	var workbook = app.ActiveWorkbook;
	var sheet = workbook.ActiveSheet;
    alert (sheet);
	var aa = sheet.Shapes.Item(0);
   //var bb =  aa.AddShape(msoShapeRectangle,50,50.100,200);
    alert (aa);
}
var _shapesitem = new CreateFunction("设置形状属性", shapesitem, []);

function Copy()
{
	var app = office.Application;
	var selection = app.Selection;
	selection.Copy();
	//var workbook = app.ActiveWorkbook;
	//var sheet = workbook.ActiveSheet;
	//var pagesetup = sheet.PageSetup;
	//alert(pagesetup.PrintArea);
	//alert(pagesetup.PrintArea = "A1:I93");
}
var _Copy = new CreateFunction("复制", Copy, []);


//20190711添加的接口。setForceBackUpEnabled()同时对二次开发和桌面程序生效。

	var workbook;
function OpenDocs_multiple()
	{
		var app = office.Application; 
		workbook = app.ActiveWorkbook; 
		app.ActiveWindow.Visible = false; 

		var count = app.Workbooks.Count;
		app.Workbooks.WorkbooksEx.Open('/home/wpsgch/桌面/xls.xls', false); 
		app.Workbooks.get_Item(count+1).Windows.get_Item(1).Visible = true;
		app.Workbooks.get_Item(count+1).Windows.get_Item(1).Activate();
	}
var _OpenDocs_multiple = new CreateFunction("打开并实例化多个文档", OpenDocs_multiple, []);

	
function OpenDocs_multiple_switch()
{
		var app = office.Application;
    	var value =app.Workbooks.get_Item(2).Windows.get_Item(1).Visible
		if (value)
	{
		workbook.Windows.get_Item(1).Visible = true;
		app.Workbooks.get_Item(2).Windows.get_Item(1).Visible = false;
	}
		else
	{
		app.Workbooks.get_Item(2).Windows.get_Item(1).Visible = true;
		workbook.Windows.get_Item(1).Visible = false;
	}
}
var _OpenDocs_multiple_switch = new CreateFunction("切换文档", OpenDocs_multiple_switch, []);
	
	
function setForceBackUpDisabled()
	{
		var app = office.Application;	
		alert(app.setForceBackUpEnabled(false));	//强制关闭实时备份功能
	}
var _setForceBackUpDisabled = new CreateFunction("关闭实时备份", setForceBackUpDisabled, []);	

function setForceBackUpEnabled()
	{
		var app = office.Application;
		alert(app.setForceBackUpEnabled(true));	//强制开启实时备份功能

	}
var _setForceBackUpEnabled = new CreateFunction("开启实时备份", setForceBackUpEnabled, []);

function DisableCtrlc()
	{
		alert(office.Application.enableCopy(false));				//禁用复制

	}
var _DisableCtrlc = new CreateFunction("禁用复制", DisableCtrlc, []);

function EnableCtrlc()
	{
		alert(office.Application.enableCopy(true));				//启用复制

	}
var _EnableCtrlc = new CreateFunction("启用复制", EnableCtrlc, []);

function DisableCtrlx()
	{

		alert(office.Application.enableCut(false));				//禁用剪切
	}
var _DisableCtrlx = new CreateFunction("禁用剪切", DisableCtrlx, []);
	
	
	
function EnableCtrlx()
	{
		alert(office.Application.enableCut(true));				//启用剪切

	}
var _EnableCtrlx = new CreateFunction("启用剪切", EnableCtrlx, []);



function setToolbarAllVisibleT()
{
	var aa = office.Application.setToolbarAllVisible(true);
    alert (aa);
}
var _setToolbarAllVisibleT = new CreateFunction("显示工具条", setToolbarAllVisibleT, []);



function setToolbarAllVisibleF()
{
	var aa = office.Application.setToolbarAllVisible(false);
    alert (aa);
}
var _setToolbarAllVisibleF = new CreateFunction("隐藏工具条", setToolbarAllVisibleF, []);


function FullScreen()
{
	var app = office.Application;
	app.FullScreen();
}
var _FullScreen = new CreateFunction("全屏", FullScreen, []);

function SaveFile()
{
	var filename = prompt("输入后缀", "/home/wpsgch/桌面/test.xlsx");
	alert(office.Application.saveAs(filename));
}
var _SaveFile = new CreateFunction("保存本地", SaveFile, []);

function SaveFile2()
{
	alert(office.Application.saveAs(""));
}
var _SaveFile2 = new CreateFunction("保存本地弹框", SaveFile2, []);



function SaveFile3()
{
    var name = prompt("输入后缀", "bbbb.xls")
	alert(office.Application.saveAs("",name));
}
var _SaveFile3 = new CreateFunction("保存本地弹框默认名字", SaveFile3, []);





function openDocument(flag)
{
	var name = prompt("输入地址", "/home/wpsgch/桌面/test.xlsx");
	alert(office.Application.openDocument(name, flag));

}
var _openDocument_readOnly = new CreateFunction("只读打开本地 \n参数为 true", openDocument, [true]);
var _openDocument_readWrite = new CreateFunction("可编辑打开本地 \n参数为 false", openDocument, [false]);

function openDocument2(flag)
{
	var name = prompt("输入地址", "/home/wpsgch/桌面/test.xlsx");
	alert(office.Application.openDocument(name, flag,"1"));

}
var _openDocument2_readOnly = new CreateFunction("只读打开本地_加密文档 \n参数为 true", openDocument2, [true]);
var _openDocument2_readWrite = new CreateFunction("可编辑打开本地_加密文档 \n参数为 false", openDocument2, [false]);

function Close()
{
  var a = office.Application.Workbooks.Close('true');
  alert(a);
}
var _Close = new CreateFunction("保存并关闭文档", Close, []);

function saveAsBase64Str()
{
	base64str_wps= office.Application.saveAsBase64Str("xlsx");
    alert(base64str_wps);
}
var _saveAsBase64Str = new CreateFunction("保存到base64", saveAsBase64Str, []);

function openDocumentFromBase64Str()
{
		if(typeof base64str_wps != "undefined" && base64str_wps != "")
		{
			var aa = office.Application.openDocumentFromBase64Str(base64str_wps, false);
			alert (aa);
               }
	       else{
               		alert("请先调用saveasbase64str");
               }
} 
var _openDocumentFromBase64Str = new CreateFunction("从base64打开", openDocumentFromBase64Str, []);

function testApplicationExopen()
{
	var app = office.Application;
	var appEx = app.ApplicationEx;

	appEx.put_EmbedTrueTypeFonts(true);
	alert("是否嵌入字体:"+appEx.EmbedTrueTypeFonts);
} 
var _testApplicationExopen = new CreateFunction("嵌入字体", testApplicationExopen, []);

function testApplicationExopenclose()
{
	var app = office.Application;
	var appEx = app.ApplicationEx;
	appEx.put_EmbedTrueTypeFonts(false);
	alert("是否嵌入字体:"+appEx.EmbedTrueTypeFonts);
} 
var _testApplicationExopenclose = new CreateFunction("不嵌入字体", testApplicationExopenclose, []);

function AddRow()
{
	var app = office.Application;
	var selection = app.ActiveCell.EntireRow;
	selection.Insert();
}
var _AddRow = new CreateFunction("添加行", AddRow, []);

function deleteRow()
{
	var app = office.Application;
	var selection = app.ActiveCell.EntireRow;
	selection.Delete();
}
var _deleteRow = new CreateFunction("删除行", deleteRow, []);

function Range_Value()
{
	office.Application.Workbooks.Add();
	var sheet1 = office.Application.ActiveSheet;	
	sheet1.get_Range("A1").put_Value(10,"kingsoft");
	sheet1.get_Range("A2").put_Value2("value2");
	sheet1.get_Range("A3").put_Formula("=sum(1)");
	sheet1.get_Range("B1:C1").put_FormulaArray("={1,2}");
	alert(sheet1.get_Range("A3").get_Value2());
	alert(sheet1.get_Range("A3").get_Text());
	alert(sheet1.get_Range("A3").get_Formula());
}
var _Range_Value = new CreateFunction("单元格设值", Range_Value, []);


function Range_Merge()
{
	var rg = office.Application.ActiveSheet.get_Range("A1:B4");
	rg.Merge();
}
var _Range_Merge = new CreateFunction("合并单元格", Range_Merge, []);

function Range_unMerge()
{
	var rg = office.Application.ActiveSheet.get_Range("A1:B4");
	rg.UnMerge();
}
var _Range_unMerge = new CreateFunction("取消合并单元格", Range_unMerge, []);

function setRangeFormat()
{
	office.Application.ActiveSheet.get_Range("A1:A3").put_Value2(1);
	office.Application.ActiveSheet.get_Range("A1").put_NumberFormatLocal("0.00_ ");
	office.Application.ActiveSheet.get_Range("A2").put_NumberFormatLocal("0.00%");
	office.Application.ActiveSheet.get_Range("A3").put_NumberFormatLocal("￥#,##0.00;￥-#,##0.00");
	office.Application.ActiveSheet.get_Range("B1").put_Locked(0);

	alert(office.Application.ActiveSheet.get_Range("A1").get_NumberFormatLocal());
	alert(office.Application.ActiveSheet.get_Range("B1").get_Locked());
}
var _setRangeFormat = new CreateFunction("设置单元格格式", setRangeFormat, []);

function Range_Row_Col()
{
	office.Application.Workbooks.Add();
	var rg = office.Application.Selection;
	rg.put_RowHeight(20);
	rg.put_ColumnWidth(20);
	alert("行高");
	alert(rg.RowHeight);
	alert("列宽");
	alert(rg.ColumnWidth);
}
var _Range_Row_Col = new CreateFunction("设置单元格行高列宽", Range_Row_Col, []);


function GetTempPath()
	{
		var app = office.Application;
		var aa = app.GetTempPath();
		alert (aa);
	}
var _GetTempPath = new CreateFunction("获取临时文件夹路径", GetTempPath, []);


//20190716新增
function setTmpFilepath()
	{
		alert(office.setTmpFilepath("/home/wpsgch/桌面"));
	}
var _setTmpFilepath = new CreateFunction("设置临时文件路径", setTmpFilepath, []);	



//20190724新增接口
function RegisterPrintOutPageSetEvent() 
{ 
	var appex = office.Application.ApplicationEx; 
	var ret = appex.registerEvent("DIID_ApplicationEventsEx","DocumentAfterPrint","EventCallBackPrintOutPageSet"); 
	alert("PrintOutPageSet--"+ret); 
}
var _RegisterPrintOutPageSetEvent = new CreateFunction("注册输出打印区域事件", RegisterPrintOutPageSetEvent, []);	

function EventCallBackPrintOutPageSet(workbook, pageset)
{
	var range = pageset.get_PrintOutRange();
	if (range == 0){
		alert("选定区域");
	}else if (range == 1){
		alert("选定工作表");
	}else if (range == 2){
		alert("选定工作簿");
	}
	alert("页码范围");
	alert(pageset.get_PrintOutPages());
}

function exitEditMode()
{
	var aa = office.Application.exitEditMode();
	alert(aa);
}
var _exitEditMode = new CreateFunction("419086_退出编辑模式", exitEditMode, []);


function protectAllSheet() {
	var workbook = office.Application.ActiveWorkbook;
	workbook.Protect('123');
	var sheets = office.Application.ActiveWorkbook.Sheets;
	var count = sheets.Count;
	for (var i = 1; i <= count; ++i) {
			sheets.get_Item(i).Protect('123');
		}
	}
var _protectAllSheet = new CreateFunction("419086_保护所有工作表", protectAllSheet, []);

function unProtectAllSheet() {
	var workbook = office.Application.ActiveWorkbook;
	workbook.Unprotect('123');
	var sheets = office.Application.ActiveWorkbook.Sheets;
	var count = sheets.Count;
	for (var i = 1; i <= count; ++i) {
			sheets.get_Item(i).Unprotect('123');
		}
	}
var _unProtectAllSheet = new CreateFunction("419086_解除保护所有工作表", unProtectAllSheet, []);


/*保存远程文档不落地-旧接口*/
/*测试旧接口的时候，需要将servlet服务删掉，并且重启Tomcat服务*/
function saveURL_s()
{
	var aa = office.Application.saveURL_s("http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/upload_l.jsp", "BUG改了.xlsx" );
        alert (aa);
}
var _saveURL_s = new CreateFunction("保存到远程不落地_保存当前打开的文档", saveURL_s, []);
/*打开远程文档不落地-旧接口*/
/*测试旧接口的时候，需要将servlet服务删掉，并且重启Tomcat服务*/
function openDocumentRemote_s()
{
	var app = office.Application;	
	var aa = app.openDocumentRemote_s("http://10.90.128.210:8080/servletTestnew/BUG改了.xlsx", false);
	alert (aa);
}
var _openDocumentRemote_s = new CreateFunction("打开远程文档不落地", openDocumentRemote_s, []);





/*保存当前打开的文档到服务器*/
function SaveUrl()
{
    var aa = office.Application.saveURL('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/upload_l.jsp', '300MB.xlsx');
	alert (aa);
}
var _SaveUrl = new CreateFunction("保存远程", SaveUrl, []);


/*打开远程文档-旧接口*/
/*测试旧接口的时候，需要将servlet服务删掉，并且重启Tomcat服务*/
function openDocumentRemote(flag)
{
	var name = prompt("输入地址", "aa.xlsx");
	alert(office.Application.openDocumentRemote('http://10.90.128.210:8080/servletTestnew/'+ name, flag));

}
var _openDocumentRemote_readOnly = new CreateFunction("只读打开远程 \n参数为 true", openDocumentRemote, [true]);
var _openDocumentRemote_readWrite = new CreateFunction("可编辑打开远程 \n参数为 false", openDocumentRemote, [false]);

function OpenRemote()
{
	
	alert(office.Application.Workbooks.OpenRemote('http://10.90.128.210:8080/servletTestnew/XLS_100MB.xls'));

}
var _OpenRemote = new CreateFunction("打开远程_oldinterface", OpenRemote, []);

function openDocumentRemote1(flag)
{
	var name = prompt("输入地址", "1.xlsx");
	alert(office.Application.openDocumentRemote('http://10.90.128.210:8080/servletTestnew/XLS_100MB.xls'+ name, flag,"1"));

}
var _openDocumentRemote1_readOnly = new CreateFunction("只读打开远程_加密文档 \n参数为 true", openDocumentRemote1, [true]);
var _openDocumentRemote1_readWrite = new CreateFunction("可编辑打开远程_加密文档 \n参数为 false", openDocumentRemote1, [false]);

//https协议打开远程
function openDocumentRemote_https(flag)
{
	var name = prompt("输入地址", "加密.xlsx");
	alert(office.Application.openDocumentRemote('https://10.90.128.230:8443/servletTestnew/X862019armhttpssession.xlsx'+ name, flag,"1"));

} 
var _openDocumentRemote_readOnly_https = new CreateFunction("只读打开远程_https \n参数为 true", openDocumentRemote_https, [true]);
var _openDocumentRemote_readWrite_https = new CreateFunction("可编辑打开远程_https \n参数为 false", openDocumentRemote_https, [false]);



/*上传远程文档到服务器-旧接口*/
/*测试旧接口的时候，需要将servlet服务删掉，并且重启Tomcat服务*/
function SendDataToServer()
{
  var a = office.Application.Workbooks.SendDataToServer('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/upload_l.jsp','/home/wpsgch/桌面/100MB.xls','test上传arm.xlsx');
  alert(a);
}
var _SendDataToServer = new CreateFunction("上传到远程", SendDataToServer, []);


/*下载远程文档到本地-旧接口*/
/*测试旧接口的时候，需要将servlet服务删掉，并且重启Tomcat服务*/
function DownLoadServerFile()
{
  var a = office.Application.Workbooks.DownLoadServerFile('http://10.90.128.210:8080/servletTestnew/XLS_100MB.xls','/home/wpsgch/桌面/XLS_100MB.xls');
  alert(a);
}
var _DownLoadServerFile = new CreateFunction("下载远程文档", DownLoadServerFile, []);

//2019年11月13日saveurl接口增加自定义数据上传的参数
/*保存到远程_上传自定义数据*/
function saveURL_CustomParam()
{
	var jsondata = {key1:"aa"}; 
	var aa = office.Application.saveURL_CustomParam("http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/upload_l.jsp", "测试自定义30M.xls", JSON.stringify(jsondata)); 
	alert(aa);
}
var _saveURL_CustomParam = new CreateFunction("保存到远程_上传自定义数据", saveURL_CustomParam, []);

//https协议保存到远程
function SaveUrl_https()
{
	var name = prompt("输入地址", "1.xlsx");
    var aa = office.Application.saveURL('https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/upload_l.jsp', name);
	alert (aa);
}
var _saveURL_https = new CreateFunction("保存到远程_https", SaveUrl_https, []);
	

//2019年12月19号新增四个接口UploadFileToServer、DownloadFileFromServer、SaveDocumentToServer、OpenDocumentFromServer
//50M用时4s,100M用时6s(7863版本)
function UploadFileToServer()
{
	var jsondata = {
	fileName:"111111111中文2.xlsx", //中文名不允许？
	isGetBodyResult:true,	//是否需要返回请求数据
	isGetResponseHead:true,		//是否返回响应头信息
	bDelLocalFile:false,				//是否删除本地
	customFormData:{
		"key1":"value1上传远程_新11111111111111",
		"key2":"value2上传远程_新11111111111"
		},
		customHeadData:{
	"Cookie":"value11111111111111111",
	"key1":"value2222上传远程_新2222222222",
	"key2":"abc"
	}
		};
	var aa = office.UploadFileToServer("http://10.90.128.210:8080/servletTestnew/HelloServlet", "/home/wpsgch/桌面/特殊.xlsx", JSON.stringify(jsondata));
	alert(aa);
}
var _UploadFileToServer = new CreateFunction("上传远程_新", UploadFileToServer, []);

//50M用时2s,100M用时3s(7863版本)
function DownloadFileFromServer()
{
	var jsondata = {
	isGetResponseHead:true,
	customHeadData:{
	"Cookie":"value1111下载远程_新",
	"Cookie":"value2222下载远程_新"
	}
	};
	var aa = office.DownloadFileFromServer("http://10.90.128.210:8080/servletTestnew/arm飞腾_特殊222.xlsx", "/home/wpsgch/桌面/arm飞腾_特殊.xlsx",JSON.stringify(jsondata));
	alert(aa);
}
var _DownloadFileFromServer = new CreateFunction("下载远程_新", DownloadFileFromServer, []);

//50M用时20s，100M用时40s
//OFD落地、不落地；PDF落地；xls落地、xls不落地；xlsx落地、xlsx不落地；遍历true/false/文件名；性能
function SaveDocumentToServer()
{
	var jsondata = {
		fileName:"龙芯loongson_删落地.xls",    //中文名不允许？
		isGetBodyResult:true,		//是否需要返回请求数据
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,				//是否不落地保存
		customFormData:{
			"key1":"value1gson_保存远程_新111111111",
			"key2":"value2gson_保存远程_新1211111111111"
},
		customHeadData:{
		"key1":"value1111保存远程_新21222222222222",
		"key2":"value2保存远程_新2222222222"
		}
};
	var aa = office.Application.SaveDocumentToServer("http://10.90.128.210:8080/servletTestnew/HelloServlet",  JSON.stringify(jsondata));
	alert(aa);
	
	}	 
	var _SaveDocumentToServer = new CreateFunction("保存远程_新", SaveDocumentToServer, []);

//50M用时21s，100M用时48s
//OFD、PDF不落地；xls落；xls不落地 ;xlsx落地；xlsx不落地-失败       
function OpenDocumentFromServer()
{
	var jsondata = {
		password:"123",					//打开文档所需的密码
		readOnly:true,					//是否只读打开文档
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,					//是否不落地保存
		customHeadData:{
		"key1":"value1111打开远程_新",
		"key2":"value2222打开远程_新"
	}				
		};
	var aa = office.Application.OpenDocumentFromServer("http://10.90.128.210:8080/servletTestnew/龙芯loongson_删落地.xls",  JSON.stringify(jsondata));
	alert(aa);
		
	}
	var _OpenDocumentFromServer = new CreateFunction("打开远程_新", OpenDocumentFromServer, []);

//二进制

/*保存远程文档不落地-带session*/
/***************************所有带session的新接口都必须在服务端部署servlet服务**************************************/
function saveURL_s_FormData()
{
	var name = prompt("输入名字", "2019最终不落地测试httpsarm.xls");
	var aa = office.Application.saveURL_s_FormData("http://10.90.128.210:8080/servletTestnew/HelloServlet", name );
    alert (aa);
}
var _saveURL_s_FormData = new CreateFunction("保存到远程不落地_保存当前打开的文档（携带session）", saveURL_s_FormData, []);


function saveURL_s_FormData_https()
{
	var name = prompt("输入名字", "X862019_buluodiarm.ofd");
	var aa = office.Application.saveURL_s_FormData("https://10.90.128.210:8443/servletTestnew/HelloServlet", name );
    alert (aa);
}
var _saveURL_s_FormData_https = new CreateFunction("保存远程不落地_当前打开_HTTPS带session", saveURL_s_FormData_https, []);



function saveURL_s_FormData_rediect()
{
	var name = prompt("输入名字", "找着.xlsx");
	var aa = office.Application.saveURL_s_FormData("http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet.jsp", name );
    alert (aa);
}
var _saveURL_s_FormData_rediect = new CreateFunction("保存远程不落地_当前打开_重定向带session", saveURL_s_FormData_rediect, []);


function saveURL_s_FormData_rediect_https()
{
	var name = prompt("输入名字", "X862019_buluodiarm.xlsx");
	var aa = office.Application.saveURL_s_FormData("https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet_https.jsp", name );
    alert (aa);
}
var _saveURL_s_FormData_rediect_https = new CreateFunction("保存远程不落地_当前打开_重定向HTTPS带session", saveURL_s_FormData_rediect_https, []);


/*打开远程文档不落地带session*/
function openDocumentRemote_s_FormData()
{
	var app = office.Application;	
	var name = prompt("输入名字", "XLS_100MB.xls");
	var aa = app.openDocumentRemote_s("http://10.90.128.210:8080/servletTestnew/"+name, false);
	alert (aa);
}
var _openDocumentRemote_s_FormData = new CreateFunction("打开远程文档不落地带session", openDocumentRemote_s_FormData, []);



function openDocumentRemote_s_FormData_https()
{
	var app = office.Application;	
	var name = prompt("输入名字", "XLS_100MB.xls");
	var aa = app.openDocumentRemote_s("https://10.90.128.210:8443/servletTestnew/"+name, false);
	alert (aa);
}
var _openDocumentRemote_s_FormData_https = new CreateFunction("打开远程文档不落地_HTTPS带session", openDocumentRemote_s_FormData_https, []);

function openDocumentRemote_s_rediect()
{
	var app = office.Application;	
	var aa = app.openDocumentRemote_s("http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_down.jsp", false);
	alert (aa);
}
var _openDocumentRemote_s_rediect = new CreateFunction("打开远程文档不落地_重定向带session", openDocumentRemote_s_rediect, []);

function openDocumentRemote_s_rediect_https()
{
	var app = office.Application;	
	var aa = app.openDocumentRemote_s("https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_down_https.jsp", false);
	alert (aa);
}
var _openDocumentRemote_s_rediect_https = new CreateFunction("打开远程文档不落地_重定向HTTPS", openDocumentRemote_s_rediect_https, []);

/*保存当前打开的文档到服务器*/
function SaveUrl_FormData()
{
	var name = prompt("输入名字", "saveURL带session2019arm.xls");
    var aa = office.Application.saveURL_FormData('http://10.90.128.210:8080/servletTestnew/HelloServlet', name);
	alert (aa);
}
var _SaveUrl_FormData = new CreateFunction("保存远程_带session", SaveUrl_FormData, []);

function SaveUrl_FormData_rediect()
{
	var name = prompt("输入名字", "X862019arm.xls");
    var aa = office.Application.saveURL_FormData('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet.jsp', name);
	alert (aa);
}
var _SaveUrl_FormData_rediect = new CreateFunction("保存远程_重定向带session", SaveUrl_FormData_rediect, []);

function saveURL_FormData_https()
{
	var name = prompt("输入名字", "X862019arm.xls");
    var aa = office.Application.saveURL_FormData('https://10.90.128.210:8443/servletTestnew/HelloServlet', name);
	alert (aa);
}
var _saveURL_FormData_https = new CreateFunction("保存远程_HTTPS带session", saveURL_FormData_https, []);

function saveURL_FormData_rediect_https()
{
	var name = prompt("输入名字", "X862019_buluodiarm.xlsx");
	var aa = office.Application.saveURL_FormData("https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet_https.jsp", name);
        alert (aa);
}
var _saveURL_FormData_rediect_https = new CreateFunction("保存远程_重定向HTTPS带session", saveURL_FormData_rediect_https, []);

/*打开远程文档-带session*/
function openDocumentRemote_FormData(flag)
{
	var name = prompt("输入地址", "X862019arm.xls");
	alert(office.Application.openDocumentRemote('http://10.90.128.210:8080/servletTestnew/'+ name, flag));

}
var _openDocumentRemote_FormData = new CreateFunction("打开远程文档带session \n参数为 false", openDocumentRemote_FormData, [false]);

function openDocumentRemote_FormData_https(flag)
{
	var name = prompt("输入地址", "X862019arm.xls");
	alert(office.Application.openDocumentRemote('https://10.90.128.210:8443/servletTestnew/'+ name, flag));

}
var _openDocumentRemote_FormData_https = new CreateFunction("打开远程文档_HTTPS带session \n参数为 false", openDocumentRemote_FormData_https, [false]);

function openDocumentRemote_rediect(flag)
{
	alert(office.Application.openDocumentRemote('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_down.jsp',name, flag));

}
var _openDocumentRemote_rediect = new CreateFunction("打开远程文档_重定向带session \n参数为 false", openDocumentRemote_rediect, [false]);


function openDocumentRemote_rediect_https(flag)
{
	alert(office.Application.openDocumentRemote('https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_down_https.jsp',name, flag));

}
var _openDocumentRemote_rediect_https = new CreateFunction("打开远程文档_重定向HTTPS \n参数为 false", openDocumentRemote_rediect_https, [false]);

/*上传远程文档到服务器-带session*/
function SendDataToServer_FormData()
{
	var headData = {};
	var name = prompt("输入要上传的文件", "/home/wpsgch/桌面/test.xlsx");
	headData.filename = "1arm.xlsx";
	var aa = office.SendDataToServer_FormData("http://10.90.128.210:8080/servletTestnew/HelloServlet",name,JSON.stringify(headData),false);
	alert(aa);
	
}

var _SendDataToServer_FormData = new CreateFunction("上传到远程_带session", SendDataToServer_FormData, []);

function SendDataToServer_FormData_https()
{
  var name = prompt("输入下载本地地址", "/home/wpsgch/桌面/test.xlsx");
  var a = office.Application.Workbooks.SendDataToServer_FormData('https://10.90.128.210:8443/servletTestnew/HelloServlet',name,'FormData_https.xlsx');
  alert(a);
}
var _SendDataToServer_FormData_https = new CreateFunction("上传到远程_HTTPS带session", SendDataToServer_FormData_https, []);


function SendDataToServer_FormData_rediect()
{
  var name = prompt("输入下载本地地址", "/home/wpsgch/桌面/123.xls");
  var a = office.Application.Workbooks.SendDataToServer_FormData('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet.jsp',name,'FormData_rediect.xlsx');
  alert(a);
}
var _SendDataToServer_FormData_rediect = new CreateFunction("上传到远程_重定向带session", SendDataToServer_FormData_rediect, []);

function SendDataToServer_FormData_rediect_https()
{
  var name = prompt("输入下载本地地址", "/home/wpsgch/桌面/123.xls");
  var a = office.Application.Workbooks.SendDataToServer_FormData('https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet_https.jsp',name,'FormData_rediect_https.xlsx');
  alert(a);
}
var _SendDataToServer_FormData_rediect_https = new CreateFunction("上传到远程_重定向HTTPS带session", SendDataToServer_FormData_rediect_https, []);

/*下载远程文档到本地-带session*/
function DownLoadServerFile_FormData()
{
  var a = office.Application.Workbooks.DownLoadServerFile('http://10.90.128.210:8080/servletTestnew/XLS_100MB.xls','/home/wpsgch/桌面/XLS_100MB.xls');
  alert(a);
}
var _DownLoadServerFile_FormData = new CreateFunction("下载远程文档带session", DownLoadServerFile_FormData, []);

function DownLoadServerFile_FormData_https()
{
  var a = office.Application.Workbooks.DownLoadServerFile('https://10.90.128.210:8443/servletTestnew/123.xlsx','/home/wpsgch/桌面/1232019down.xlsx');
  alert(a);
}
var _DownLoadServerFile_FormData_https = new CreateFunction("下载远程文档_HTTPS带session", DownLoadServerFile_FormData_https, []);

function DownLoadServerFile_rediect()
{
  var a = office.Application.Workbooks.DownLoadServerFile('http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_down.jsp','/home/wpsgch/桌面/100MB.xls');
  alert(a);
}
var _DownLoadServerFile_rediect = new CreateFunction("下载远程文档_重定向带session", DownLoadServerFile_rediect, []);

//2019年11月13日saveurl接口增加自定义数据上传的参数
/*保存到远程_上传自定义数据_带session*/
function saveURL_CustomParam_FormData()
{
	var jsondata = {key1:"ccesddf"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("http://10.90.128.210:8080/servletTestnew/HelloServlet", "带session.ofd", JSON.stringify(jsondata)); 
	alert(aa);
}
var _saveURL_CustomParam_FormData = new CreateFunction("保存到远程_上传自定义数据_带session", saveURL_CustomParam_FormData, []);

/*保存到远程_上传自定义数据_重定向带session*/
function saveURL_CustomParam_FormData_rediect()
{
	var jsondata = {key1:"wwwwwwwwwww"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("http://10.90.128.210:8080/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet.jsp", "测试自定义rediect.xls", JSON.stringify(jsondata)); 
	alert(aa);
}

var _saveURL_CustomParam_FormData_rediect = new CreateFunction("保存到远程_上传自定义数据_重定向带session", saveURL_CustomParam_FormData_rediect, []);


function DownLoadServerFile_rediect_https()
{
  var a = office.Application.Workbooks.DownLoadServerFile('https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_down_https.jsp','/home/wpsgch/桌面/down_rediect_https.xlsx');
  alert(a);
}

var _DownLoadServerFile_rediect_https = new CreateFunction("下载远程文档_重定向HTTPS带session", DownLoadServerFile_rediect_https, []);

//2019年11月13日saveurl接口增加自定义数据上传的参数
/*HTTPS保存到远程_上传自定义数据_带session*/
function saveURL_CustomParam_FormData_https()
{
	var jsondata = {key1:"FormData_https"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("https://10.90.128.210:8443/servletTestnew/HelloServlet", "带sessionhttps.uos", JSON.stringify(jsondata)); 
	alert(aa);
}

var _saveURL_CustomParam_FormData_https = new CreateFunction("保存到远程_HTTPS上传自定义数据_带session", saveURL_CustomParam_FormData_https, []);

//2019年11月13日saveurl接口增加自定义数据上传的参数
/*HTTPS保存到远程_上传自定义数据_重定向带session*/
function saveURL_CustomParam_FormData_rediect_https()
{
	var jsondata = {key1:"FormData_rediect_https"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("https://10.90.128.210:8443/servletTestnew/wps_webdemo/linux/src/et/rediect_servlet_https.jsp", "测试自定义HTTPS_rediect.xls", JSON.stringify(jsondata)); 
	alert(aa);
}

var _saveURL_CustomParam_FormData_rediect_https = new CreateFunction("保存到远程_上传自定义数据_重定向带session", saveURL_CustomParam_FormData_rediect_https, []);


// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 函数调用区 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// 加载后执行
window.onload = function () {
    InitLayui();
}
