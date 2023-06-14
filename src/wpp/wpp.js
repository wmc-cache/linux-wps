/*
 *Created by chenyuxin
 */
var office = null;

function verify(obj)
{
    console.info(obj);
}

// 初始化插件及启动wpp
function InitWpp() 
{
    var app = null;
    app = Init("wpp");
    if (!app) {
        // 为了兼容
        office.setAttribute('data', './newfile.dps');
		var Interval_control = setInterval(
            function () {
                var app = office.Application;
                if (app) {
                    clearInterval(Interval_control);
                    Presentations_Add();
                }
            }, 500);
    }
    else {
        Presentations_Add();
    }
}
var _InitWpp = new CreateFunction("初始化wpp", InitWpp, []);
/* *****注释掉以前的初始化dmeo****
function Init(tagID) {
    if (office != undefined)
        office.Application.Quit();

    var iframe;
    iframe = document.getElementById(tagID);
    if (iframe.innerHTML.indexOf("application/x-wpp") > -1) {
        iframe.innerHTML = "";
    }
    var codes = [];
    codes.push('<object name="rpcwpp" id="rpcwpp_id" type="application/x-wpp" wpsshieldbutton="false" width="100%" height="100%">');
    codes.push('<param name="quality" value="high" />');
    codes.push('<param name="bgcolor" value="#ffffff" />');
    codes.push('<param name="Enabled" value="1" />');
    codes.push('<param name="allowFullScreen" value="true" />');
    codes.push('</object>');
    iframe.innerHTML = codes.join("");
    office = document.getElementById("rpcwpp_id");
    return office.Application;
}
******修改日期：2019-05-21****** */

var office;
function Init(tagID)
{
	var iframe;
	iframe = document.getElementById(tagID);
	iframe.innerHTML = '';
	var codes = [];
	codes.push("<object  name='webwpp' id='webwpp_id' type='application/x-wpp' width='100%' height='100%'> <param name='Enabled' value='1' />  </object>");

	iframe.innerHTML = codes.join("");
	office = document.getElementById("webwpp_id");

	window.onbeforeunload = function () 
	{
		office.Application.Quit();
	};

	window.onresize = function () 
	{
		console.log("ondrag");
		office.sltReleaseKeyboard();
	};

	return office.Application;
}

function PresentationBeforeSave11(pres, cancel) 
{
    alert("pres: " + pres);
    alert("cancel: " + cancel);
    cancel.SetValue(true);

}

function registerPresentationBeforeSave() 
{
    alert("registerEventHandler	");
    var app = office.Application;
    var ret = app.registerEvent("IID_EApplication", "PresentationBeforeSave", "PresentationBeforeSave11");
    alert("registerEvent ret: " + ret);
}

//四院封装接口
function openDocumentF()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/test.ppt", false);
    alert(aa);
}
var _openDocumentF = new CreateFunction("可编辑打开本地文档", openDocumentF, []);

function openDocumentRemote() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("https://10.90.128.241:8443/wps/保存到远程https.ppt", false);
    alert(aa);
}
var _openDocumentRemote = new CreateFunction("打开远程文档", openDocumentRemote, []);

function saveAsQ() 
{
    var app = office.Application;
    var aa = app.saveAs("/home/wpsgch/桌面/20190815.ppt");
    alert(aa);
}
var _saveAsQ = new CreateFunction("保存本地不弹框", saveAsQ, []);

function saveAs() 
{
    var app = office.Application;
    alert(app.saveAs());
}
var _saveAs = new CreateFunction("保存本地弹框", saveAs, []);

function saveURL() 
{
    var app = office.Application;
    var aa = app.saveURL("https://10.90.128.241:8443/wps/upload_l.jsp", "保存到远程https.ppt");
    alert(aa);
}
var _saveURL = new CreateFunction("保存到远程", saveURL, []);

function saveURL_url() 
{
    var app = office.Application;
    var aa = app.saveURL("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/save_rediect.jsp", "保存到远程url.pptx");
    alert(aa);
}
var _saveURL_url = new CreateFunction("保存到远程_重定向", saveURL_url, []);

function closewpp() 
{
    var app = office.Application;
    app.close();
}
var _closewpp = new CreateFunction("关闭文档", closewpp, []);

function AppIsLoad() 
{
    var app = office.Application;
    var aa = app.IsLoad();
    alert(aa);
}
var _AppIsLoad = new CreateFunction("IsLoad", AppIsLoad, []);

//以下是新增不落地接口测试--2018-08-22
function openDocumentRemote_s() {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://10.90.128.241:8080/wps/不落地测试.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s = new CreateFunction("打开远程文档不落地", openDocumentRemote_s, []);

function saveURL_s() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://10.90.128.241:8080/wps/upload_l.jsp", "不落地测试.ppt");
    alert(aa);
}
var _saveURL_s = new CreateFunction("保存到远程不落地", saveURL_s, []);

function saveURL_s_url()
{
	var app = office.Application;		
	var aa = app.saveURL_s("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/save_rediect.jsp", "不落地测试.pptx" );
    alert (aa);
}
var _saveURL_s_url = new CreateFunction("保存到远程不落地_重定向", saveURL_s_url, []);

function saveAsBase64Str() 
{
    base64str_dps = office.Application.saveAsBase64Str("pptx");
    alert(base64str_dps);
}
var _saveAsBase64Str = new CreateFunction("保存到base64", saveAsBase64Str, []);

function openDocumentFromBase64Str() 
{
    if (typeof base64str_dps != "undefined" && base64str_dps != "") {
        var aa = office.Application.openDocumentFromBase64Str(base64str_dps, false);
        alert(aa);
    } else {
        alert("请先调用saveasbase64str");
    }
}
var _openDocumentFromBase64Str = new CreateFunction("从base64打开", openDocumentFromBase64Str, []);

//以上是新增--2018-08-22

function GetName() 
{
    var app = office.Application;
    verify(app);
    alert(app.Name);
}
var _GetName = new CreateFunction("获取APP名", GetName, []);

function GetActivePresentationName() 
{
    var act = office.Application.ActivePresentation;
    verify(act);
    alert(act.Name);
}
var _GetActivePresentationName = new CreateFunction("获取ActivePresentation名", GetActivePresentationName, []);

//通过!String 没问题
function SetAppCaption() 
{
    var app = office.Application;
    verify(app);
    app.put_Caption("OldCaption");
    alert(app.Caption);

    alert("接下来换一个新的Caption:");
    app.put_Caption("NewCaption");
    alert(app.Caption);
}
var _SetAppCaption = new CreateFunction("设置窗口名称", SetAppCaption, []);

function GetPres1Name() 
{
    //alert(office.Application.Presentations.Count);
    alert(office.Application.Presentations.get_Count());
    var pres = office.Application.ActivePresentation;
    alert(pres.Name);
}
var _GetPres1Name = new CreateFunction("Int_Param_Sheet1_Name", GetPres1Name, []);

function GetUserName() 
{
    var app = office.Application;
    verify(app);
    alert(app.UserName);
}


function CloseActivePresentation() 
{
    var app = office.Application;
    //app.ActivePresentation.Close()
    app.ActivePresentation.Close()
}
var _CloseActivePresentation = new CreateFunction("关闭ActivePresentation", CloseActivePresentation, []);


function FullName() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    alert(pres.FullName);
}
var _FullName = new CreateFunction("Presentation_FullName", FullName, []);

function Presentation_Path() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    alert(pres.Path);
}
var _Presentation_Path = new CreateFunction("Presentation_Path", Presentation_Path, []);

function Presentation_HasTitleMaster() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    alert(pres.HasTitleMaster);
}
var _Presentation_HasTitleMaster = new CreateFunction("Presentation_HasTitleMaster", Presentation_HasTitleMaster, []);

function Presentation_NewWindow() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    pres.NewWindow();
}
var _Presentation_NewWindow = new CreateFunction("Presentation_NewWindow", Presentation_NewWindow, []);


function Presentation_Slides() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    //alert(pres.Slides.Count)
    alert(pres.Slides.get_Count())
}
var _Presentation_Slides = new CreateFunction("Presentation_Slides", Presentation_Slides, []);

function Presentation_Windows() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    alert(pres.Windows.get_Count())
}
var _Presentation_Windows = new CreateFunction("Presentation_Windows", Presentation_Slides, []);

function Presentation_FollowHyperlink() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    pres.FollowHyperlink('http://www.w3school.com.cn/');
}
var _Presentation_FollowHyperlink = new CreateFunction("Presentation_Windows", Presentation_FollowHyperlink, []);


function ActivePresentation_SaveAs() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    verify(pres);
    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('ActivePresentation_SaveAs');
    pres.SaveAs('/home/wpsgch/桌面/saveas0815.dps');

    console.timeEnd('ActivePresentation_SaveAs'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _ActivePresentation_SaveAs = new CreateFunction("当前文档保存另存为", ActivePresentation_SaveAs, []);


function Presentations_Add() 
{
    var app = office.Application;
    verify(app);
    var Presentations = app.Presentations;
    verify(Presentations);

    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('Presentations_Add');

    Presentations.Add();

    console.timeEnd('Presentations_Add'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _Presentations_Add = new CreateFunction("新建文档(没有slide)", Presentations_Add, []);

function Presentations_Close() 
{
    var app = office.Application;

    var pres = app.ActivePresentation;
    verify(pres);

    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('ActivePresentation.Close');
    pres.Close();
    console.timeEnd('ActivePresentation.Close'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _Presentations_Close = new CreateFunction("关闭当前文档", Presentations_Close, []);

function Presentations_Item() 
{
    var app = office.Application;
    //alert(app.Presentations.Count);
    alert(app.Presentations.get_Count());

    var pres = app.Presentations.get_Item(1);
    verify(pres)
    alert(pres.Name);
}
var _Presentations_Item = new CreateFunction("Presentations_Item", Presentations_Item, []);


function Presentations_Open() 
{
    var app = office.Application;
    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('Presentations.Open');
    var pres = app.Presentations.Open("sample.ppt");
    console.timeEnd('Presentations.Open'); // 结束定时器，打印时间。
    console.profileEnd();
    verify(pres);
}

function SlideShowSettings_Run() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    verify(pres);
    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('SlideShowSettings.Run');
    pres.SlideShowSettings.Run();
    console.timeEnd('SlideShowSettings.Run'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _SlideShowSettings_Run = new CreateFunction("进入播放", SlideShowSettings_Run, []);

function Slides_Add() 
{
    var app = office.Application;
    var pres = app.ActivePresentation;
    verify(pres);

    var slides = pres.Slides;
    verify(slides);
    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('Slides_Add');
    slides.Add(1, 12); //ppLayoutBlank
    console.timeEnd('Slides_Add'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _Slides_Add = new CreateFunction("添加幻灯片（ppLayoutBlank）", Slides_Add, []);

function Slides_Delete() 
{
    var slides = office.Application.ActivePresentation.Slides;
    console.profile('性能分析'); // 两种方法打印时间：console.profile及console.time，都可以获得函数的执行时间；
    console.time('Slides_Delete');
    slides.Item(1).Delete();
    console.timeEnd('Slides_Delete'); // 结束定时器，打印时间。
    console.profileEnd();
}
var _Slides_Delete = new CreateFunction("删除第一张幻灯片", Slides_Delete, []);

function Application_CommandBars() 
{
    alert(office.Application.CommandBars.Count);
}


function Application_Quit() 
{
    office.Application.Quit();
}

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

function InsertPic() 
{
    /*
	var app = office.Application;
	var pres = app.ActivePresentation;
	var slide = pres.Slides(1);
	verify(slide);
	slide.Shapes.AddPicture("123.gif", 0,1,10,10,500,500);
	*/
    var app = office.Application;
    var pres = app.ActivePresentation;
    var slides = pres.Slides;
    var slide = slides.Item(1);
    slide.Shapes.AddPicture("/home/wpsgch/桌面/1.png", 0, 1, 10, 10, 500, 500);
}
var  _InsertPic = new CreateFunction("插入图片", InsertPic, []);

// alert(rgItem.Count);
function EventCallBack(arg0) 
{
    alert("EventCallBack");
    alert(arg0);
    alert(arg0.Caption);
}

function RegisterEvent() 
{
    if (office_ctrl == undefined)
        return;
    alert(office_ctrl.registerEvent("DIID__CommandBarButtonEvents", "Click", "EventCallBack"));

}

function UnRegisterEvent() 
{
    if (office_ctrl == undefined)
        return;
    alert(office_ctrl.unRegisterEvent("DIID__CommandBarButtonEvents", "Click", "EventCallBack"));
}

function setToolbarAllVisibleT() 
{
    var app = office.Application;
    var aa = app.setToolbarAllVisible(true);
    alert(aa);
}
var _setToolbarAllVisibleT = new CreateFunction("显示工具栏", setToolbarAllVisibleT, []);

function setToolbarAllVisibleF() 
{
    var app = office.Application;
    var aa = app.setToolbarAllVisible(false);
    alert(aa);
}
var _setToolbarAllVisibleF = new CreateFunction("隐藏工具栏", setToolbarAllVisibleF, []);

//添加输入OFD格式_嵌入字体 日期：20190516
function ActivePresentation_SaveAsOFD()
{
	var app = office.Application;
	var pres = app.ActivePresentation;
	verify(pres);
	alert(pres.SaveAs('/home/wpsgch/桌面/saveasofd.ofd',102,-1));
}
var _ActivePresentation_SaveAsOFD = new CreateFunction("当前文档另存为OFD_嵌入字体", ActivePresentation_SaveAsOFD, []);

function ActivePresentation_SaveAsPDF()
{
	var app = office.Application;
	var pres = app.ActivePresentation;
	verify(pres);
	alert(pres.SaveAs('/home/wpsgch/桌面/saveaspdf.pdf',32));
}
var _ActivePresentation_SaveAsPDF = new CreateFunction("当前文档另存为PDF", ActivePresentation_SaveAsPDF, []);

//支持sessionid的接口 --20190528
function SendDataToServer_FormData_session()
{
	var headData = {};
	headData.filename = "测试.ppt";
	var aa = office.SendDataToServer_FormData("http://10.90.128.241:8080/servletTest_N/HelloServlet", "/home/wpsgch/桌面/test.ppt", JSON.stringify(headData), false);
	alert(aa);
}
var _SendDataToServer_FormData_session = new CreateFunction("上传文档至服务器_带seesion",SendDataToServer_FormData_session, []); 

function SendDataToServer_FormData_session_url()
{
	var headData = {};
	headData.filename = "测试url.pptx";
	var aa = office.SendDataToServer_FormData("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/sendserver_rediect.jsp", "/home/wpsgch/桌面/test.pptx", JSON.stringify(headData), false);
	alert(aa);
}
var _SendDataToServer_FormData_session_url = new CreateFunction("上传文档至服务器_带seesion_重定向",SendDataToServer_FormData_session_url, []);

function SendDataToServer_FormData_session_https()
{
	var headData = {};
	headData.filename = "测试https.dps";
	var aa = office.SendDataToServer_FormData("https://10.90.128.241:8443/servletTest_N/HelloServlet", "/home/wpsgch/桌面/test.dps", JSON.stringify(headData), false);
	alert(aa);
}
var _SendDataToServer_FormData_session_https = new CreateFunction("上传文档至服务器_带seesion_https",SendDataToServer_FormData_session_https, []);

function SendDataToServer_FormData_session_url_https()
{
	var headData = {};
	headData.filename = "httpsurl.dps";
	var aa = office.SendDataToServer_FormData("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/sendserver_rediect_https.jsp", "/home/wpsgch/桌面/test.dps", JSON.stringify(headData), false);
	alert(aa);
}
var _SendDataToServer_FormData_session_url_https = new CreateFunction("上传文档至服务器_带seesion_重定向_https",SendDataToServer_FormData_session_url_https, []);

function saveURL_FormData_session()
{ 
	var app = office.Application;
	var aa = app.saveURL_FormData("http://10.90.128.241:8080/servletTest/HelloServlet", "保存到远程.pptx" );
    alert (aa);
}
var _saveURL_FormData_session = new CreateFunction("保存到远程_带seesion",saveURL_FormData_session, []);

function saveURL_FormData_session_url()
{ 
    var app = office.Application;
	var aa = app.saveURL_FormData("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect.jsp", "保存到远程url.pptx" );
    alert (aa);
}
var _saveURL_FormData_session_url = new CreateFunction("保存到远程_带seesion_重定向",saveURL_FormData_session_url, []);
	
function saveURL_FormData_session_https()
{ 
    var app = office.Application;
	var aa = app.saveURL_FormData("https://10.90.128.241:8443/servletTest/HelloServlet", "保存到远程.pptx" );
    alert (aa);
}
var _saveURL_FormData_session_https = new CreateFunction("保存到远程_带seesion_https",saveURL_FormData_session_https, []);
	
function saveURL_FormData_session_url_https()
{ 
    var app = office.Application;
	var aa = app.saveURL_FormData("https://10.90.128.241:8443/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect_https.jsp", "保存到远程url.pptx" );
    alert (aa);
}
var _saveURL_FormData_session_url_https = new CreateFunction("保存到远程_带seesion_重定向_https",saveURL_FormData_session_url_https, []);

function saveURL_s_FormData_session()
{
	var app = office.Application;		
	var aa = app.saveURL_s_FormData("http://10.90.128.241:8080/servletTest/HelloServlet", "保存到远程不落地.pptx" );
    alert (aa);
}
var _saveURL_s_FormData_session = new CreateFunction("保存到远程不落地_带seesion", saveURL_s_FormData_session, []);
	
function saveURL_s_FormData_session_url()
{
	var app = office.Application;		
	var aa = app.saveURL_s_FormData("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect.jsp", "保存到远程不落地url.pptx" );
    alert (aa);
}
var _saveURL_s_FormData_session_url = new CreateFunction("保存到远程不落地_带seesion_重定向", saveURL_s_FormData_session_url, []);
	
function saveURL_s_FormData_session_https()
{
	var app = office.Application;		
	var aa = app.saveURL_s_FormData("https://10.90.128.241:8443/servletTest/HelloServlet", "保存到远程不落地.pptx" );
    alert (aa);
}
var _saveURL_s_FormData_session_https = new CreateFunction("保存到远程不落地_带seesion_https", saveURL_s_FormData_session_https, []);
	
function saveURL_s_FormData_session_url_https()
{
	var app = office.Application;		
	var aa = app.saveURL_s_FormData("https://10.90.128.241:8443/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect_https.jsp", "保存到远程不落地url.pptx" );
    alert (aa);
}
var _saveURL_s_FormData_session_url_https = new CreateFunction("保存到远程不落地_带seesion_重定向_https", saveURL_s_FormData_session_url_https, []);
	
function openDocumentRemote_FormData_session()
{
    var app = office.Application;
	var aa = app.openDocumentRemote("http://10.90.128.241:8080/servletTest/保存到远程.pptx", false);
    alert (aa);
} 
var _openDocumentRemote_FormData_session = new CreateFunction("打开远程文档_带seesion", openDocumentRemote_FormData_session, []);
	
function openDocumentRemote_FormData_session_url()
{
    var app = office.Application;
	var aa = app.openDocumentRemote("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/open_rediect.jsp", false);
    alert (aa);
} 
var _openDocumentRemote_FormData_session_url = new CreateFunction("打开远程文档_带seesion_重定向", openDocumentRemote_FormData_session_url, []);
	
function openDocumentRemote_FormData_session_https()
{
    var app = office.Application;
	var aa = app.openDocumentRemote("https://10.90.128.241:8443/servletTest/保存到远程.pptx", false);
    alert (aa);
} 
var _openDocumentRemote_FormData_session_https = new CreateFunction("打开远程文档_带seesion_https", openDocumentRemote_FormData_session_https, []);
	
function openDocumentRemote_FormData_session_url_https()
{
    var app = office.Application;
	var aa = app.openDocumentRemote("https://10.90.128.241:8443/servletTest/wps_webdemo/linux/src/wpp/open_rediect_https.jsp", false);
    alert (aa);
}
var _openDocumentRemote_FormData_session_url_https = new CreateFunction("打开远程文档_带seesion_重定向_https", openDocumentRemote_FormData_session_url_https, []);
	
function openDocumentRemote_s_FormData_session()
{
	var app = office.Application;		
	var aa = app.openDocumentRemote_s("http://10.90.128.241:8080/servletTest/保存到远程不落地.pptx", false);
    alert (aa);
}
var _openDocumentRemote_s_FormData_session = new CreateFunction("打开远程文档不落地_带seesion", openDocumentRemote_s_FormData_session, []); 
	
function openDocumentRemote_s_FormData_session_url()
{
	var app = office.Application;		
	var aa = app.openDocumentRemote_s("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/opens_rediect.jsp", false);
    alert (aa);
}
var _openDocumentRemote_s_FormData_session_url = new CreateFunction("打开远程文档不落地_带seesion_重定向", openDocumentRemote_s_FormData_session_url, []);
	
function openDocumentRemote_s_FormData_session_https()
{
	var app = office.Application;		
	var aa = app.openDocumentRemote_s("https://10.90.128.241:8443/servletTest/测试url.pptx", false);
    alert (aa);
} 
var _openDocumentRemote_s_FormData_session_https = new CreateFunction("打开远程文档不落地_带seesion_https", openDocumentRemote_s_FormData_session_https, []);
	
function openDocumentRemote_s_FormData_session_url_https()
{
	var app = office.Application;		
	var aa = app.openDocumentRemote_s("https://10.90.128.241:8443/servletTest/wps_webdemo/linux/src/wpp/opens_rediect_https.jsp", false);
    alert (aa);
}
var _openDocumentRemote_s_FormData_session_url_https = new CreateFunction("打开远程文档不落地_带seesion_重定向_https", openDocumentRemote_s_FormData_session_url_https, []);

//新增禁用实时备份以及复制剪切接口--2019-07-11
function Disable_BackUp()
{
	var app = office.Application;
	var aa = app.setForceBackUpEnabled(false);	//禁用实时备份功能
	alert (aa);
}
var _Disable_BackUp = new CreateFunction("禁用实时备份", Disable_BackUp,[]);
	
function Enable_BackUp()
{
	var app = office.Application;
	var aa = app.setForceBackUpEnabled(true);		//启用实时备份功能
	alert (aa);
}
var _Enable_BackUp = new CreateFunction("启用实时备份", Enable_BackUp, []);
	
function Disable_Copyandpaste()
{
	var app = office.Application;
	app.enableCopy(false);				//禁用复制
	app.enableCut(false);				//禁用剪切
}
var _Disable_Copyandpaste = new CreateFunction("禁止复制剪切", Disable_Copyandpaste, []);
 
function Enable_Copyandpaste()
{
	var app = office.Application;
	app.enableCopy(true);				//启用复制
	app.enableCut(true);				//启用剪切
}
var _Enable_Copyandpaste = new CreateFunction("启用复制剪切", Enable_Copyandpaste, []);
	
//新增设置临时文件路径接口---20190716
function setTmpFilepath()
{
	alert(office.setTmpFilepath("/home/wpsgch/桌面"));//设置后，调用打开远程文档接口打开文档后保存，路径将变为设置后的路径
}
var _setTmpFilepath = new CreateFunction("设置临时路径", setTmpFilepath, []);

//新增注册打印事件接口
function RegisterPrintOutPageSetEvent() 
{ 
	var appex = office.Application.ApplicationEx; 
	var ret = appex.registerEvent("DIID_ApplicationEventsEx","DocumentAfterPrint","EventCallBackPrintOutPageSet"); 
	alert("PrintOutPageSet--"+ret); 
}
var _RegisterPrintOutPageSetEvent = new CreateFunction("注册打印事件", RegisterPrintOutPageSetEvent, []);

function EventCallBackPrintOutPageSet(pres, pageset)
{
	var range = pageset.get_PrintOutRange();
	if (range == 1){
		alert("全选");
	}else if (range == 2){
		alert("选中");
	}else if (range == 3){
		alert("当前");
	}
	alert("页码范围");
	alert(pageset.get_PrintOutPages());
}

//增加设置默认保存文件名接口---20190819
function setFramesaveAs()
{
	var app = office.Application;
    alert(app.saveAs("","test0819.ppt"));
}
var _setFramesaveAs = new CreateFunction("设置默认保存文件名", setFramesaveAs, []);

//增加打印大纲方式，打印到文件接口----20190912
function PrintOutOutline()
{
	var printOptions = office.Application.ActivePresentation.get_PrintOptions();
	//设置打印大纲
	printOptions.put_OutputType(6);
	printOptions.put_RangeType(1);
	//实际打印到文档
	alert(office.Application.ActivePresentation.PrintOut(-1, -1, "/home/wpsgch/桌面/outline.pdf"));
	
}
var _PrintOutOutline = new CreateFunction("当前文档打印到文件_大纲", PrintOutOutline, []);

//增加合并文档接口----20191022
function InsertFormFile()
{
	var fileName = "/home/wpsgch/桌面/test.pptx";
	var fromfile = office.Application.ActivePresentation.Slides.InsertFromFile(fileName,1,1,2);
	alert(fromfile);  
}
var _InsertFormFile = new CreateFunction("合并文档", InsertFormFile, []);

//增加返回自定义请求头接口----20191112
function saveURL_CustomParam_FormData()
{
	var jsondata = {haha:"aaa"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("http://10.90.128.241:8080/servletTest/HelloServlet", "test.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam_FormData = new CreateFunction("保存到远程_返回自定义数据_带session", saveURL_CustomParam_FormData, []);

function saveURL_CustomParam_FormData_url()
{
	var jsondata = {haha:"bbb"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect.jsp", "testurl.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam_FormData_url = new CreateFunction("保存到远程_返回自定义数据_带session_重定向", saveURL_CustomParam_FormData_url, []);

function saveURL_CustomParam_FormData_https()
{
	var jsondata = {haha:"ccc"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("https://10.90.128.241:8443/servletTest/HelloServlet", "testhttps.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam_FormData_https = new CreateFunction("保存到远程_返回自定义数据_带session_https", saveURL_CustomParam_FormData_https, []);

function saveURL_CustomParam_FormData_url_https()
{
	var jsondata = {haha:"dddd"}; 
	var aa = office.Application.saveURL_CustomParam_FormData("https://10.90.128.241:8443/servletTest/wps_webdemo/linux/src/wpp/sendserver_rediect_https.jsp", "testurlhttps.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam_FormData_url_https = new CreateFunction("保存到远程_返回自定义数据_带session_重定向_https", saveURL_CustomParam_FormData_url_https, []);

function saveURL_CustomParam()
{
	var jsondata = {key1:"aaa"}; 
	var aa = office.Application.saveURL_CustomParam("http://10.90.128.241:8080/servletTest/upload_l.jsp", "111.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam = new CreateFunction("保存到远程_返回自定义数据", saveURL_CustomParam, []);

function saveURL_CustomParam_url()
{
	var jsondata = {key1:"bbb"}; 
	var aa = office.Application.saveURL_CustomParam("http://10.90.128.241:8080/servletTest/wps_webdemo/linux/src/wpp/save_rediect.jsp", "111url.pptx", JSON.stringify(jsondata)); 
	alert(aa); 
}
var _saveURL_CustomParam_url = new CreateFunction("保存到远程_返回自定义数据_重定向", saveURL_CustomParam_url, []);

//新增四个接口---20191220
//1、UploadFileToServer(url, loaclpath, paraminfo)，上传文档到服务器，3个参数为必填项；
function UploadFileToServer()
{
	var jsondata = 
	{
		fileName:"上传!@#$%.ppt",  		//上传到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,		 	//是否返回响应头信息
		bDelLocalFile:false,			//是否删除本地
		customFromData:
		{
			"key1":"Upload1测试",
			"key2":"Upload2测试"
		},
		customHeadData:
		{
			"Cookie":"Uploadcookieppt",
			"key1":"Upload111金山",
			"key2":"Upload222金山"
		}
	};
	var aa = office.UploadFileToServer("http://10.90.128.241:8080/servletTest_N/HelloServlet", "/home/wpsgch/桌面/test.ppt", JSON.stringify(jsondata));
	alert(aa);
}
var _UploadFileToServer = new CreateFunction("上传远程", UploadFileToServer, []);

function UploadFileToServer_url()
{
	var jsondata = 
	{
		fileName:"上传重定向.pptx",  //上传到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,		 	//是否返回响应头信息
		bDelLocalFile:false,			//是否删除本地
		customFromData:
		{
			"key1":"Upload重定向",
			"key2":"Upload重定向"
		},
		customHeadData:
		{
			"Cookie":"UploadCookieurlpptx",
			"key1":"Uploadurl111测试",
			"key2":"Uploadurl222测试"
		}
	};
	var aa = office.UploadFileToServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect.jsp", "/home/wpsgch/桌面/test.pptx", JSON.stringify(jsondata));
	alert(aa);
}
var _UploadFileToServer_url = new CreateFunction("上传远程_重定向", UploadFileToServer_url, []);

function UploadFileToServer_https()
{
	var jsondata = 
	{
		fileName:"上传https.dps",  		//中文名不允许？
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,		 	//是否返回响应头信息
		bDelLocalFile:false,			//是否删除本地
		customFromData:
		{
			"key1":"Uploadhttps",
			"key2":"Uploadhttps"
		},
		customHeadData:
		{
			"Cookie":"UploadCookiehttpsdps",
			"key1":"Uploadhttps111测试",
			"key2":"Uploadhttps222测试"
		}
	};
	var aa = office.UploadFileToServer("https://10.90.128.241:8443/servletTest_N/HelloServlet", "/home/wpsgch/桌面/test.dps", JSON.stringify(jsondata));
	alert(aa);
}
var _UploadFileToServer_https = new CreateFunction("上传远程_https", UploadFileToServer_https, []);

function UploadFileToServer_https_url()
{
	var jsondata = 
	{
		fileName:"上传https重定向.pptx",  		//中文名不允许？
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,		 	//是否返回响应头信息
		bDelLocalFile:false,			//是否删除本地
		customFromData:
		{
			"key1":"Uploadhttps重定向",
			"key2":"Uploadhttps重定向"
		},
		customHeadData:
		{
			"Cookie":"UploadCookiehttpspptx",
			"key1":"Uploadhttps111重定向",
			"key2":"Uploadhttps222重定向"
		}
	};
	var aa = office.UploadFileToServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect_https.jsp", "/home/wpsgch/桌面/test.pptx", JSON.stringify(jsondata));
	alert(aa);
}
var _UploadFileToServer_https_url = new CreateFunction("上传远程_https_重定向", UploadFileToServer_https_url, []);

//2、DownloadFileFromServer(url, loaclpath, isGetResponseHead)，下载远程文档到本地，3个参数为必填项；
function DownloadFileFromServer()
{
	var jsondata = 
	{
		isGetResponseHead:true,		//是否返回响应头信息
		customHeadData:
		{
			"Cookie":"DownloadCookie",
			"key1":"Download",
			"key2":"Download"
		}
	}
	var aa = office.DownloadFileFromServer("http://10.90.128.241:8080/servletTest_N/http测试.pptx", "/home/wpsgch/桌面/aa/!@#$%/Download.pptx", JSON.stringify(jsondata));
	alert(aa);
}
var _DownloadFileFromServer = new CreateFunction("下载远程_新", DownloadFileFromServer, []);

function DownloadFileFromServer_url()
{
	var jsondata = 
	{
		isGetResponseHead:true,		//是否返回响应头信息
		customHeadData:
		{
			"Cookie":"DownloadCookieurl",
			"key1":"Download重定向",
			"key2":"Download重定向"
		}
	}
	var aa = office.DownloadFileFromServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/open_rediect.jsp", "/home/wpsgch/桌面/Download重定向.pptx", JSON.stringify(jsondata));
	alert(aa);
}
var _DownloadFileFromServer_url = new CreateFunction("下载远程_重定向", DownloadFileFromServer_url, []);

function DownloadFileFromServer_https()
{
	var jsondata = 
	{
		isGetResponseHead:true,		//是否返回响应头信息
		customHeadData:
		{
			"Cookie":"DownloadCookiehttps",
			"key1":"Downloadhttps",
			"key2":"Downloadhttps"
		}
	}
	var aa = office.DownloadFileFromServer("https://10.90.128.241:8443/servletTest_N/上传https.pptx", "/home/wpsgch/桌面/Downloadhttps.ppt", JSON.stringify(jsondata));
	alert(aa);
}
var _DownloadFileFromServer_https = new CreateFunction("下载远程_https", DownloadFileFromServer_https, []);

function DownloadFileFromServer_https_url()
{
	var jsondata = 
	{
		isGetResponseHead:true,		//是否返回响应头信息
		customHeadData:
		{
			"Cookie":"DownloadhttpsCookieurl",
			"key1":"Downloadhttps重定向",
			"key2":"Downloadhttps重定向"
		}
	}
	var aa = office.DownloadFileFromServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/open_rediect_https.jsp", "/home/wpsgch/桌面/Downloadhttps重定向.pptx", JSON.stringify(jsondata));
	alert(aa);
}
var _DownloadFileFromServer_https_url = new CreateFunction("下载远程_https_重定向", DownloadFileFromServer_https_url, []);

//3、SaveDocumentToServer(url, paraminfo)，保存远程，支持落地、不落地，2个参数为必填项；
//保存远程
function SaveDocumentToServer_F()
{
	var jsondata = 
	{
		fileName:"http落地测试.uop",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:false,				//是否不落地保存 false：落地 
		customFromData:
		{
			"key1":"SaveDocument测试",
			"key2":"SaveDocument测试"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookiedps",
			"key1":"SaveDocument111测试",
			"key2":"SaveDocument222测试"
		}
	}
	var aa = office.Application.SaveDocumentToServer("http://10.90.128.241:8080/servletTest_N/HelloServlet",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_F = new CreateFunction("保存远程", SaveDocumentToServer_F, []);

function SaveDocumentToServer_F_url()
{
	var jsondata = 
	{
		fileName:"http落地测试重定向.pptx",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:false,				//是否不落地保存 false：落地 
		customFromData:
		{
			"key1":"SaveDocument重定向",
			"key2":"SaveDocument重定向"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookieurl",
			"key1":"SaveDocument111重定向",
			"key2":"SaveDocument222重定向"
		}
	}
	var aa = office.Application.SaveDocumentToServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect.jsp",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_F_url = new CreateFunction("保存远程_重定向", SaveDocumentToServer_F_url, []);

function SaveDocumentToServer_F_https()
{
	var jsondata = 
	{
		fileName:"https落地测试.ppt",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:false,				//是否不落地保存 false：落地 
		customFromData:
		{
			"key1":"SaveDocument测试https",
			"key2":"SaveDocument测试https"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookiehttps",
			"key1":"SaveDocument111测试https",
			"key2":"SaveDocument222测试https"
		}
	}
	var aa = office.Application.SaveDocumentToServer("https://10.90.128.241:8443/servletTest_N/HelloServlet",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_F_https = new CreateFunction("保存远程_https", SaveDocumentToServer_F_https, []);

function SaveDocumentToServer_F_https_url()
{
	var jsondata = 
	{
		fileName:"https落地测试重定向.pptx",  //保存到远程文件名
		isGetBodyResult:true,				//是否需要返回请求数据
		isGetResponseHead:true,				//是否返回响应头信息
		isNoTmpFile:false,					//是否不落地保存 false：落地 
		customFromData:
		{
			"key1":"SaveDocumenthttps重定向",
			"key2":"SaveDocumenthttps重定向"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumenthttpsCookieurl",
			"key1":"SaveDocument111重定向https",
			"key2":"SaveDocument222重定向https"
		}
	}
	var aa = office.Application.SaveDocumentToServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect_https.jsp",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_F_https_url = new CreateFunction("保存远程_https_重定向", SaveDocumentToServer_F_https_url, []);

//保存远程不落地
function SaveDocumentToServer_T()
{
	var jsondata = 
	{
		fileName:"http不落地测试.ppt",  //保存到远程文件名
		isGetBodyResult:true,				//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:true,				//是否不落地保存 true：不落地 
		customFromData:
		{
			"key1":"SaveDocument不落地",
			"key2":"SaveDocument不落地"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookie",
			"key1":"SaveDocument不落地111",
			"key2":"SaveDocument不落地222"
		}
	}
	var aa = office.Application.SaveDocumentToServer("http://10.90.128.241:8080/servletTest_N/HelloServlet",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_T = new CreateFunction("保存远程", SaveDocumentToServer_T, []);

function SaveDocumentToServer_T_url()
{
	var jsondata = 
	{
		fileName:"http不落地测试重定向.pptx",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:true,				//是否不落地保存 true：不落地 
		customFromData:
		{
			"key1":"SaveDocument不落地url",
			"key2":"SaveDocument不落地url"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookieurl",
			"key1":"SaveDocument不落地111url",
			"key2":"SaveDocument不落地222url"
		}
	}
	var aa = office.Application.SaveDocumentToServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect.jsp",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_T_url = new CreateFunction("保存远程不落地_重定向", SaveDocumentToServer_T_url, []);

function SaveDocumentToServer_T_https()
{
	var jsondata = 
	{
		fileName:"https不落地测试.dps",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:true,				//是否不落地保存 true：不落地 
		customFromData:
		{
			"key1":"SaveDocument不落地https",
			"key2":"SaveDocument不落地"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookiehttps",
			"key1":"SaveDocument不落地111",
			"key2":"SaveDocument不落地222https"
		}
	}
	var aa = office.Application.SaveDocumentToServer("https://10.90.128.241:8443/servletTest_N/HelloServlet",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_T_https = new CreateFunction("保存远程不落地_https", SaveDocumentToServer_T_https, []);

function SaveDocumentToServer_T_https_url()
{
	var jsondata = 
	{
		fileName:"https不落地测试重定向.pptx",  //保存到远程文件名
		isGetBodyResult:true,			//是否需要返回请求数据
		isGetResponseHead:true,			//是否返回响应头信息
		isNoTmpFile:true,				//是否不落地保存 true：不落地 
		customFromData:
		{
			"key1":"SaveDocument不落地_https_url",
			"key2":"SaveDocument不落地_https_url"
		},
		customHeadData:
		{
			"Cookie":"SaveDocumentCookiehttpsurl",
			"key1":"SaveDocument不落地111_https_url",
			"key2":"SaveDocument不落地222_https_url"
		}
	}
	var aa = office.Application.SaveDocumentToServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/send_rediect_https.jsp",  JSON.stringify(jsondata));
	alert(aa);
}	
var _SaveDocumentToServer_T_https_url = new CreateFunction("保存远程不落地_https_重定向", SaveDocumentToServer_T_https_url, []);


//4、OpenDocumentFromServer(url, paraminfo)，打开远程文档，支持落地、不落地，2个参数为必填项。
//打开远程
function OpenDocumentFromServer_F()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:true,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,			//是否不落地保存 false：落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookie0",
			"key1":"OpenDocument落地1",
			"key2":"OpenDocument落地2"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("http://10.90.128.241:8080/servletTest_N/http落地测试.uop",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_F = new CreateFunction("打开远程", OpenDocumentFromServer_F, []);

function OpenDocumentFromServer_F_url()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:true,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,			//是否不落地保存 false：落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookieurl0",
			"key1":"OpenDocument落地1url",
			"key2":"OpenDocument落地2url"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/open_rediect.jsp",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_F_url = new CreateFunction("打开远程_重定向", OpenDocumentFromServer_F_url, []);

function OpenDocumentFromServer_F_https()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:true,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,			//是否不落地保存 false：落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookiehttps0",
			"key1":"OpenDocument落地1_https",
			"key2":"OpenDocument落地2_https"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("https://10.90.128.241:8443/servletTest_N/https落地测试.ppt",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_F_https = new CreateFunction("打开远程_https", OpenDocumentFromServer_F_https, []);

function OpenDocumentFromServer_F_https_url()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:false,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:false,			//是否不落地保存 false：落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookiehttpsurl0",
			"key1":"OpenDocument落地1_https_url",
			"key2":"OpenDocument落地2_https_url"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/open_rediect_https.jsp",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_F_https_url = new CreateFunction("打开远程_https_重定向", OpenDocumentFromServer_F_https_url, []);

//打开远程不落地
function OpenDocumentFromServer_T()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:false,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:true,			//是否不落地保存 true：不落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookie000",
			"key1":"OpenDocument不落地111",
			"key2":"OpenDocument不落地222"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("http://10.90.128.241:8080/servletTest_N/http不落地测试.ppt",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_T = new CreateFunction("打开远程", OpenDocumentFromServer_T, []);

function OpenDocumentFromServer_T_url()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:false,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:true,			//是否不落地保存 true：不落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookieurl000",
			"key1":"OpenDocument不落地111_url",
			"key2":"OpenDocument不落地222_url"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("http://10.90.128.241:8080/servletTest_N/wps_webdemo/linux/src/wpp/opens_rediect.jsp",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_T_url = new CreateFunction("打开远程不落地_重定向", OpenDocumentFromServer_T_url, []);

function OpenDocumentFromServer_T_https()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:false,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:true,			//是否不落地保存 true：不落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookie000_https",
			"key1":"OpenDocument不落地111_https",
			"key2":"OpenDocument不落地222_https"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("https://10.90.128.241:8443/servletTest_N/https不落地测试.dps",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_T_https = new CreateFunction("打开远程不落地_https", OpenDocumentFromServer_T_https, []);

function OpenDocumentFromServer_T_https_url()
{
	var jsondata = 
	{
		password:"1",				//打开文档所需的密码
		readOnly:false,				//是否只读打开文档 
		isGetResponseHead:true,		//是否返回响应头信息
		isNoTmpFile:true,			//是否不落地保存 true：不落地
		customHeadData:
		{
			"Cookie":"OpenDocumentCookie000_https_url",
			"key1":"OpenDocument不落地111_https_url",
			"key2":"OpenDocument不落地222_https_url"
		}
	}
	var aa = office.Application.OpenDocumentFromServer("https://10.90.128.241:8443/servletTest_N/wps_webdemo/linux/src/wpp/opens_rediect_https.jsp",  JSON.stringify(jsondata));
	alert(aa);	
}
var _OpenDocumentFromServer_T_https_url = new CreateFunction("打开远程不落地_重定向", OpenDocumentFromServer_T_https_url, []);

//补充上传、下载接口---20191225
function SendDataToServer() 
{
    var aa = office.SendDataToServer("https://10.90.128.241:8443/wps/upload_l.jsp", "/home/wpsgch/桌面/test.pptx", "测试https.pptx", false);
    alert(aa);
}
var _SendDataToServer = new CreateFunction("本地文档上传到远程", SendDataToServer, []);

function DownLoadServerFile()
{
    var aa = office.DownLoadServerFile("https://10.90.128.241:8443/wps/测试https.pptx", "/home/wpsgch/桌面/测试下载https.pptx");
    alert(aa);
}
var _DownLoadServerFile = new CreateFunction("下载远程文档至本地", DownLoadServerFile, []);

//添加注册关闭事件---20200116
function wppCloseCallback()
{
	alert("wppCloseCallback");
}
function RegisterBeforeCloseEvent()
{
	var app = office.Application; 
	var ret = app.registerEvent("IID_EApplication","PresentationClose","wppCloseCallback"); 
	alert("PresentationClose--"+ret); 
}
var _RegisterBeforeCloseEvent = new CreateFunction("注册关闭事件", RegisterBeforeCloseEvent, []);

//添加取消注册关闭事件---20200311
function unRegisterBeforeCloseEvent()
{
	var app = office.Application; 
	var ret = app.unRegisterEvent("IID_EApplication","PresentationClose","wppCloseCallback"); 
	alert("PresentationClose--"+ret); 
}
var _unRegisterBeforeCloseEvent = new CreateFunction("注册关闭事件", unRegisterBeforeCloseEvent, []);

//添加隐藏/显示关闭所有文档、保存所有文档按钮接口---20200408
function CommandBarsL_false()
{
	var app = office.Application; 
	var aa = app.CommandBars.get_Item("File").Controls.get_Item("关闭所有文档(&L)").Visible = false;
	alert(aa); 
}
var _CommandBarsL_false = new CreateFunction("隐藏关闭所有文档按钮", CommandBarsL_false, []);
function CommandBarsL_true()
{
	var app = office.Application; 
	var aa = app.CommandBars.get_Item("File").Controls.get_Item("关闭所有文档(&L)").Visible = true;
	alert(aa); 
}
var _CommandBarsL_true = new CreateFunction("显示关闭所有文档按钮", CommandBarsL_true, []);

function CommandBarsE_false()
{
	var app = office.Application; 
	var aa = app.CommandBars.get_Item("File").Controls.get_Item("保存所有文档(&E)").Visible = false;
	alert(aa); 
}
var _CommandBarsE_false = new CreateFunction("隐藏保存所有文档按钮", CommandBarsE_false, []);
function CommandBarsE_true()
{
	var app = office.Application; 
	var aa = app.CommandBars.get_Item("File").Controls.get_Item("保存所有文档(&E)").Visible = true;
	alert(aa); 
}
var _CommandBarsE_true = new CreateFunction("显示保存所有文档按钮", CommandBarsE_true, []);


// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 函数调用区 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// 加载后执行
window.onload = function () {
    InitLayui();
}

// <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 函数调用区 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
