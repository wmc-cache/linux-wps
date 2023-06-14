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

//100kb
function openDocumentF()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/二次开发效率样张/WPP/PPT_100KB.ppt", false);
    alert(aa);
}
var _openDocumentF = new CreateFunction("打开本地文档_100kb", openDocumentF, []);

function saveURL() 
{
    var app = office.Application;
    var aa = app.saveURL("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程100KB.ppt");
    alert(aa);
}
var _saveURL = new CreateFunction("保存到远程_100kb", saveURL, []);

function openDocumentRemote() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("http://192.168.39.80:8080/wps/保存到远程100KB.ppt", false);
    alert(aa);
}
var _openDocumentRemote = new CreateFunction("打开远程文档_100kb", openDocumentRemote, []);

function saveURL_s() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程不落地100KB.ppt");
    alert(aa);
}
var _saveURL_s = new CreateFunction("保存到远程不落地_100kb", saveURL_s, []);

function openDocumentRemote_s()
 {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://192.168.39.80:8080/wps/保存到远程不落地100KB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s = new CreateFunction("打开远程文档不落地_100kb", openDocumentRemote_s, []);


//10mb
function openDocumentF_10mb()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/二次开发效率样张/WPP/PPT_10MB.ppt", false);
    alert(aa);
}
var _openDocumentF_10mb = new CreateFunction("打开本地文档_10mb", openDocumentF_10mb, []);

function saveURL_10mb() 
{
    var app = office.Application;
    var aa = app.saveURL("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程10MB.ppt");
    alert(aa);
}
var _saveURL_10mb = new CreateFunction("保存到远程_10mb", saveURL_10mb, []);

function openDocumentRemote_10mb() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("http://192.168.39.80:8080/wps/保存到远程10MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_10mb = new CreateFunction("打开远程文档_10mb", openDocumentRemote_10mb, []);

function saveURL_s_10mb() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程不落地10MB.ppt");
    alert(aa);
}
var _saveURL_s_10mb = new CreateFunction("保存到远程不落地_10mb", saveURL_s_10mb, []);

function openDocumentRemote_s_10mb()
 {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://192.168.39.80:8080/wps/保存到远程不落地10MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s_10mb = new CreateFunction("打开远程文档不落地_10mb", openDocumentRemote_s_10mb, []);


//50mb
function openDocumentF_50mb()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/二次开发效率样张/WPP/PPT_50MB.ppt", false);
    alert(aa);
}
var _openDocumentF_50mb = new CreateFunction("打开本地文档_50mb", openDocumentF_50mb, []);

function saveURL_50mb() 
{
    var app = office.Application;
    var aa = app.saveURL("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程50MB.ppt");
    alert(aa);
}
var _saveURL_50mb = new CreateFunction("保存到远程_50mb", saveURL_50mb, []);

function openDocumentRemote_50mb() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("http://192.168.39.80:8080/wps/保存到远程50MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_50mb = new CreateFunction("打开远程文档_50mb", openDocumentRemote_50mb, []);

function saveURL_s_50mb() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程不落地50MB.ppt");
    alert(aa);
}
var _saveURL_s_50mb = new CreateFunction("保存到远程不落地_50mb", saveURL_s_50mb, []);

function openDocumentRemote_s_50mb()
 {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://192.168.39.80:8080/wps/保存到远程不落地50MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s_50mb = new CreateFunction("打开远程文档不落地_50mb", openDocumentRemote_s_50mb, []);

//100mb
function openDocumentF_100mb()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/二次开发效率样张/WPP/PPT_100MB.ppt", false);
    alert(aa);
}
var _openDocumentF_100mb = new CreateFunction("打开本地文档_100mb", openDocumentF_100mb, []);

function saveURL_100mb() 
{
    var app = office.Application;
    var aa = app.saveURL("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程100MB.ppt");
    alert(aa);
}
var _saveURL_100mb = new CreateFunction("保存到远程_100mb", saveURL_100mb, []);

function openDocumentRemote_100mb() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("http://192.168.39.80:8080/wps/保存到远程100MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_100mb = new CreateFunction("打开远程文档_100mb", openDocumentRemote_100mb, []);

function saveURL_s_100mb() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程不落地100MB.ppt");
    alert(aa);
}
var _saveURL_s_100mb = new CreateFunction("保存到远程不落地_100mb", saveURL_s_100mb, []);

function openDocumentRemote_s_100mb()
 {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://192.168.39.80:8080/wps/保存到远程不落地100MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s_100mb = new CreateFunction("打开远程文档不落地_100mb", openDocumentRemote_s_100mb, []);


//200mb
function openDocumentF_200mb()
{
    var app = office.Application;
    var aa = app.openDocument("/home/wpsgch/桌面/二次开发效率样张/WPP/PPT_200MB.ppt", false);
    alert(aa);
}
var _openDocumentF_200mb = new CreateFunction("打开本地文档_200mb", openDocumentF_200mb, []);

function saveURL_200mb() 
{
    var app = office.Application;
    var aa = app.saveURL("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程200MB.ppt");
    alert(aa);
}
var _saveURL_200mb = new CreateFunction("保存到远程_200mb", saveURL_200mb, []);

function openDocumentRemote_200mb() 
{
    var app = office.Application;
    var aa = app.openDocumentRemote("http://192.168.39.80:8080/wps/保存到远程200MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_200mb = new CreateFunction("打开远程文档_200mb", openDocumentRemote_200mb, []);

function saveURL_s_200mb() 
{
    var app = office.Application;
    var aa = app.saveURL_s("http://192.168.39.80:8080/wps/upload_l.jsp", "保存到远程不落地200MB.ppt");
    alert(aa);
}
var _saveURL_s_200mb = new CreateFunction("保存到远程不落地_200mb", saveURL_s_200mb, []);

function openDocumentRemote_s_200mb()
 {
    var app = office.Application;
    var aa = app.openDocumentRemote_s("http://192.168.39.80:8080/wps/保存到远程不落地200MB.ppt", false);
    alert(aa);
}
var _openDocumentRemote_s_200mb = new CreateFunction("打开远程文档不落地_200mb", openDocumentRemote_s_200mb, []);

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


// >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 函数调用区 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
// 加载后执行
window.onload = function () 
{
    InitLayui();
}

// <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 函数调用区 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
