// 函数说明, 代码显示控件 对应的 elementID
var elementId_codeDoc = "codeDoc";
var elementId_code = "code";

var layer, form, $;
function InitLayui() {
    // use layer
    layui.use(['layer', 'form'], function () {
        layer = layui.layer;
        form = layui.form;
        $ = layui.jquery;
        // layer.msg('Hello World');
    });

    // use element (折叠面板依赖它)
    layui.use('element', function () {
        // element = layui.element;
    });

    // window.alert = function (text) {
    //     layui.layer.alert(text);
    // }
    //加载code模块
    // layui.use('code', function () {
    //     // do nothing
    // });
}

function FlushCodeInfo(tips, code) {
    document.getElementById(elementId_codeDoc).innerText = tips;
    document.getElementById(elementId_code).innerText = code;
    // flush code style 
    // layui.code({
    //     title: 'JavaScript'
    //     , about: false      //剔除关于
    // });
}

function CreateFunction(tips, func, args) {
	if (tips == "初始化wpp")
		console.log(tips);
    return function() {
        FlushCodeInfo(tips, func.toString());
        func.apply(window, args);
    }
}
// 使用示例: var OnClicked = new CreateFunction("hello world", fn, [5, 10]);