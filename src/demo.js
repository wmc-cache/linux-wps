function InitLayui() {
    // use layer
    layui.use(['layer', 'form'], function () {
        var layer = layui.layer,
            form = layui.form;

        // layer.msg('Hello World');
    });

    // use element (折叠面板依赖它)
    layui.use('element', function () {
        element = layui.element;

        // tab 切换
        element.on('nav(filter_tab)', function (elem) {
            // Tips: 使用的是 jQuery 操作 DOM 的方式.
            var indexName = elem.attr("id");
            if (indexName === "a_wps") {
                SwitchTab(0);
                UpdateDocBar(0);
            } else if (indexName === "a_wpp") {
                SwitchTab(1);
                UpdateDocBar(1);
            } else if (indexName === "a_et") {
                SwitchTab(2);
                UpdateDocBar(2);
            }
        });
    });
    // window.alert = function (text) {
    //     layui.layer.alert(text);
    // }
}

// 切换到相应 tab. 0: wps  1: wpp  2: et
function SwitchTab(crtTabIndex) {
    var iframe_wps = document.getElementById("iframe_wps");
    var iframe_wpp = document.getElementById("iframe_wpp");
    var iframe_et = document.getElementById("iframe_et");

    if (iframe_wps && iframe_wpp && iframe_et) {
        switch (crtTabIndex) {
            case 0:
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wps.setAttribute("height", "100%");
                break;
            case 1:
                iframe_wps.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "100%");
                break;
            case 2:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "100%");
                break;
            default:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                break;
        }

        console.log("wps height: " + iframe_wps.getAttribute("height"));
        console.log("wpp height: " + iframe_wpp.getAttribute("height"));
        console.log("et height: " + iframe_et.getAttribute("height"));
    }
}

// 显示tab对应的文档 0: wps  1: wpp  2: et
function UpdateDocBar(indexId) {
    var docBar = document.getElementById("divLayout_docBar");
    var docChilds = docBar.children;
    for (var i=0; i<docChilds.length; i++) {
        var docTypeVal = docChilds[i].getAttribute("itemShowType");
        var bShow = false;
        if (docTypeVal == -1) {
            bShow = true;
        } else {
            if (docTypeVal == indexId) {
                bShow = true;
            }
        }

        if (bShow) {
            docChilds[i].style.display="";      // show
        } else {
            docChilds[i].style.display="none";  // hide
        }
    }
}


window.onload = function () {
    InitLayui();

    // 默认切换到 wps
    SwitchTab(0);
    UpdateDocBar(0);
}