/// <reference path="./node_modules/wps-jsapi/index.d.ts" />

declare var XDomainRequest:{
    prototype: XMLHttpRequest;
    new(): XMLHttpRequest;
};
class name1  {
    constructor(){}
} 

function getHttpObj():XMLHttpRequest {
    
    var httpobj:XMLHttpRequest;
    if (IEVersion() < 10) {
        try {
            httpobj = new XDomainRequest();
        } catch (e1) {
            httpobj =  createXHR();
            
        }
    } else {
        httpobj = new name1();
    }
    return httpobj;
}

//兼容IE低版本的创建xmlhttprequest对象的方法
function createXHR() {
    
    if (typeof XMLHttpRequest != 'undefined') { //兼容高版本浏览器
        return new XMLHttpRequest();
    } else if (typeof ActiveXObject != 'undefined') { //IE6 采用 ActiveXObject， 兼容IE6
        var versions = [ //由于MSXML库有3个版本，因此都要考虑
            'MSXML2.XMLHttp.6.0',
            'MSXML2.XMLHttp.3.0',
            'MSXML2.XMLHttp'
        ];

        for (var i = 0; i < versions.length; i++) {
            try {
                return new ActiveXObject(versions[i]);
            } catch (e) {
                //跳过
            }
        }
    } else {
        throw new Error('您的浏览器不支持XHR对象');
    }
}
function IEVersion():Number {
    var userAgent = navigator.userAgent; //取得浏览器的userAgent字符串  
    var isIE = userAgent.indexOf("compatible") > -1 && userAgent.indexOf("MSIE") > -1; //判断是否IE<11浏览器  
    var isEdge = userAgent.indexOf("Edge") > -1 && !isIE; //判断是否IE的Edge浏览器  
    var isIE11 = userAgent.indexOf('Trident') > -1 && userAgent.indexOf("rv:11.0") > -1;
    if (isIE) {
        var reIE = new RegExp("MSIE (\\d+\\.\\d+);");
        reIE.test(userAgent);
        var fIEVersion = parseFloat(RegExp["$1"]);
        if (fIEVersion == 7) {
            return 7;
        } else if (fIEVersion == 8) {
            return 8;
        } else if (fIEVersion == 9) {
            return 9;
        } else if (fIEVersion == 10) {
            return 10;
        } else {
            return 6; //IE版本<=7
        }
    } else if (isEdge) {
        return 20; //edge
    } else if (isIE11) {
        return 11; //IE11  
    } else {
        return 30; //不是ie浏览器
    }
}