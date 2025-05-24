# WPS加载项开发实战之一安装、部署与调试

#### 介绍
WPS 加载项（简称WPSJS）是一套基于 Web 技术用来扩展 WPS 应用程序的解决方案，开发完成一个WPSJS加载项后首先要面对的是用户怎么安装和使用，为了让用户使用WPSJ就必须发布到一个WEB服务器上，在这个项目中就是实现一个加载项的安装。以WPS文档助手加载项为例，这个加载项的具体实现在第二部分展现。为了便于掌握wpsjs加载项的实现机制所以没有使用wpsjs工具包创建而是从0创建，为了便于理解把功能尽可能简化。

#### 开发环境准备
开发WPSJS需要熟悉WEB和WPS客户端两种技术。基本配置是一个静态网页服务器和WPS软件。静态网页服务使用gitee.com的pages服务功能，所以对应的代码也存储在了gitee.com仓库中。
 [打开代码仓库https://gitee.com/wpsoa/wpsoa](https://gitee.com/wpsoa/wpsoa) 在这儿可以看所有的代码，[访问https://wpsoa.gitee.io](https://wpsoa.gitee.io) 安装或卸载加载项。编程和调试还需一套开发环境这个可根据个人喜好，从gitee.com拉取或推送代码，最好选用支持git命令的集成开发环境如VSCODE或VSSTUDIO等。WPS客户提供了非常多的接口来操作对应的文档，wpsjs提供了几个.d.ts扩展名的文件，这是TypeScript的声明文件，在这里面基本上把所有的接口作了描述，使用Typescript编写代码时这些接口会出现在编辑提示中，而且还可以在编译成javascript之前发现错误，提高编程效率，因些建议使用Typescript语言。安装与管理的加载项是仅实现在WPS表格的视图选项卡中显示一个文档助手按钮[加载项代码仓库https://gitee.com/wpsoa/xlsapp](https://gitee.com/wpsoa/xlsapp) ,后续在其中一步步实现具体功能。

#### 文件目录

1.  index.html            用户访问首页，进行加载项的安装或卸载操作
2.  js/wpsjsrpcsdk.js     WPS封装好的wpsjs管理工具（与官方提供的相比做了修改解决了https支持的问题）

#### 理解管理加载项的实现原理

在用户的电脑上安装过WPS软件后，WPS会在用户的电脑上安装WEB服务（就相当于一个小的WEB服务器），这个服务器的地址是`http://127.0.0.1:58890`和`https://127.0.0.1:58891`，wpsjsrpcsdk.js中就是用js的发送http请求到上述两个地址来管理加载项或启动WPS客户端。具体用法参加WPS开放平台。
#### index.html 代码分析
以下内容是index.html中的定义了要管理的加载项的信息
```
 window.element = { "name": "wpsapp", "addonType": "et", "online": "true", "multiUser": "false", "url": "https://wpsoa.gitee.io/xlsapp/"};  
```
点击安装按钮后执行加载项安装，在这里没使用检测是否安装而是先卸载后再重新安装
```
document.getElementById("setup").onclick = function () { WpsAddonMgr.disable(element); WpsAddonMgr.enable(element); };
```
单进程和多进程启动，其中openBook是启动WPS客户端后执行的加载项中的方法，执行这个方法需要先安装加载项，在下一步的加载项开发时实现。

```
document.getElementById("start").onclick = function () {
                WpsInvoke.InvokeAsHttps(WpsInvoke.ClientType.et,
                    "wpsapp",
                    "openBook",
                    JSON.stringify(""), (e) => {console.log(e) }, true, "https://fxzqf.github.io/wpsapp/jsplugins.xml", false);
            };
            document.getElementById("client").onclick = function () {
                var wpsClient = new WpsClient(WpsInvoke.ClientType.et);
                wpsClient.jsPluginsXml = "https://fxzqf.gitee.io/wpsapp/jsplugins.xml";
                wpsClient.InvokeAsHttps("wpsapp", "openBook", JSON.stringify(""),(e)=>{console.log(e)});
            }

```
#### 使用和调试
在安装了WPS客户端的电脑上[访问https://wpsoa.gitee.io](https://wpsoa.gitee.io)点击安装按钮，弹出用户确认对话框后安装加载项到本机，打开WPS客户端新建或打开一个WPS表格，你会看到在视图选项卡的最后多了一个“文档助手”按钮，该按钮还没有任何功能，在下一步的开发中实现。访问本机文件路径可以在C:\Users\{用户名如administrator}\AppData\Roaming\kingsoft\wps\jsaddons路径下安装的wps加载项信息，删除所有文件就可以卸载所有加载项了。

1、启用调WPSJS调试功能

 打开 WPS 安装目录，进入 wps.exe 程序同级目录，在cfgs文件夹下新建或打开oem.ini文件，在文档最后添加如下配置
 ```
 [Support]
##启用加载项
JsApiPlugin = true
##启用网页调试，详情请参考 WPS 基础接口说明
JsApiShowWebDebugger=true
[Server]
##指定加载项配置文件 jsplugins.xml 链接地址
JSPluginsServer = http://***/jsplugins.xml
```
JsApiShowWebDebugger=true是启用调试

2、调试

重新启动WPS客户端，在视图选项卡中的新出现"打开JS调试器"按钮，点击该按钮打开调试窗口进行调试。