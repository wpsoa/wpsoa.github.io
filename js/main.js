"use strict";
/// <reference path=".././node_modules/wps-jsapi/src/index.d.ts" />
//这个文件由index.html包含
var WPS_Enum = {
    msoCTPDockPositionLeft: 0,
    msoCTPDockPositionRight: 2
};
var taskPanes = new Array();
/**
 * 当打开新的工作簿时创新与工作簿对应的操作窗格
 * @param wb
 * @returns
 */
function onWorkbookOpen(wb) {
    var obj = wb.CustomDocumentProperties;
    var name = wb.Name;
    for (var x = obj.Count; x > 0; x--) {
        if (obj.Item(x).Name == "TaskPane") {
            wps.ActiveTaskPane = wps.CreateTaskPane("https://fxzqf.github.io/" + obj.Item(x).Value, "表格助手");
            taskPanes.push({ Name: name, ID: wps.ActiveTaskPane.ID });
            wps.ActiveTaskPane.Visible = true;
            return;
        }
    }
}
/**
 * 当有工作簿关闭时删除工作簿对应的操作窗格
 * @param wb
 */
function onWorkbookBeforeClose(wb) {
    var name = wb.Name;
    var x = -1;
    for (var _i = 0, taskPanes_1 = taskPanes; _i < taskPanes_1.length; _i++) {
        var taskPane = taskPanes_1[_i];
        x++;
        if (taskPane.Name == name) {
            wps.GetTaskPane(taskPane.ID).Delete();
            break;
        }
    }
    if (x >= 0)
        taskPanes.splice(x, 1);
}
/**
 * 当窗口激活时显示工作簿对应的操作窗格
 * @param wb
 * @param win
 * @returns
 */
function onWindowActivate(wb, win) {
    var name = wb.Name;
    for (var _i = 0, taskPanes_2 = taskPanes; _i < taskPanes_2.length; _i++) {
        var taskPane = taskPanes_2[_i];
        if (taskPane.Name == name) {
            (wps.ActiveTaskPane = wps.GetTaskPane(taskPane.ID)).Visible = true;
            return;
        }
    }
}
function GetUrlPath() {
    var e = document.location.toString();
    return -1 != (e = decodeURI(e)).indexOf("/") && (e = e.substring(0, e.lastIndexOf("/"))), e;
}
/**
 * 通过wps提供的接口执行一段脚本
 * @param {*} param 需要执行的脚本
 */
function shellExecuteByOAAssist(param) {
    if (wps != null) {
        wps.OAAssist.ShellExecute(param, "");
    }
}
//wps.ribbonUI.InvalidateControl("btnShowTaskPane")
//wps.ribbonUI.Invalidate(); //这行代码打开则是刷新所有的按钮状态
//alert(typeof (Workbook));
//wps.ActiveTaskPane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html", "表格助手")
//wps.ActiveTaskPane.Visible=true;
//const eleId = control.Id;
//let tsId: any = wps.PluginStorage.getItem("taskpane_id")
//if (!tsId) {
//    let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html", "表格助手")
//    let id = tskpane.ID
//    wps.PluginStorage.setItem("taskpane_id", id)
//    tskpane.Visible = true
//} else {
//    let tskpane = wps.GetTaskPane(tsId)
//    tskpane.Visible = !tskpane.Visible
//}
//document.write("<script language='javascript' src='js/util.js'></script>");
//document.write("<script language='javascript' src='js/ribbon.js'></script>");
//document.write("<script language='javascript' src='js/systemdemo.js'></script>");
//alert(wps.EtApplication().ActiveWorkbook.Name);
//alert(wps.WPSCloudService.UserInfo);
//wps.WPSCloudService.Login();
/**
 * 当应用程序窗口变成非活动窗口，关闭当前活动操作窗格
 * @param wb
 * @param win
 */
function onWindowDeactivate(wb, win) {
    if (wps.ActiveTaskPane) {
        wps.ActiveTaskPane.Visible = false;
        wps.ActiveTaskPane = null;
    }
}
/**
 * 获取一个控件的图标
 * @param control 要获取图标的控件
 * @returns 图标的SVG图像的URL
 */
function GetImage(control) {
    var eleId = control.Id;
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg";
        case "btnShowDialog":
            return "images/2.svg";
        case "btnShowTaskPane":
            return "images/3.svg";
        default:
            ;
    }
    return "images/newFromTemp.svg";
}
/**
 * 获一个控件的可用状态
 * @param control 要获取状态的控件
 * @returns true 可用,false 不可用
 */
function OnGetEnabled(control) {
    return true;
}
function OnGetVisible(control) {
    return true;
}
/**
 * 表格助手按钮单击显示表格助手任务窗格
 * @param control
 * @returns
 */
function OnAction(control) {
    if (wps.ActiveTaskPane)
        wps.ActiveTaskPane.Visible = true;
    return true;
}
window.onload = function () {
    if (wps.Enum)
        wps.Enum = WPS_Enum; // 如果没有内置枚举值
    if (wps.Application)
        wps.Application = wps.EtApplication();
    wps.ApiEvent.AddApiEventListener("WindowActivate", onWindowActivate);
    wps.ApiEvent.AddApiEventListener("WindowDeactivate", onWindowDeactivate);
    wps.ApiEvent.AddApiEventListener("WorkbookBeforeClose", onWorkbookBeforeClose);
    wps.ApiEvent.AddApiEventListener("WorkbookOpen", onWorkbookOpen);
};
//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
    if (wps.RibbonUI)
        wps.RibbonUI = ribbonUI;
    if (wps.Application.ActiveWorkbook)
        onWorkbookOpen(wps.Application.ActiveWorkbook);
    return true;
}
function openOfficeFileFromSystemDemo(param) {
    var jsonObj = (typeof (param) == 'string' ? JSON.parse(param) : param);
    alert("从业务系统传过来的参数为：" + JSON.stringify(jsonObj));
    return { wps加载项项返回: jsonObj.filepath + ", 这个地址给的不正确" };
}
function newWorkbook() {
    wps.Application.Workbooks.Add(undefined);
}
