// 系统工具
import System from './system.js';
// 自定义工具
import { WPS_Enum } from './utils.js';
// 公共自定义函数
import { onAbout, onGitHub, onTestFunc } from '../custom/PublicFunction.js';
// 每日计划自定义函数
import { onDailyPlan } from '../custom/DailyPlan.js';

//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
    if (typeof window.Application.ribbonUI != 'object') {
        window.Application.ribbonUI = ribbonUI;
    }

    if (typeof window.Application.Enum != 'object') {
        // 如果没有内置枚举值
        window.Application.Enum = WPS_Enum;
    }

    //这几个导出函数是给外部业务系统调用的
    window.openOfficeFileFromSystemDemo = System.openOfficeFileFromSystemDemo;
    window.InvokeFromSystemDemo = System.InvokeFromSystemDemo;
    return true;
}

// 按钮点击事件
function OnAction(control) {
    const eleId = control.Id;
    switch (eleId) {
        // About按钮
        case 'About':
            onAbout();
            break;
        // GitHub按钮
        case 'GitHub':
            onGitHub();
            break;
        // 每日计划按钮
        case 'MRJH':
            onDailyPlan();
            break;
        // 测试按钮
        case 'TEST':
            onTestFunc();
            break;
        default:
            break;
    }
    return true;
}

// 返回按钮图片
function GetImage(control) {
    const eleId = control.Id;
    switch (eleId) {
        case 'About':
            return 'images/about.svg';
        case 'GitHub':
            return 'images/link.svg';
        case 'MRJH':
            return 'images/list.svg';
        case 'TEST':
            return 'images/test.svg';
        default:
    }
    return 'images/default.svg';
}

// 是否可用
function OnGetEnabled(control) {
    const eleId = control.Id;
    switch (eleId) {
        case '':
            return false;
        default:
            break;
    }
    return true;
}

// 是否可见
function OnGetVisible(control) {
    const eleId = control.Id;
    console.log(eleId);
    return true;
}

// 动态label
function OnGetLabel(control) {
    const eleId = control.Id;
    switch (eleId) {
        case '':
            {
            }
            break;
    }
    return '';
}

//这些函数是给wps客户端调用的
export default {
    OnAddinLoad,
    OnAction,
    GetImage,
    OnGetEnabled,
    OnGetVisible,
    OnGetLabel
};
