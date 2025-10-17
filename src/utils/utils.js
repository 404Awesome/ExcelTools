//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft: 0,
    msoCTPDockPositionRight: 2
};

function GetUrlPath() {
    // 在本地网页的情况下获取路径
    if (window.location.protocol === 'file:') {
        const path = window.location.href;
        // 删除文件名以获取根路径
        return path.substring(0, path.lastIndexOf('/'));
    }

    // 在非本地网页的情况下获取根路径
    const { protocol, hostname, port } = window.location;
    const portPart = port ? `:${port}` : '';
    return `${protocol}//${hostname}${portPart}`;
}

function GetRouterHash() {
    if (window.location.protocol === 'file:') {
        return '';
    }

    return '/#';
}

// 打开侧边任务窗格
function OpenTaskpPane(path, Width) {
    let tsId = window.Application.PluginStorage.getItem('taskpane_id');
    if (!tsId) {
        let tskpane = window.Application.CreateTaskPane(GetUrlPath() + GetRouterHash() + path);
        tskpane.MaxWidth = Width;
        tskpane.MinWidth = Width;
        let id = tskpane.ID;
        window.Application.PluginStorage.setItem('taskpane_id', id);
        tskpane.Visible = true;
    } else {
        let tskpane = window.Application.GetTaskPane(tsId);
        tskpane.MaxWidth = Width;
        tskpane.MinWidth = Width;
        tskpane.Visible = !tskpane.Visible;
    }
}

// 关闭任务窗格
function CloseTaskPane() {
    let tsId = window.Application.PluginStorage.getItem('taskpane_id');
    if (tsId) {
        let tskpane = window.Application.GetTaskPane(tsId);
        tskpane.Visible = false;
        window.Application.PluginStorage.removeItem('taskpane_id');
    }
}

// 节流函数
function throttle(func, delay) {
    let timeout = null;
    return function (...args) {
        if (!timeout) {
            timeout = setTimeout(() => {
                func.apply(this, args);
                timeout = null; // 清除定时器
            }, delay);
        }
    };
}

// 防抖函数
function debounce(func, wait) {
    let timerId = null;
    let flag = true;
    return function () {
        clearTimeout(timerId);
        if (flag) {
            func.apply(this, arguments);
            flag = false;
        }
        timerId = setTimeout(() => {
            flag = true;
        }, wait);
    };
}

export { WPS_Enum, GetUrlPath, GetRouterHash, OpenTaskpPane, CloseTaskPane, throttle, debounce };
