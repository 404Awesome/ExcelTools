// 关于工具箱信息
function onAbout() {
    alert(`WPS(Excel)工具集\rJS / Vue / Vite / Vscode开发\r@一介俗人`);
}

// 打开GitHub连接
function onGitHub() {
    wps.TabPages.OpenWebUrl('https://github.com/404Awesome');
}

// 测试函数
function onTestFunc() {
    alert('\r暂无测试！');
}

export { onAbout, onGitHub, onTestFunc };
