import { OpenTaskpPane } from '../utils/utils';

// 关于工具箱信息
function onAbout() {
    OpenTaskpPane('/', 400);
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
