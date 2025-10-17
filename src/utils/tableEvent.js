import { CloseTaskPane, debounce } from './utils';

// 当激活任一工作表时触发此事件 使用防抖函数降低触发频率
wps.ApiEvent.AddApiEventListener(
    'SheetActivate',
    debounce(() => {
        // 关闭任务窗格
        CloseTaskPane();
    }, 1000)
);
// 任一工作簿窗口被激活时，将触发此事件 使用防抖函数降低触发频率
wps.ApiEvent.AddApiEventListener(
    'WindowActivate',
    debounce(() => {
        // 关闭任务窗格
        CloseTaskPane();
    }, 1000)
);
