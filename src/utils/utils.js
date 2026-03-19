//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft: 0,
    msoCTPDockPositionRight: 2,
    xlHAlignCenter: 'xlHAlignCenter',
    xlVAlignCenter: 'xlVAlignCenter'
};

// 获取Url路径
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

// 获取路由中的hash值
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

// 新建工作表
function newSheets(title, array, colWight = 10, rowHeight = 20) {
    Application.Worksheets.Add();
    Application.ActiveSheet.Name = title;
    let activeSheet = Application.Worksheets.Item(title);
    let resultAreas = activeSheet.Range('A1').Resize(array.length, array[0].length);
    resultAreas.Borders.LineStyle = 'Double';
    resultAreas.Borders.Color = 0x000000;
    resultAreas.HorizontalAlignment = Application.Enum.xlHAlignCenter;
    resultAreas.VerticalAlignment = Application.Enum.xlVAlignCenter;
    resultAreas.NumberFormat = '@';
    resultAreas.RowHeight = rowHeight;
    resultAreas.Columns.ColumnWidth = colWight;
    resultAreas.Value2 = array;
    return resultAreas;
}

/* 将一维数组分割成多维数组
 * 列如：[1, 2, 3, 4, 5, 6]
 * 生成：[
 *		  [1, 2, 3],
 *		  [4, 5, 6]
 *	    ]
 */
function chunkArray(array, number) {
    if (!array.length || !number || number < 1) return [];
    let [start, end, result] = [null, null, []];
    for (let i = 0; i < Math.ceil(array.length / number); i++) {
        start = i * number;
        end = start + number;
        result.push(array.slice(start, end));
    }
    return result;
}

/* 根据仪表量程范围返回3点数值
 * 0~1Mpa -> 0 / 0.5 / 1
 */
function getInputThreeData(inputVal, fixed) {
    // 获取输入值 格式如：0~1.6Mpa 下限~上限单位
    ['MPA', 'KPA', 'MM'].map(item => {
        if (inputVal.toString().toUpperCase().indexOf(item) !== -1) {
            inputVal = inputVal.toString().toUpperCase().replace(item, '');
        }
    })[0];

    // 分割数组
    let [botton, top] = [...inputVal.split('~')];
    botton = new Number(botton);
    top = new Number(top);

    // 三点对象
    let Three = {
        4: 0,
        12: 0.5,
        20: 1
    };

    // 计算出结果并传入数组
    let result = [];
    Object.keys(Three).map(item => {
        result.push(new Number(botton < 0 ? botton + (top - botton) * Three[item] : (top - botton) * Three[item] + botton).toFixed(fixed));
    });

    // 返回结果
    return result;
}

/* 根据仪表量程范围返回5点数值
 * 0~1Mpa -> 0 / 0.25 / 0.5 / 0.75 / 1
 */
function getInputFiveData(inputVal, fixed) {
    // 获取输入值 格式如：0~1.6Mpa 下限~上限单位
    ['MPA', 'KPA'].map(item => {
        if (inputVal.toString().toUpperCase().indexOf(item) !== -1) {
            inputVal = inputVal.toString().toUpperCase().replace(item, '');
        }
    })[0];

    // 分割数组
    let [botton, top] = [...inputVal.split('~')];
    botton = new Number(botton);
    top = new Number(top);

    // 五点对象
    let Five = {
        4: 0,
        8: 0.25,
        12: 0.5,
        16: 0.75,
        20: 1
    };

    // 计算出结果并传入数组
    let result = [];
    Object.keys(Five).map(item => {
        result.push(new Number(botton < 0 ? botton + (top - botton) * Five[item] : (top - botton) * Five[item] + botton).toFixed(fixed));
    });

    // 返回结果
    return result;
}

export { WPS_Enum, GetUrlPath, GetRouterHash, OpenTaskpPane, CloseTaskPane, throttle, debounce, newSheets, chunkArray, getInputThreeData, getInputFiveData };
