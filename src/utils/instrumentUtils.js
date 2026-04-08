/* 单位数组
 * 涉及以下函数
 * getInputThreeData
 * getInputFiveData
 */
let unitArrays = ['MPA', 'KPA', 'MM', '℃', '°'];
let unitObj = {
    MPA: 'Mpa',
    KPA: 'Kpa',
    MM: 'mm',
    '℃': '℃',
    '°': '°'
};

/* 根据仪表量程范围返回单位
 * 0~1Mpa -> Mpa
 */
function getInputUnit(inputVal) {
    let result = '';
    unitArrays.map(item => {
        if (inputVal.toUpperCase().indexOf(item.toUpperCase()) !== -1) result = unitObj[item];
    });
    return result;
}

/* 根据仪表量程范围返回3点数值
 * 0~1Mpa -> 0 / 0.5 / 1
 */
function getInputThreeData(inputVal, fixed) {
    // 获取输入值 格式如：0~1.6Mpa 下限~上限单位
    unitArrays.map(item => {
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
    unitArrays.map(item => {
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

export { unitArrays, getInputUnit, getInputThreeData, getInputFiveData };
