import { NameSpace } from '../constant';
import { getInputUnit, getInputThreeData, getInputFiveData } from '../utils/instrumentUtils';

/* 根据仪表量程范围返回3点数值
 * 0~1Mpa -> 0 / 0.5 / 1
 */
wps?.AddCustomFunction(NameSpace, 'getInputThreeData', getInputThreeData, {
    description: '获取仪表量程范围的三点数据!',
    result: { type: 'number[]' },
    parameters: [
        { name: '量程范围', type: 'string', description: '例如0~1Mpa, 分隔符使用~' },
        { name: '小数点', type: 'number', description: '保留几位小数点!' }
    ]
});

/* 根据仪表量程范围返回5点数值
 * 0~1Mpa -> 0 / 0.25 / 0.5 / 0.75 / 1
 */
wps?.AddCustomFunction(NameSpace, 'getInputFiveData', getInputFiveData, {
    description: '获取仪表量程范围的五点数据!',
    result: { type: 'number[]' },
    parameters: [
        { name: '量程范围', type: 'string', description: '例如0~1Mpa, 分隔符使用~' },
        { name: '小数点', type: 'number', description: '保留几位小数点!' }
    ]
});

/* 根据仪表量程范围返回单位
 * 0~1Mpa -> Mpa
 */
wps?.AddCustomFunction(NameSpace, 'getInputUnit', getInputUnit, {
    description: '获取仪表量程范围中的单位!',
    result: { type: 'string' },
    parameters: [{ name: '量程范围', type: 'string', description: '例如0~1Mpa, 返回Mpa!' }]
});

/* 调节阀/执行器/开关阀调校记录
 * 根据FC/FO类型返回标准三点行程
 */
wps?.AddCustomFunction(
    NameSpace,
    'getValveStandardThreeData',
    function (inputVal, type, fixed) {
        return type.toUpperCase() === 'FC' ? getInputThreeData(inputVal, fixed) : getInputThreeData(inputVal, fixed).reverse();
    },
    {
        description: '根据FC/FO类型返回标准三点行程!',
        result: { type: 'array' },
        parameters: [
            { name: '量程范围', type: 'string', description: '例如0~1Mpa, 返回Mpa!' },
            { name: '阀门类型', type: 'string', description: '例如FO/FC!' },
            { name: '小数点位数', type: 'number', description: '例如2' }
        ]
    }
);

/* 调节阀/执行器/开关阀调校记录
 * 根据精确度及行程返回允许误差
 * ±0.1% / 0~14.3mm -> ±0.143mm
 */
wps?.AddCustomFunction(
    NameSpace,
    'getAllowableError',
    function (accuracy, itinerary, unit, fixed) {
        // 处理精确度
        let handleAccuracy = accuracy.replace('±', '').replace('%', '');
        // 处理行程
        let handleItinerary = itinerary.split('~')[1].replace(unit, '');
        // 运算允许误差
        return `${(handleItinerary * (handleAccuracy / 10)).toFixed(fixed)}${unit}`;
    },
    {
        description: '根据精确度及行程返回允许误差！',
        result: { type: 'string' },
        parameters: [
            { name: '精确度', type: 'string', description: '例如±0.1%' },
            { name: '行程', type: 'string', description: '例如0~14.3mm' },
            { name: '行程单位', type: 'string', description: '例如mm' },
            { name: '小数点', type: 'number', description: '保留几位小数点!' }
        ]
    }
);

/* 调节阀/执行器/开关阀调校记录
 * 根据实测行程返回误差
 */
wps?.AddCustomFunction(
    NameSpace,
    'getError',
    function (standardData, testedData, direction, fixed) {
        // 运算单位
        let unit = getInputUnit(standardData);

        // 获取标准行程的三点数据
        let standardThreeData = direction.toUpperCase() === 'FC' ? getInputThreeData(standardData, fixed) : getInputThreeData(standardData, fixed).reverse();

        // 误差结果
        let result = [];

        // 填入实测行程单位
        result.push(unit);

        // 运算误差
        [...standardThreeData, ...standardThreeData].map((item, index) => {
            result.push((new Number(item) - new Number(testedData[0][index])).toFixed(fixed).toString().replace('-', ''));
        });

        // 填入误差单位
        result.push(unit);

        // 运算回差
        for (let index = 0; index < 3; index++) {
            result.push(Math.max(result[index + 1], result[index + 4]).toFixed(fixed));
        }

        // 判断单位是否是 ° ，如果是则代表是开关阀，对结果数组进行处理
        if (unit === '°') result[2] = result[5] = result[9] = '';

        // 返回结果
        return result;
    },
    {
        description: '根据实测行程返回误差!',
        result: { type: 'array' },
        parameters: [
            { name: '实测行程', type: 'string', description: '例如0~14.3mm' },
            { name: '实测行程数据', type: 'array', description: '例如正:0.00/7.15/14.30' },
            { name: '阀门作用方向', type: 'string', description: '例如:FC/FO' },
            { name: '小数点', type: 'number', description: '保留几位小数点!' }
        ]
    }
);
