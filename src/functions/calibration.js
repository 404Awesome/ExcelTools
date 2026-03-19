import { NameSpace } from '../constant';
import { getInputThreeData } from '../utils/utils';

/* 根据仪表量程范围返回5点数值
 * 0~1Mpa -> 0 / 0.25 / 0.5 / 0.75 / 1
 */
wps?.AddCustomFunction(
    NameSpace,
    'getInputFiveData',
    function (inputVal, fixed) {
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
    },
    {
        description: '获取仪表量程范围的五点数据!',
        result: { type: 'number[]' },
        parameters: [
            { name: '量程范围', type: 'string', description: '例如0~1Mpa, 分隔符使用~' },
            { name: '小数点', type: 'number', description: '保留几位小数点!' }
        ]
    }
);

/* 根据仪表量程范围返回3点数值
 * 0~1Mpa -> 0 / 0.5 / 1
 */
wps?.AddCustomFunction(
    NameSpace,
    'getInputThreeData',
    function (inputVal, fixed) {
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
    },
    {
        description: '获取仪表量程范围的三点数据!',
        result: { type: 'number[]' },
        parameters: [
            { name: '量程范围', type: 'string', description: '例如0~1Mpa, 分隔符使用~' },
            { name: '小数点', type: 'number', description: '保留几位小数点!' }
        ]
    }
);

/* 根据仪表量程范围返回单位
 * 0~1Mpa -> Mpa
 */
wps?.AddCustomFunction(
    NameSpace,
    'getInputUnit',
    function (inputVal) {
        let result = '';
        ['Mpa', 'Kpa'].map(item => {
            if (inputVal.toUpperCase().indexOf(item.toUpperCase()) !== -1) result = item;
        });
        return result;
    },
    {
        description: '获取仪表量程范围中的单位!',
        result: { type: 'string' },
        parameters: [{ name: '量程范围', type: 'string', description: '例如0~1Mpa, 返回Mpa!' }]
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
    function (standardData, testedData) {
        // 获取标准行程的三点数据 正1/正2/正3/反1/反2/反3
        let standardThreeData = [...getInputThreeData(standardData, 2), ...getInputThreeData(standardData, 2).reverse()];

        // 获取实测行程的三点数据 正1/正2/正3/反1/反2/反3
        let testedThreeData = testedData[0];

        // 运算单位
        let unit = '';
        ['MM'].map(item => {
            if (standardData.toUpperCase().indexOf(item) > 0) {
                unit = item.toLocaleLowerCase();
            }
        });

        // 误差结果
        let result = [];

        // 填入实测行程单位
        result.push(unit);

        // 运算误差
        standardThreeData.map((item, index) => {
            result.push((new Number(item) - new Number(testedThreeData[index])).toFixed(2));
        });

        // 填入误差单位
        result.push(unit);

        // 运算回差
        for (let index = 0; index < 3; index++) {
            result.push(Math.max(result[index + 2], result[index + 4]).toFixed(2));
        }

        // 返回结果
        return result;
    },
    {
        description: '根据实测行程返回误差!',
        result: { type: 'string' },
        parameters: [
            { name: '实测行程', type: 'string', description: '例如0~14.3mm' },
            { name: '实测行程数据', type: 'array', description: '例如正:0.00/7.15/14.30' }
        ]
    }
);
