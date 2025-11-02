import { NameSpace } from '../constant';

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
        description: '获取仪表量程范围的五点数据！',
        result: { type: 'number[]' },
        parameters: [
            { name: '量程范围', type: 'string', description: '例如0~1Mpa, 分隔符使用~' },
            { name: '小数点', type: 'number', description: '保留几位小数点！' }
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
        description: '获取仪表量程范围中的单位！',
        result: { type: 'string' },
        parameters: [{ name: '量程范围', type: 'string', description: '例如0~1Mpa, 返回Mpa！' }]
    }
);
