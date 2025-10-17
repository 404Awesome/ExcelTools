import { NameSpace } from '../constant';

// 根据身份证号返回年龄
wps?.AddCustomFunction(
    NameSpace,
    'cardIDForAge',
    function (id) {
        // 判断身份证号长度是否是18位
        id = id.toString();
        if (id.length === 18) {
            //处理18位的身份证号码从号码中得到生日和性别代码
            let strBirthday = id.substr(6, 4) + '/' + id.substr(10, 2) + '/' + id.substr(12, 2);

            // 时间字符串里，必须是“/”
            var birthDate = new Date(strBirthday);
            var nowDateTime = new Date();
            var age = nowDateTime.getFullYear() - birthDate.getFullYear();
            if (nowDateTime.getMonth() < birthDate.getMonth() || (nowDateTime.getMonth() == birthDate.getMonth() && nowDateTime.getDate() < birthDate.getDate())) {
                age--;
            }
            return age.toString();
        } else {
            return '身份证号长度不符合！';
        }
    },
    {
        description: '选择或输入身份证号，返回当前年龄！',
        result: { type: 'string' },
        parameters: [{ name: '身份证号', type: 'string', description: '选择身份证号！' }]
    }
);
