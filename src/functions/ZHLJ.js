import { NameSpace } from '../constant';

// 计算实名制考勤的实际上班时间 每9小时为1个工，剩余时间为加班
wps?.AddCustomFunction(
    NameSpace,
    'workTime',
    function (timeStr) {
        /**
         * 打卡时间计算工时函数
         *
         * 函数说明:
         * 根据员工打卡时间字符串，计算实际工作时间
         *
         * 输入参数:
         * - timeStr: 打卡时间字符串，格式如 "06:11,11:03,12:04,17:32"
         *   支持2-4次打卡，用逗号分隔
         *   支持中英文逗号
         *
         * 输出格式(换行分隔):
         * - 第1行: 天数(00-01)
         * - 第2行: 小时(00-23)
         * - 第3行: 分钟(00-59)
         * - 第4行: 打卡次数(如04次打卡)
         *
         * 计算规则:
         * - 4次打卡: 1-2为上午时段，3-4为下午时段
         * - 3次打卡: 1-2为上午，第3次为下午下班
         * - 2次打卡: 1为上班，2为下班
         * - 9小时=1天，剩余为加班时间
         *
         * 参考标准:
         * - 按9小时计算为1个工作日
         * - 剩余时间按实际加班计算
         */

        // 参数处理
        var inputStr = '';
        if (timeStr !== null && timeStr !== undefined) {
            inputStr = String(timeStr).trim();
        }

        // 空输入检查
        if (!inputStr) {
            return '请输入打卡时间';
        }

        // 分割打卡时间
        var timeParts = inputStr.split(/[,，]/);

        // 解析有效打卡次数
        var validCount = 0;
        var parsedTimes = [];
        for (var i = 0; i < timeParts.length; i++) {
            var timeStrItem = timeParts[i].trim();
            if (!timeStrItem) continue;

            // 解析时间格式 HH:mm 或 HH:mm:ss
            var timeParts2 = timeStrItem.split(':');
            if (timeParts2.length < 2 || timeParts2.length > 3) continue;

            var hour = parseInt(timeParts2[0]);
            var minute = parseInt(timeParts2[1]);

            if (isNaN(hour) || isNaN(minute) || hour > 23 || minute > 59) continue;

            validCount++;
            parsedTimes.push({ hours: hour, minutes: minute });
        }

        // 打卡次数检查
        if (validCount < 2) {
            return '请至少输入2个有效打卡时间';
        }

        // 按时间排序
        parsedTimes.sort(function (a, b) {
            return a.hours * 60 + a.minutes - (b.hours * 60 + b.minutes);
        });

        // 计算总工时(分钟)
        var totalMinutes = 0;
        if (validCount === 4) {
            var morningMinutes = parsedTimes[1].hours * 60 + parsedTimes[1].minutes - (parsedTimes[0].hours * 60 + parsedTimes[0].minutes);
            var afternoonMinutes = parsedTimes[3].hours * 60 + parsedTimes[3].minutes - (parsedTimes[2].hours * 60 + parsedTimes[2].minutes);
            totalMinutes = morningMinutes + afternoonMinutes;
        } else if (validCount === 3) {
            var morningMinutes = parsedTimes[1].hours * 60 + parsedTimes[1].minutes - (parsedTimes[0].hours * 60 + parsedTimes[0].minutes);
            var afternoonMinutes = parsedTimes[2].hours * 60 + parsedTimes[2].minutes - (parsedTimes[1].hours * 60 + parsedTimes[1].minutes);
            totalMinutes = morningMinutes + afternoonMinutes;
        } else if (validCount === 2) {
            totalMinutes = parsedTimes[1].hours * 60 + parsedTimes[1].minutes - (parsedTimes[0].hours * 60 + parsedTimes[0].minutes);
        }

        // 计算天数和剩余时间(9小时=1天)
        var workDayMinutes = 9 * 60;
        var days = totalMinutes >= workDayMinutes ? 1 : 0;
        var remainingMinutes = totalMinutes >= workDayMinutes ? totalMinutes - workDayMinutes : totalMinutes;
        var hours = Math.floor(remainingMinutes / 60);
        var minutes = Math.round(remainingMinutes % 60);

        // 构建输出字符串(换行分隔)
        var resultStr = '';

        // 天数
        resultStr += `天: ${days}\n`;

        // 小时
        resultStr += `时: ${hours}\n`;

        // 分钟
        if (minutes < 10) minutes = ` ${minutes}`;
        resultStr += `分:${minutes}\n`;

        // 打卡次数
        resultStr += `次: ${validCount}`;

        return resultStr;
    },
    {
        description: '根据打卡时间计算实际工作时间(9小时=1天)',
        result: { type: 'string' },
        parameters: [
            {
                name: 'timeStr',
                type: 'string',
                description: '打卡时间字符串，如06:11,11:03,12:04,17:32'
            }
        ]
    }
);
