import { NameSpace } from '../utils/utils';

// 计算实名制考勤的实际上班时间 每9小时为1个工，剩余时间为加班
wps?.AddCustomFunction(
    NameSpace,
    'getZHLJWorkTime',
    function (timeStr) {
        /**
         * 打卡时间计算工时函数
         *
         * 函数说明:
         * 根据员工打卡时间字符串，计算实际工作时间
         *
         * 输入参数:
         * - timeStr: 打卡时间字符串，格式如 "06:11,11:03,12:04,17:32,18:30,21:00"
         *   支持2次以上任意次数打卡，用逗号/换行分隔
         *   支持中英文逗号、换行符
         *   支持 HH:mm 或 HH:mm:ss 格式
         *
         * 输出格式(换行分隔):
         * - 第1行: 天数(00-01)
         * - 第2行: 小时(00-23)
         * - 第3行: 分钟(00-59)
         * - 第4行: 打卡次数(如06次打卡)
         *
         * 计算规则:
         * - 两两配对计算时长：第1-2次、第3-4次、第5-6次...（上班-下班配对）
         * - 奇数次打卡：最后一次与倒数第二次配对（视为加班下班打卡）
         * - 9小时=1天，剩余为加班时间
         * - 自动按时间排序，兼容乱序输入
         *
         * 示例:
         * - "08:00,12:00,13:00,17:00"      -> 4次打卡(上午+下午) = 8小时
         * - "08:00,12:00,13:00,18:00,19:00,22:00" -> 6次打卡(含加班) = 11小时
         * - "08:00,20:00"                  -> 2次打卡 = 12小时
         * - "08:00,12:00,18:00"            -> 3次打卡 = 10小时(最后两次配对)
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
        var timeParts = inputStr.split(/[,，\n\r]+/);

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

        // 计算总工时(分钟) - 支持任意偶数次打卡(2/4/6/8...)，奇数次按最后一次为下班时间处理
        var totalMinutes = 0;
        if (validCount >= 2) {
            // 两两配对计算时长：第1-2次、第3-4次、第5-6次...
            // 奇数次打卡时，最后一次视为下班时间，与倒数第2次配对
            var pairCount = Math.floor(validCount / 2);
            for (var i = 0; i < pairCount; i++) {
                var startIdx = i * 2;
                var endIdx = startIdx + 1;
                var startMinutes = parsedTimes[startIdx].hours * 60 + parsedTimes[startIdx].minutes;
                var endMinutes = parsedTimes[endIdx].hours * 60 + parsedTimes[endIdx].minutes;
                var duration = endMinutes - startMinutes;
                if (duration > 0) {
                    totalMinutes += duration;
                }
            }
            
            // 处理奇数次打卡：最后一次与倒数第二次配对（视为加班下班）
            if (validCount % 2 === 1 && validCount >= 3) {
                var lastStartIdx = validCount - 2;
                var lastEndIdx = validCount - 1;
                var startMinutes = parsedTimes[lastStartIdx].hours * 60 + parsedTimes[lastStartIdx].minutes;
                var endMinutes = parsedTimes[lastEndIdx].hours * 60 + parsedTimes[lastEndIdx].minutes;
                var duration = endMinutes - startMinutes;
                if (duration > 0) {
                    totalMinutes += duration;
                }
            }
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
        description: '根据打卡时间计算实际工作时间(9小时=1天)，支持2/4/6/8...次打卡及奇数次打卡',
        result: { type: 'string' },
        parameters: [
            {
                name: 'timeStr',
                type: 'string',
                description: '打卡时间字符串，如06:11,11:03,12:04,17:32,18:30,21:00(支持6次打卡含加班)'
            }
        ]
    }
);
