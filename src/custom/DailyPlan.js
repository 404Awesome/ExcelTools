import { OpenTaskpPane } from '../utils/utils';

// 作业周期 及 作业剩余时间
function onWorkTime() {
    try {
        // 激活工作表 导入模板
        let activeSheet = Application.Worksheets.Item('导入模板');

        // 判断第一个单元格内容是否为 企业名称 容错处理
        if (activeSheet.Range('A1').Value2 !== '企业名称') {
            throw new Error("请检查是否存在 '导入模板' 工作表!");
        }

        // 在P列生成 作业周期, 在Q列生成 剩余时间
        activeSheet.Range('P1').Value2 = '作业周期';
        activeSheet.Range('Q1').Value2 = '剩余时间';

        // 获取当前表格已使用的行数
        let useLine = Application.ActiveSheet.UsedRange.Rows.Count;

        // 作业票作业 剩余时间
        let timeRemaining = 0;

        // 遍历每条计划, 并处理
        for (let index = 2; index <= useLine + 1; index++) {
            // 判断开始时间和结束时间是否存在, 如果存在, 在P列生成作业周期, 在Q列生成剩余时间
            if (activeSheet.Range(`J${index}`).Value2 && activeSheet.Range(`K${index}`).Value2) {
                // 作业周期
                activeSheet.Range(`P${index}`).Value2 = `${(new Date(activeSheet.Range(`K${index}`).Value2) - new Date(activeSheet.Range(`J${index}`).Value2)) / 1000 / 60 / 60 / 24}天`;

                // 剩余时间
                timeRemaining = new Date(activeSheet.Range(`K${index}`).Value2).getTime() - Date.now();
                activeSheet.Range(`Q${index}`).Value2 = `${parseInt(timeRemaining / (1000 * 60 * 60 * 24))}天:${parseInt((timeRemaining % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60))}时`;

                // 如果剩余时间过期填充红色，小于一天则填充黄色提醒
                if (timeRemaining < 0) {
                    activeSheet.Range(`P${index}:Q${index}`).Interior.Color = 0x7b68ee;
                } else if (timeRemaining < 86400000) {
                    activeSheet.Range(`P${index}:Q${index}`).Interior.Color = 0x61b0fa;
                } else {
                    activeSheet.Range(`P${index}:Q${index}`).Interior.ColorIndex = 0;
                }
            }
        }

        // 为 P列 & Q列 添加边框
        useLine = Application.ActiveSheet.UsedRange.Rows.Count;
        activeSheet.Range(`P2:Q${useLine}`).Borders.LineStyle = 'Double';
        activeSheet.Range(`P2:Q${useLine}`).Borders.Color = 0x000000;
    } catch (err) {
        alert(`${err.name}\r${err.message}`);
    }
}

/* 每日计划
 * 功能如下
 * 1. 只需选择风险等级名称，即可自动生成风险等级编码
 * 2. 检查作业类型名称是否和作业级别名称对应
 * 3. 如果作业时间空白，则根据作业类型名称和作业级别生成作业时间
 * 4. 如果作业部位及内容空白，则根据作业区域名称和日计划内容描述自动生成
 * 5. 如果作业时间空白，则日计划状态为新增
 * 6. 如果作业时间存在，则日计划状态为作业中
 * 7. 若作业结束时间过期则删除此条计划
 * 8. 初始化基础信息
 * 9. 若作业类型名称和作业级别名称是否匹配，不匹配填充颜色提示
 */
function onDailyPlan() {
    try {
        // 判断当前是否打开了 导入模板 这个表
        if (Application.ActiveSheet.Name !== '导入模板') {
            alert(`请选择明日计划中的\r    《导入模板》\r    工作表后使用!`);
            return false;
        }

        /* 初始化基础信息 如下
         * 企业名称 施工单位负责人姓名 施工单位负责人联系方式  队伍名称
         */
        let CompanyName = '南化公司';
        let Name = '杲良彬';
        let Number = '18762230909';
        let TeamName = '电仪二队';

        // 激活工作表 导入模板
        let activeSheet = Application.Worksheets.Item('导入模板');

        // 判断第一个单元格内容是否为 企业名称 容错处理
        if (activeSheet.Range('A1').Value2 !== '企业名称') {
            throw new Error("请检查是否存在 '导入模板' 工作表!");
        }

        // 是否生成 每日计划附加文本
        let IsOnDailyPlanText = true;

        // 天数对应的毫秒数;
        let days = {
            // 一天毫秒数
            oneDay: 86400000,
            // 三天毫秒数
            threeDay: 259200000,
            // 五天毫秒数
            fiveDay: 432000000,
            // 七天毫秒数
            sevenDay: 604800000,
            // 十天毫秒数
            tenDay: 864000000,
            // 十五天毫秒数
            fifteenDay: 1.296e9,
            // 三十天毫秒数
            thirtyDay: 2.592e9
        };

        // 风险等级名称 对应 风险等级编码
        let RiskLevel = {
            低风险: '11316004',
            重大风险: '11316001',
            较大风险: '11316002',
            一般风险: '11316003'
        };

        // 过期计划索引
        let PlanExpired = [];

        // 返回单元格引用
        function GetValue2(range) {
            return activeSheet.Range(range);
        }

        // 日期单数补0函数
        function add0(date) {
            return date.toString().length === 1 ? `0${date}` : date;
        }

        // 生成带格式的日期
        function createDate(milliseconds, SpecifiedDate) {
            let date = new Date(new Date(SpecifiedDate).getTime() + milliseconds);
            return `${date.getFullYear()}-${add0(date.getMonth() + 1)}-${add0(date.getDate())} 09:00:00`;
        }

        // 单元格填写错误 填充颜色警示
        function errorTip(range) {
            // 填充红色
            activeSheet.Range(range).Interior.Color = 0x7b68ee;

            // 将 每日计划附加文本（IsOnDailyPlanText）设置为不生成 false
            IsOnDailyPlanText = false;
        }

        // 生成明天作业开始时间
        let tomorrowDate = createDate(days.oneDay, Date.now());

        // 获取当前表格已使用的行数
        let useLine = Application.ActiveSheet.UsedRange.Rows.Count - 1;

        // 遍历每行计划
        for (let index = 2; index <= useLine + 1; index++) {
            // 存在计划的行 --> index
            /* 列号对应内容 如下
             * 企业名称			->	A
             * 日计划内容描述		->	B
             * 风险等级名称		->  C
             * 风险等级编码		->  D
             * 作业区域名称		->  E
             * 作业类型名称		->  F
             * 作业级别名称		->  G
             * 作业部位及内容		->  H
             * 作业人数			->  I
             * 作业开始时间		->  J
             * 作业结束时间		->  K
             * 施工单位负责人姓名		->  L
             * 施工单位负责人联系方式 ->  M
             * 日计划状态			->  N
             */

            // 初始化基础信息 如顶部参数所示
            GetValue2(`A${index}`).Value2 = CompanyName;
            GetValue2(`L${index}`).Value2 = Name;
            GetValue2(`M${index}`).Value2 = Number;

            // 初始化单元格填充颜色为白色
            activeSheet.Range(`A${index}:N${index}`).Interior.ColorIndex = 0;

            // 判断 风险等级名称 / 作业区域名称 / 作业类型名称 / 作业级别名称 / 作业人数 是否填写
            let empyt = false;
            for (let item of ['B', 'C', 'E', 'F', 'G', 'I']) {
                if (GetValue2(`${item}${index}`).Value2 === null && !'一般作业（项目部）临时用电（项目部）受限空间（项目部）'.includes(GetValue2(`F${index}`).Value2)) {
                    // 未填写则为单元格填充红色
                    errorTip(`${item}${index}`);
                    // 存在未填写的单元格，记录状态，跳过本行计划
                    empyt = true;
                }
            }
            if (empyt) continue;

            /* 生成 <作业部位及内容> 并填入
             * 如果 <作业区域名称> 后一个字是数字, 添加 "装置" 二字
             */
            GetValue2(`H${index}`).Value2 = `中化六建${GetValue2(`I${index}`).Value2}人` + GetValue2(`E${index}`).Value2 + (!!parseInt(GetValue2(`E${index}`).Value2.slice(-1)) ? '装置' : '') + GetValue2(`B${index}`).Value2;

            // 根据 风险等级名称 生成 风险等级编码 并填入
            GetValue2(`D${index}`).Value2 = RiskLevel[GetValue2(`C${index}`).Value2];

            // 处理作业开始时间
            if (GetValue2(`J${index}`).Value2 === null) {
                GetValue2(`J${index}`).Value2 = tomorrowDate;
            }

            // 处理作业结束时间
            switch (GetValue2(`F${index}`).Value2) {
                case '一般作业（项目部）':
                    // 作业级别名称 无级别
                    GetValue2(`G${index}`).Value2 = '';
                    // 作业周期为10天
                    GetValue2(`K${index}`).Value2 = createDate(days.tenDay, GetValue2(`J${index}`));
                    break;
                case '动土作业（项目部）':
                    /* 作业周期如下
                     * 深度<3m --> 30天
                     * 深度>=3m --> 15天
                     */
                    if (GetValue2(`G${index}`).Value2 === '3m以下') {
                        GetValue2(`K${index}`).Value2 = createDate(days.thirtyDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '3m以上') {
                        GetValue2(`K${index}`).Value2 = createDate(days.fifteenDay, GetValue2(`J${index}`).Value2);
                    } else {
                        // 作业级别名称填写错误
                        errorTip(`G${index}`);
                        GetValue2(`J${index}`).Value2 = '';
                        GetValue2(`K${index}`).Value2 = '';
                        GetValue2(`N${index}`).Value2 = '';
                        continue;
                    }
                    break;
                case '临时用电（项目部）':
                    // 作业级别名称 无级别
                    GetValue2(`G${index}`).Value2 = '';
                    // 作业周期为15天
                    GetValue2(`K${index}`).Value2 = createDate(days.fifteenDay, GetValue2(`J${index}`).Value2);
                    break;
                case '起重吊装（项目部）':
                    /* 作业周期如下
                     * 一级 --> 1天
                     * 二级 --> 3天
                     * 三级 --> 7天
                     */
                    if (GetValue2(`G${index}`).Value2 === '一级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.oneDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '二级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.threeDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`) === '三级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.sevenDay, GetValue2(`J${index}`).Value2);
                    } else {
                        // 作业级别名称填写错误
                        errorTip(`G${index}`);
                        GetValue2(`J${index}`).Value2 = '';
                        GetValue2(`K${index}`).Value2 = '';
                        GetValue2(`N${index}`).Value2 = '';
                        continue;
                    }
                    break;
                case '受限空间（项目部）':
                    // 作业级别名称 无级别
                    GetValue2(`G${index}`).Value2 = '';
                    // 作业周期为1天
                    GetValue2(`K${index}`).Value2 = createDate(days.oneDay, GetValue2(`J${index}`).Value2);
                    break;
                case '高处作业（项目部）':
                    /* 作业周期如下
                     * 一级 --> 7天
                     * 二级 --> 7天
                     * 三级 --> 5天
                     * 四级 --> 3天
                     */
                    if (GetValue2(`G${index}`).Value2 === '一级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.sevenDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '二级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.sevenDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '三级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.fiveDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '四级') {
                        GetValue2(`K${index}`).Value2 = createDate(days.threeDay, GetValue2(`J${index}`).Value2);
                    } else {
                        // 作业级别名称填写错误
                        errorTip(`G${index}`);
                        GetValue2(`J${index}`).Value2 = '';
                        GetValue2(`K${index}`).Value2 = '';
                        GetValue2(`N${index}`).Value2 = '';
                        continue;
                    }
                    break;
                case '动火作业（项目部）':
                    /* 作业周期如下
                     * 施工动火区 --> 7天
                     * 固定动火区 --> 30天
                     */
                    if (GetValue2(`G${index}`).Value2 === '施工动火区') {
                        GetValue2(`K${index}`).Value2 = createDate(days.sevenDay, GetValue2(`J${index}`).Value2);
                    } else if (GetValue2(`G${index}`).Value2 === '固定动火区') {
                        GetValue2(`K${index}`).Value2 = createDate(days.thirtyDay, GetValue2(`J${index}`).Value2);
                    } else {
                        // 作业级别名称填写错误
                        errorTip(`G${index}`);
                        GetValue2(`J${index}`).Value2 = '';
                        GetValue2(`K${index}`).Value2 = '';
                        GetValue2(`N${index}`).Value2 = '';
                        continue;
                    }
                    break;
                case '射线作业（项目部）':
                    /* 作业周期如下
                     * X射线 --> 天
                     * γ射线 --> 天
                     */
                    if (GetValue2(`G${index}`).Value2 === 'X射线') {
                    } else if (GetValue2(`G${index}`).Value2 === 'γ射线') {
                    } else {
                        // 作业级别名称填写错误
                        errorTip(`G${index}`);
                        GetValue2(`J${index}`).Value2 = '';
                        GetValue2(`K${index}`).Value2 = '';
                        GetValue2(`N${index}`).Value2 = '';
                        continue;
                    }
                    alert('射线作业（项目部）未开发! 自主填写作业时间!');
                    break;
            }

            // 计划中：查看结束时间是否到期，如果到期，标记行号至 PlanExpired
            if (Date.parse(new Date(GetValue2(`K${index}`).Value2)) < Date.now()) {
                PlanExpired.push(index);
            }

            // 判断 当前时间 是否大于 作业开始时间，大于则 日计划状态 为 作业中，否则 为 新增
            if (Date.parse(new Date(GetValue2(`J${index}`).Value2)) - 32400000 < Date.now()) {
                GetValue2(`N${index}`).Value2 = '作业中';
            } else {
                GetValue2(`N${index}`).Value2 = '新增';
            }
        }

        // 过期计划行标红处理
        PlanExpired.reverse().map(index => {
            GetValue2(`A${index}:N${index}`).Interior.Color = 0x7b68ee;
        });

        // 重新获取已使用的行并对已存在数据添加表格边框
        useLine = Application.ActiveSheet.UsedRange.Rows.Count;
        activeSheet.Range(`A2:N${useLine}`).Borders.LineStyle = 'Double';
        activeSheet.Range(`A2:N${useLine}`).Borders.Color = 0x000000;

        // 刷新时间周期
        if (IsOnDailyPlanText) {
            onWorkTime();
        }

        // 打开 每日计划附加文本 任务窗格
        if (IsOnDailyPlanText) {
            OpenTaskpPane(`/dailyplan?tomorrowDate=${tomorrowDate}&TeamName=${TeamName}`, 400);
        } else {
            alert('\r表格中有必填项未填写！');
        }
    } catch (err) {
        alert(`${err.name}\r${err.message}`);
    }
}

export { onDailyPlan };
