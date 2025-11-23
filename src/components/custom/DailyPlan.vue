<template>
    <div id="DailyPlan">
        <!-- 基本信息 -->
        <div class="info">
            <!-- 标题 -->
            <n-gradient-text style="text-align: center" :size="24" type="success">{{ new Date(tomorrowDate).getMonth() + 1 }}月{{ new Date(tomorrowDate).getDate() }}日计划汇总</n-gradient-text>
            <!-- 提示文本 -->
            <p style="margin-top: 15px; margin-bottom: 15px; text-align: center">请输入监护人的数量后，点击生成文本!</p>
            <!-- 数字输入框及生成按钮 -->
            <div style="display: flex">
                <n-input-number v-model:value="guardianVal" clearable style="flex: 1; margin-right: 15px" />
                <n-button @click="onGenerate" :disabled="guardianVal <= 0" type="primary">生成文本</n-button>
            </div>
            <!-- 等待生成提示信息及复制按钮 -->
            <div>
                <n-divider dashed v-if="resultText === ''">等待生成</n-divider>
                <n-button v-else @click="onCopy" type="success" style="min-width: 100%; margin-top: 15px; margin-bottom: 15px">点击复制</n-button>
            </div>
        </div>

        <!-- 生成后的文本 -->
        <div v-show="resultText" class="resultText">
            <pre style="white-space: pre-wrap; :break-word">{{ resultText }}</pre>
        </div>
    </div>
</template>

<script setup>
import { useRoute } from 'vue-router';
import { onMounted, ref } from 'vue';
import { useMessage } from 'naive-ui';

// 使用消息提示框
let message = '';

// 明日时间
let tomorrowDate = ref('');

// 班组名称
let TeamName = ref('');

// 监护人数量
let guardianVal = ref(0);

// 生成后的文本
let resultText = ref('');

// 开始生成文本
function onGenerate() {
    try {
        // 激活工作表 导入模板
        let activeSheet = Application.Worksheets.Item('导入模板');

        // 返回单元格引用
        function GetValue2(range) {
            return activeSheet.Range(range);
        }

        // 程序所需的基础信息
        let baseInfo = {
            // 队伍名称
            TeamName,
            // 队伍总在场人数 含队长、安全员等管理人员和后勤
            total: 40,
            // 管理人员数量 (队长、技术员、安全员)
            manageStaff: 6,
            // 后勤人员数量
            logistics: 2
        };

        // 获取当前表格已使用的行数
        let useLine = Application.ActiveSheet.UsedRange.Rows.Count;

        // 数据行的索引值
        let index = 2;

        // 作业区域名称
        let workAreaNames = [];

        // 作业类型名称
        let workTypeNames = [];

        // 作业人数
        let workingPeoples = 0;

        // 作业部位及内容
        let workContexts = [];

        // 处理每行数据
        while (useLine > 0) {
            // 判断 日计划内容描述 内容是否为空，以此判断是否存在计划
            if (GetValue2(`B${index}`).Value2 !== null) {
                // 判断当前计划是否是 "临时用电（项目部）", 如果是, 则跳过
                if (GetValue2(`F${index}`).Value2 !== '临时用电（项目部）') {
                    // 将 非临时用电的 作业部位及内容 填入 workContexts
                    workContexts.push(`${GetValue2(`E${index}`).Value2}: ${GetValue2(`B${index}`).Value2}`);

                    // 将 作业区域名称 填入 workAreaNames
                    workAreaNames.push(GetValue2(`E${index}`).Value2);
                }

                // 将 作业类型名称 填入 workTypeNames
                workTypeNames.push(GetValue2(`F${index}`).Value2.replace('（项目部）', ''));
            }

            // 结束循环次数条件减一
            useLine--;

            //数据行的索引值 + 1
            index++;
        }

        // 如果作业部位及内容不为0, 新建文档并填充结果
        if (workContexts.length) {
            // 作业人数 = 总在场人数 - 管理人员 - 后勤 - 输入的监护人数(guardians)
            workingPeoples = baseInfo.total - baseInfo.manageStaff - baseInfo.logistics - guardianVal.value;

            // 处理明日时间
            tomorrowDate.value = tomorrowDate.value.replace(' 09:00:00', '').replaceAll('-', '.');

            // 生成 作业内容
            let context = '';
            let newSetWorkContexts = Array.from(new Set(workContexts));
            Array.from(new Set(workAreaNames)).map((area, index) => {
                let resText = '';
                newSetWorkContexts.map(item => {
                    let itemArr = item.split(':');
                    if (area === itemArr[0]) {
                        resText = `${resText}${resText === '' ? '' : '、'}${itemArr[1]}`;
                    }
                });
                context += `${index + 1}. ${area}: ${resText}\n`;
            });

            // 开始生成最后数据
            resultText.value = `南京工程/中化六建\n${tomorrowDate.value} ${baseInfo.TeamName.value}\n作业人数：${workingPeoples}人\n监护人：${guardianVal.value}人\n作业票区域：${Array.from(new Set(workAreaNames)).join('、')}\n作业票名称：${Array.from(new Set(workTypeNames)).join('、')}\n作业内容:\n${context}\n具体防范措施:机械、吊装作业半径严禁站人，周边设警示闭合围护，起重工持证上岗，做好履职，高处作业系挂好安全带，临时用电设备做好检查，作业过程中喇叭播放警示标语，进行安全交底，作业过程中互相提醒，专人监护，专职监督！`;
        }
    } catch (err) {
        alert(err.name);
    }
}

// 复制生成后的文本到剪切板
function onCopy() {
    if (resultText.value !== '') {
        // 复制文本到剪贴板
        navigator.clipboard.writeText(resultText.value).then(
            () => {
                message.success('复制成功，可以粘贴使用!');
            },
            () => {
                message.error('复制失败，请手动复制!');
            }
        );
    }
}

// 监听监护人数量变化，清空结果文本
watch(guardianVal, () => {
    resultText.value = '';
});

onMounted(() => {
    message = useMessage();

    const route = useRoute();
    tomorrowDate.value = route.query.tomorrowDate;
    TeamName.value = route.query.TeamName;
});
</script>

<style scoped>
#DailyPlan {
    width: 100%;
    height: 100%;
    font-family: '得意黑', sans-serif;
}

#DailyPlan .info {
    display: flex;
    flex-direction: column;
    justify-content: center;
}

#DailyPlan .resultText {
    max-height: 65vh;
    overflow-y: auto;
    padding: 15px;
    border: 1px solid #d9d9d9;
    border-radius: 5px;
    background-color: #f5f5f5;
}
</style>
