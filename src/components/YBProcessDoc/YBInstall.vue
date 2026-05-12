<template>
    <div id="YBInstall">
        <!-- 标题 -->
        <h1 class="title">仪表安装检查记录</h1>

        <!-- 选择仪表按钮 -->
        <div class="btn">
            <n-button strong type="primary" block="false" @click="onChoose" v-show="handleRange.length === 0">选择仪表位号区域</n-button>
            <n-button strong type="warning" block="false" @click="onChoose" v-show="handleRange.length !== 0">重新选择仪表位号区域</n-button>
        </div>

        <!-- 复制结果的抽屉 -->
        <n-drawer :show="showDrawer" :default-height="'87.3vh'" placement="bottom">
            <n-drawer-content title="复制结果" header-style="display: none;">
                <!-- 返回上级按钮 -->
                <n-button type="error" @click="updateDrawer" class="title" style="margin: 22px 15px 0px 15px; width: 92%">返回上级</n-button>
                <!-- 分割线 -->
                <n-divider dashed style="padding: 0px 15px">复制结果</n-divider>
                <!-- 复制结果按钮组 -->
                <n-scrollbar style="max-height: 65vh">
                    <n-grid :x-gap="12" :y-gap="12" :cols="3" style="padding: 0px 15px">
                        <n-grid-item v-for="(value, index) in pendingArr">
                            <n-button @click="onCopy(value, index)" style="width: 100%">第{{ index + 1 }}页</n-button>
                        </n-grid-item>
                    </n-grid>
                </n-scrollbar>
            </n-drawer-content>
        </n-drawer>

        <!-- 选择仪表后显示内容 -->
        <div v-if="handleRange.length">
            <!-- 选择结果的明细 -->
            <n-scrollbar style="max-height: 60.2vh">
                <n-table :bordered="true" :single-line="false" size="small">
                    <thead style="text-align: center">
                        <tr>
                            <th>序号</th>
                            <th>仪表位号</th>
                        </tr>
                    </thead>
                    <tbody style="text-align: center">
                        <tr v-for="(value, index) in handleRange">
                            <td>{{ index + 1 }}</td>
                            <td>{{ value }}</td>
                        </tr>
                    </tbody>
                </n-table>
            </n-scrollbar>

            <!-- 选择的结果汇总 -->
            <h3 style="text-align: center; margin: 20px 0px">共选择 {{ handleRange.length }} 台仪表!</h3>

            <!-- 复制处理后的结果 -->
            <n-button type="success" @click="updateDrawer" class="title" style="width: 100%">复制处理后的结果</n-button>
        </div>
        <!-- 未选择仪表显示内容 -->
        <div v-else>
            <n-empty size="large" description="未选择任何数据!" style="margin-top: 20vh"></n-empty>
        </div>
    </div>
</template>

<script setup>
import { onMounted, ref } from 'vue';
import { useMessage } from 'naive-ui';
import { chunkArray } from '../../utils/utils';

// 使用消息提示框
let message = '';

// 处理后的数组
let handleRange = ref([]);

// 分割后的数组
let pendingArr = ref([]);

// 显示抽屉
const showDrawer = ref(false);
function updateDrawer() {
    showDrawer.value = !showDrawer.value;
}

// 复制结果
function onCopy(result, index) {
    // 复制文本到剪贴板
    navigator.clipboard.writeText(result).then(
        () => message.success(`第${index + 1}页复制成功`),
        () => message.error(`第${index + 1}页复制失败`)
    );
}

// 选择需要处理的仪表位号区域
function onChoose() {
    try {
        // 重置处理后的数组
        handleRange.value = [];

        // 获得选择的区域
        let myRange = Application.InputBox('选择需要处理的仪表位号区域！', '仪表安装检查记录', undefined, undefined, undefined, undefined, undefined, 8).Value2;

        //判断输入值是否存在 或者 是否只选择了一个单元格
        if (!myRange || typeof myRange == 'string') {
            myRange = [[myRange]];
        }

        // 遍历区域
        myRange.map(item => {
            handleRange.value.push(...item);
        });

        // 二次分割数组;
        let maxLength = 0;
        let defaultBool = 'null';
        pendingArr.value = chunkArray(chunkArray(handleRange.value, 3), 6).map(item => {
            // 判断这个数组内最长的仪表位号，求字符串长度
            item.map(no1 =>
                no1.map(no2 => {
                    if (no2.length > maxLength) {
                        maxLength = no2.length;
                    }
                })
            );

            // 根据最长长度为低于这个长度的仪表位号追加空格
            item = item.map(no1 => {
                return no1.map(no2 => {
                    if (no2.length < maxLength) {
                        return (no2 += Array(maxLength - no2.length)
                            .fill(' ')
                            .join(''));
                    } else {
                        return no2;
                    }
                });
            });

            // 判断第一行的长度，如果过长则削减 \t 否则默认处理
            defaultBool = defaultBool === 'null' && item.map(no => `\t\t${no.join('\t\t\t')}`)[0].length > 52 ? true : false;
            return defaultBool ? item.map(no => `\t${no.join('\t\t')}`).join('\r\n') : item.map(no => `\t\t${no.join('\t\t\t')}`).join('\r\n');
        });
    } catch (err) {}
}

// 初始化消息提示框插件
onMounted(() => {
    message = useMessage();
});
</script>

<style scoped>
#YBInstall {
    width: 100%;
    height: 100%;
    margin-bottom: 0px;
}

.title {
    text-align: center;
    font-family: '得意黑', sans-serif;
}

.btn {
    margin: 20px 0px;
    font-family: '得意黑', sans-serif;
}
</style>
