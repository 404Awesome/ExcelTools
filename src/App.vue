<template>
    <n-message-provider>
        <RouterView />
    </n-message-provider>
</template>

<script setup>
import { onMounted } from 'vue';
import ribbon from './utils/ribbon.js';
import { NMessageProvider } from 'naive-ui';

// 引入表格事件
import './utils/tableEvent.js';

// 引入自定义函数 - 字符串处理
import './functions/stringHandle.js';

// 引入自定义函数 - 仪表调校
import './functions/instrument.js';

// 引入自定义函数 - 六化建处理
import './functions/ZHLJ.js';

onMounted(() => {
    window.ribbon = ribbon;
});
</script>

<style scoped></style>
