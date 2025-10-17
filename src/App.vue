<template>
    <n-message-provider>
        <RouterView />
    </n-message-provider>
</template>

<script setup>
import { onMounted } from 'vue';
import ribbon from './utils/ribbon.js';
import { NMessageProvider } from 'naive-ui';
// 统一引入js资源
import './introduce.js';

onMounted(() => {
    window.ribbon = ribbon;
});
</script>

<style scoped></style>
