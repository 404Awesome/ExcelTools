<template>
    <n-message-provider>
        <RouterView />
    </n-message-provider>
</template>

<script setup>
import { onMounted } from 'vue';
import ribbon from './utils/ribbon.js';
import './utils/tableEvent.js';
import { NMessageProvider } from 'naive-ui';

onMounted(() => {
    window.ribbon = ribbon;
});
</script>

<style scoped></style>
