<template>
    <div>
        <h1>{{ msg }}</h1>
    </div>
</template>

<script>
import('@/utils/functions.js');
export default {
    name: 'HelloWps',
    data() {
        return {
            msg: '欢迎来到wps加载项的世界!'
        };
    }
};
</script>

<style scoped></style>
