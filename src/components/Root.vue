<template>
    <div>
        <h1 class="title">Hello ExcelTools</h1>
        <n-divider>
            <span>描述</span>
        </n-divider>
        <p>欢迎使用本工具!</p>
        <p>WPS(Excel)工具集</p>
        <p>主要是解决电气或仪表过程资料,提供相关函数,用已加快速度!</p>
        <p>或是解决平日工作中遇到的问题</p>
        <p>爱来自中国!</p>

        <n-divider>
            <span>技术栈</span>
        </n-divider>
        <n-grid :cols="3">
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-vscode-icons:file-type-js-official" />
                    <span>JavaScript</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-hugeicons:wps-office text-2xl" />
                    <span>WPS JS</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-logos-vue" />
                    <span>Vue3</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-devicon:vitejs" />
                    <span>Vite</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-devicon:vscode" />
                    <span>VSCode</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-material-icon-theme:unocss" />
                    <span>UnoCSS</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-devicon:sass text-2xl" />
                    <span>Sass</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-logos:naiveui" />
                    <span>NaiveUI</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-simple-icons:axios" />
                    <span>Axios</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-logos-vue" />
                    <span>VueRouter</span>
                </div>
            </n-grid-item>
            <n-grid-item>
                <div class="gridItem">
                    <span class="i-material-icon-theme:eslint text-2xl" />
                    <span>Eslint</span>
                </div>
            </n-grid-item>
        </n-grid>
    </div>
</template>

<script></script>

<style scoped>
.title {
    text-align: center;
    font-family: '得意黑', sans-serif;
}

.gridItem {
    height: 55px;
    display: flex;
    justify-content: center;
    align-items: center;
    column-gap: 10px;
    font-family: '得意黑', sans-serif;
    cursor: pointer;
    font-size: 20px;
    color: #999999;
}

.gridItem:hover {
    background-color: #f1f0f0;
    color: #444444;
}
</style>
