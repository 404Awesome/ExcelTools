export default [
    {
        path: '/',
        name: '默认页',
        component: () => import('../components/Root.vue')
    },
    {
        path: '/function',
        name: '自定义函数信息',
        component: () => import('../components/FunctionInfo.vue')
    },
    {
        path: '/dailyplan',
        name: '每日计划',
        component: () => import('../components/NanHua/DailyPlan.vue')
    },
    {
        path: '/ybinstall',
        name: '仪表安装检查记录',
        component: () => import('../components/YBProcessDoc/YBInstall.vue')
    }
];
