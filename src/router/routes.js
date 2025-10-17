export default [
    {
        path: '/',
        name: '默认页',
        component: () => import('../components/Root.vue')
    },
    {
        path: '/dailyplan',
        name: '默认页',
        component: () => import('../components/custom/DailyPlan.vue')
    }
];
