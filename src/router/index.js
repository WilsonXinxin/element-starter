import Vue from 'vue'
import Router from 'vue-router'

// 解决重复点击路由报错
const originalPush = Router.prototype.push
Router.prototype.push = function push(location) {
  return originalPush.call(this, location).catch(err => err)
}

Vue.use(Router)

const routers = [
  {
    path: '/',
    name: 'Home',
    component: () => import('@/views/home'),
  },
  {
    path: '/stockOrder',
    component: () => import('@/views/export/stockOrder'),
  },
  {
    path: '/orderMerge',
    component: () => import('@/views/export/orderMerge'),
  },
  {
    path: '/mergeTable',
    component: () => import('@/views/export/mergeTable'),
  },
  {
    path: '/systemOrderMerge',
    component: () => import('@/views/export/systemOrderMerge'),
  },
  {
    path: '/404',
    component: () => import('@/views/common/error-page/404'),
    hidden: true
  }
]

const createRouter = () => new Router({
  mode: 'history', // require service support
  scrollBehavior: () => ({ y: 0 }),
  routes: routers
})

const router = createRouter()

// Detail see: https://github.com/vuejs/vue-router/issues/1234#issuecomment-357941465
export function resetRouter() {
  const newRouter = createRouter()
  router.matcher = newRouter.matcher // reset router
}

export default router