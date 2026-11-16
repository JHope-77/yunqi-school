const { call } = require('../../utils/request')

Page({
  data: {
    news: []
  },

  onShow() {
    const app = getApp()
    if (!app.globalData.isLogin) {
      wx.redirectTo({ url: '/pages/login/login' })
      return
    }
    this.loadNews()
  },

  loadNews() {
    call('newsList').then(res => {
      this.setData({ news: (res.data || []).slice(0, 3) })
    })
  },

  goGrid(e) {
    const url = e.currentTarget.dataset.url
    if (url) wx.navigateTo({ url })
  },

  goSearch() {
    wx.showToast({ title: '搜索演示', icon: 'none' })
  }
})
