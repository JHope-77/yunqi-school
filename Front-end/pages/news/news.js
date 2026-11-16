const { call } = require('../../utils/request')

Page({
  data: { list: [] },

  onShow() {
    call('newsList').then(res => this.setData({ list: res.data || [] }))
  },

  goDetail(e) {
    wx.navigateTo({ url: '/pages/news-detail/news-detail?id=' + e.currentTarget.dataset.id })
  }
})
