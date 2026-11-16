const { call } = require('../../utils/request')

Page({
  data: {
    cats: ['全部', '教材', '电子', '生活用品', '代步'],
    curCat: '全部',
    list: []
  },

  onShow() { this.loadList() },

  loadList() {
    call('marketList').then(res => {
      const all = res.data || []
      const list = this.data.curCat === '全部' ? all : all.filter(g => g.cat === this.data.curCat)
      this.setData({ list })
    })
  },

  pickCat(e) {
    this.setData({ curCat: e.currentTarget.dataset.c }, () => this.loadList())
  },

  goDetail(e) {
    wx.navigateTo({ url: '/pages/goods-detail/goods-detail?id=' + e.currentTarget.dataset.id })
  },

  goPublish() {
    wx.navigateTo({ url: '/pages/publish-goods/publish-goods' })
  }
})
