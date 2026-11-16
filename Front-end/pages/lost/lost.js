const { call } = require('../../utils/request')

Page({
  data: {
    tabs: ['寻物', '招领'],
    cur: 0,
    list: []
  },

  onShow() { this.loadList() },

  loadList() {
    call('lostList').then(res => {
      const all = res.data || []
      const type = this.data.cur === 0 ? 'miss' : 'found'
      this.setData({ list: all.filter(x => x.type === type) })
    })
  },

  pickTab(e) {
    this.setData({ cur: Number(e.currentTarget.dataset.i) }, () => this.loadList())
  },

  goPublish() {
    wx.navigateTo({ url: '/pages/publish-lost/publish-lost' })
  }
})
