const { call } = require('../../utils/request')

Page({
  data: { detail: null },

  onLoad(options) {
    call('marketDetail', { id: options.id }).then(res => {
      this.setData({ detail: res.data })
    })
  },

  contact() {
    wx.showToast({ title: '演示：调起微信咨询', icon: 'none' })
  }
})
