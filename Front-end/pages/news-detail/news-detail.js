const { call } = require('../../utils/request')

Page({
  data: { detail: null },

  onLoad(options) {
    call('newsDetail', { id: options.id }).then(res => this.setData({ detail: res.data }))
  }
})
