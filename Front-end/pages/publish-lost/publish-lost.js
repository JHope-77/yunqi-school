Page({
  data: {
    type: 'miss',
    title: '',
    place: '',
    time: '',
    desc: '',
    contact: ''
  },

  onInput(e) {
    this.setData({ [e.currentTarget.dataset.field]: e.detail.value })
  },

  pickType(e) {
    this.setData({ type: e.currentTarget.dataset.t })
  },

  onSubmit() {
    const { title, place, contact } = this.data
    if (!title || !place || !contact) {
      wx.showToast({ title: '请填写标题、地点和联系方式', icon: 'none' })
      return
    }
    wx.showToast({ title: '发布成功（演示）', icon: 'success' })
    setTimeout(() => wx.navigateBack(), 800)
  }
})
