Page({
  data: {
    title: '',
    cat: '教材',
    price: '',
    desc: '',
    contact: ''
  },

  onInput(e) {
    this.setData({ [e.currentTarget.dataset.field]: e.detail.value })
  },

  pickCat(e) {
    this.setData({ cat: e.currentTarget.dataset.c })
  },

  onSubmit() {
    const { title, price, desc, contact } = this.data
    if (!title || !price || !contact) {
      wx.showToast({ title: '请填写标题、价格和联系方式', icon: 'none' })
      return
    }
    wx.showToast({ title: '发布成功（演示）', icon: 'success' })
    setTimeout(() => wx.navigateBack(), 800)
  }
})
