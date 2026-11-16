const app = getApp()

Page({
  data: { userInfo: {} },

  onShow() {
    this.setData({ userInfo: app.globalData.userInfo || {} })
  },

  logout() {
    wx.showModal({
      title: '提示',
      content: '确定退出登录？',
      success: (res) => {
        if (res.confirm) {
          app.logout()
          wx.redirectTo({ url: '/pages/login/login' })
        }
      }
    })
  }
})
