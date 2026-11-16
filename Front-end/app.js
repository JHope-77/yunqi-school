App({
  globalData: {
    userInfo: null,
    isLogin: false
  },

  onLaunch() {
    const userInfo = wx.getStorageSync('userInfo')
    if (userInfo && userInfo.avatar) {
      this.globalData.userInfo = userInfo
      this.globalData.isLogin = true
    }
  },

  login(userInfo) {
    this.globalData.userInfo = userInfo
    this.globalData.isLogin = true
    wx.setStorageSync('userInfo', userInfo)
  },

  logout() {
    this.globalData.userInfo = null
    this.globalData.isLogin = false
    wx.removeStorageSync('userInfo')
  }
})
