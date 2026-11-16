const app = getApp()

Page({
  data: {
    avatar: '',
    nickName: ''
  },

  onChooseAvatar(e) {
    this.setData({ avatar: e.detail.avatarUrl })
  },

  onNickInput(e) {
    this.setData({ nickName: e.detail.value.trim() })
  },

  onLogin() {
    const { avatar, nickName } = this.data
    if (!avatar || !nickName) {
      wx.showToast({ title: '请先选择头像和昵称', icon: 'none' })
      return
    }
    app.login({ avatar, nickName, role: 'user' })
    wx.showToast({ title: '登录成功', icon: 'success' })
    setTimeout(() => {
      wx.switchTab({ url: '/pages/index/index' })
    }, 600)
  }
})
