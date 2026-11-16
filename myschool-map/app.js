App({
  onLaunch() {
    // 申请定位权限，失败也不影响地图展示（仅不显示"我的位置"）
    wx.getLocation({
      type: 'gcj02',
      success: () => {},
      fail: () => {}
    })
  }
})
