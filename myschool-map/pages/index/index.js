const qqmap = require('../../utils/qqmap')

const CENTER = { latitude: 22.686699, longitude: 108.266767 }

const CAMPUS = [
  { id: 1, name: '主教学楼', latitude: 22.687200, longitude: 108.265500 },
  { id: 2, name: '第一食堂', latitude: 22.686800, longitude: 108.268200 },
  { id: 3, name: '图书馆', latitude: 22.687600, longitude: 108.267000 },
  { id: 4, name: '体育馆', latitude: 22.688100, longitude: 108.268000 },
  { id: 5, name: '主校门', latitude: 22.685200, longitude: 108.266200 }
]

Page({
  data: {
    latitude: CENTER.latitude,
    longitude: CENTER.longitude,
    scale: 17,
    activeTab: 0,
    keyword: '',
    markers: [],
    polyline: [],
    results: [],
    userLat: CENTER.latitude,
    userLng: CENTER.longitude
  },

  onLoad() {
    this.setData({ markers: this.buildMarkers(CAMPUS) })
    wx.getLocation({
      type: 'gcj02',
      success: (res) => this.setData({ userLat: res.latitude, userLng: res.longitude }),
      fail: () => {}
    })
  },

  buildMarkers(list) {
    return list.map((m, i) => ({
      id: i,
      latitude: m.latitude,
      longitude: m.longitude,
      title: m.name,
      callout: {
        content: m.name,
        color: '#1A1D26',
        fontSize: 13,
        borderRadius: 8,
        bgColor: '#FFFFFF',
        padding: 8,
        display: 'BYCLICK',
        textAlign: 'center'
      }
    }))
  },

  switchTab(e) {
    const t = Number(e.currentTarget.dataset.t)
    this.setData({ activeTab: t, results: [] })
    if (t === 0) {
      this.setData({ markers: this.buildMarkers(CAMPUS), polyline: [] })
    }
    if (t === 1 && !this.data.results.length) {
      this.setData({ keyword: '食堂' })
      this.doSearch()
    }
  },

  onInput(e) { this.setData({ keyword: e.detail.value }) },

  doSearch() {
    const kw = this.data.keyword.trim()
    if (!kw) return
    wx.showLoading({ title: '搜索中' })
    qqmap.search(kw, CENTER.latitude, CENTER.longitude, 1500).then(list => {
      wx.hideLoading()
      this.setData({
        results: list,
        activeTab: 1,
        markers: this.buildMarkers(list.map(p => ({
          name: p.title,
          latitude: p.location.lat,
          longitude: p.location.lng
        })))
      })
    }).catch(() => {
      wx.hideLoading()
      wx.showToast({ title: '搜索失败，请检查 Key 配置', icon: 'none' })
    })
  },

  startRoute(target) {
    this.setData({ activeTab: 2 })
    wx.showLoading({ title: '路线规划中' })
    qqmap.walkingRoute(this.data.userLat, this.data.userLng, target.latitude, target.longitude).then(res => {
      wx.hideLoading()
      const pts = qqmap.decodePolyline(res.routes[0].polyline)
      this.setData({
        polyline: [{ points: pts, color: '#2E6BFF', width: 8, arrowLine: true }],
        markers: this.buildMarkers([
          { name: '我的位置', latitude: this.data.userLat, longitude: this.data.userLng },
          { name: target.name, latitude: target.latitude, longitude: target.longitude }
        ])
      })
    }).catch(() => {
      wx.hideLoading()
      wx.showToast({ title: '路线规划失败', icon: 'none' })
    })
  },

  pickResult(e) {
    const r = this.data.results[e.currentTarget.dataset.index]
    this.startRoute({ name: r.title, latitude: r.location.lat, longitude: r.location.lng })
  },

  onMarkerTap(e) {
    const id = e.markerId
    let m = null
    if (this.data.activeTab === 0) {
      m = CAMPUS[id]
    } else if (this.data.results[id]) {
      const r = this.data.results[id]
      m = { name: r.title, latitude: r.location.lat, longitude: r.location.lng }
    }
    if (m) this.startRoute(m)
  }
})
