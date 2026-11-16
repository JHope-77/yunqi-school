// 腾讯位置服务 WebService API 封装
// 使用前：把下面 YOUR_KEY 替换为你在 lbs.qq.com 申请的 Key（需勾选 WebService API）
const KEY = 'FXVBZ-2JQEZ-ZRLXT-773SW-FUTOO-VBFCQ'

// 地点搜索：在学校周边 radius 米内搜索 keyword
function search(keyword, lat, lng, radius = 1500) {
  return new Promise((resolve, reject) => {
    wx.request({
      url: 'https://apis.map.qq.com/ws/place/v1/search',
      data: {
        keyword,
        boundary: 'nearby(' + lat + ',' + lng + ',' + radius + ')',
        page_size: 10,
        key: KEY
      },
      success: (res) => {
        if (res.data && res.data.status === 0) resolve(res.data.data)
        else reject(res.data)
      },
      fail: reject
    })
  })
}

// 步行路线规划：from/to 为 'lat,lng'
function walkingRoute(fromLat, fromLng, toLat, toLng) {
  return new Promise((resolve, reject) => {
    wx.request({
      url: 'https://apis.map.qq.com/ws/direction/v1/walking',
      data: {
        from: fromLat + ',' + fromLng,
        to: toLat + ',' + toLng,
        key: KEY
      },
      success: (res) => {
        if (res.data && res.data.status === 0) resolve(res.data.result)
        else reject(res.data)
      },
      fail: reject
    })
  })
}

// 腾讯 polyline 解码：返回值为 [lat0*1e6, lng0*1e6, dlat1*1e6, dlng1*1e6, ...]，后续为偏移
function decodePolyline(plist) {
  const pts = []
  let lat = plist[0] / 1e6
  let lng = plist[1] / 1e6
  pts.push({ latitude: lat, longitude: lng })
  for (let i = 2; i < plist.length; i += 2) {
    lat += plist[i] / 1e6
    lng += plist[i + 1] / 1e6
    pts.push({ latitude: lat, longitude: lng })
  }
  return pts
}

module.exports = { search, walkingRoute, decodePolyline }
