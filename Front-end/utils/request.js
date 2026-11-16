// 请求封装：当前阶段返回 Mock 数据；后端联调时改为 wx.cloud.callFunction({ name, data })
const mock = require('./mock')

function call(name, data = {}) {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      try {
        let payload = null
        switch (name) {
          case 'marketList':
            payload = mock.goods
            break
          case 'marketDetail':
            payload = mock.goods.find(g => g.id === Number(data.id)) || null
            break
          case 'lostList':
            payload = mock.lost
            break
          case 'newsList':
            payload = mock.news
            break
          case 'newsDetail':
            payload = mock.news.find(n => n.id === Number(data.id)) || null
            break
          case 'schedule':
            payload = mock.schedule
            break
          default:
            payload = []
        }
        resolve({ code: 0, data: payload })
      } catch (e) {
        reject(e)
      }
    }, 200)
  })
}

module.exports = { call }
