const { call } = require('../../utils/request')

Page({
  data: {
    weeks: [5, 6, 7, 8, 9],
    currentWeek: 8,
    days: ['一', '二', '三', '四', '五'],
    currentDay: 1,
    schedule: {},
    rows: []
  },

  onLoad() {
    call('schedule').then(res => {
      const schedule = res.data || {}
      this.setData({ schedule }, () => this.buildRows())
    })
  },

  pickWeek(e) {
    this.setData({ currentWeek: Number(e.currentTarget.dataset.w) })
  },

  pickDay(e) {
    this.setData({ currentDay: Number(e.currentTarget.dataset.d) }, () => this.buildRows())
  },

  buildRows() {
    const { schedule, currentDay } = this.data
    const list = schedule[currentDay] || []
    const rows = []
    for (let i = 1; i <= 5; i++) {
      const c = list.find(x => x.sec === i)
      rows.push({
        sec: i,
        name: c ? c.name : '',
        room: c ? c.room : '',
        teacher: c ? c.teacher : '',
        color: c ? c.color : '#F0F2F7',
        empty: !c
      })
    }
    this.setData({ rows })
  }
})
