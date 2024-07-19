<template>
  <div>
    <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false"
      :on-change="handleChange" :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? '解析中' : '解析并下载' }}</el-button>
      <div slot="tip" v-if="showTip" class="el-upload__tip" style="color:red">请先上传文件~</div>
    </el-upload>
  </div>
</template>

<script>
import Papa from 'papaparse'
import { parseTime, extractNumberFromStart, extractStr, exportExcelFile } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      csvData: [],
      exportData: [],
      showTip: false,
      parseLoading: false,
      lastIndex: 1,
      timer: null
    }
  },
  methods: {
    handleChange(file, fileList) {
      console.log(file, fileList);
      this.fileList = fileList
      if (this.fileList.length) this.showTip = false
    },
    parseCsv(fileList) {
      return new Promise((resolve, reject) => {
        this.csvData = []
        fileList.forEach(item => {
          Papa.parse(item.raw, {
            encoding: 'GB2312', // 编码格式
            complete: ({ data }) => {
              const index = data.findIndex(item => item[0] === '流水号') + 1
              const list = [...data.slice(index, data.length - 2)]
              const accountIndex = data.findIndex(item => item[0] === '#账户名') + 1
              const accountName = data[accountIndex][0].slice(1, data[accountIndex][0].length)
              list.forEach(item => {
                item.push(accountName)
              })
              this.csvData.push(...list)
            }
          })
        })

        // 定时器判断文件是否解析完
        this.timer = setInterval(() => {
          if (this.lastIndex !== this.csvData.length) {
            this.lastIndex = this.csvData.length
          } else {
            if (this.timer !== null) clearInterval(this.timer)
            this.csvData.length ? resolve() : reject('csv table is empty!')
          }
        }, 200);
      })
    },
    async btnParse() {
      if (!this.fileList.length) {
        this.showTip = true
        return
      }

      try {
        this.parseLoading = true
        await this.parseCsv(this.fileList)
        const exportData = []
        this.csvData.forEach(item => {
          const list = {
            '日期': item[1],
            '账户名': item[8],
            '订单号': (item[3] && item[3].indexOf('订单号') !== -1) ? extractNumberFromStart(item[3], '订单号') : '',
            '收入': item[4],
            '支出': item[5],
            '摘要': (item[3] && item[3].indexOf('订单号') !== -1) ? extractStr(item[3], '订单号') : item[3],
            '支付方式': item[2],
            '余额': item[6],
            '资金渠道': item[7]
          }
          exportData.push(list)
        })
        exportExcelFile(exportData, 'table1', `example_${parseTime(new Date())}.xlsx`)
      } catch (error) {
        console.error(error)
      } finally {
        this.parseLoading = false
      }
    }
  }
}
</script>