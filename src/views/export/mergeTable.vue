<template>
  <div>
    <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false" :on-change="handleChange"
      :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? '解析中' : '解析并下载' }}</el-button>
      <div slot="tip" v-if="showTip" class="el-upload__tip" style="color:red">请先上传文件~</div>
    </el-upload>
  </div>
</template>

<script>
import xlsx from 'xlsx';
import Papa from 'papaparse'
import { parseTime } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      importData: [],
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
    btnParse() {
      if (!this.fileList.length) {
        this.showTip = true
        return
      }

      this.parseLoading = true
      this.importData = []
      this.exportData = []
      this.fileList.forEach(item => {
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
            this.importData.push(...list)
          }
        })
      })

      // 定时器判断文件是否解析完
      this.timer = setInterval(() => {
        if (this.firstIndex !== this.importData.length) {
          this.firstIndex = this.importData.length
        } else {
          if (this.timer !== null) clearInterval(this.timer)
          this.handleExportData()
        }
      }, 200);
    },
    extractNumberFromStart(str, startChar) {
      const startIndex = str.indexOf(startChar) + 1;
      const numStr = str.slice(startIndex).match(/\d+/);
      return numStr ? numStr[0] : null;
    },
    extractStr(str, startChar) {
      const startIndex = str.indexOf(startChar) - 1;
      let sliceStr = str.slice(0, startIndex)
      if(str.indexOf('订单号:')) sliceStr += ')'
      return sliceStr
    },
    handleExportData() {
      console.log(this.importData);
      this.importData.forEach(item => {
        const list = {
          '日期': item[1],
          '账户名': item[8],
          '订单号': (item[3] && item[3].indexOf('订单号') !== -1) ? this.extractNumberFromStart(item[3], '订单号') : '',
          '收入': item[4],
          '支出': item[5],
          '摘要': (item[3] && item[3].indexOf('订单号') !== -1) ? this.extractStr(item[3], '订单号') : item[3],
          '支付方式': item[2],
          '余额': item[6],
          '资金渠道': item[7]
        }

        this.exportData.push(list)
      })
      this.exportExcelFile(this.exportData, 'table1', `example_${parseTime(new Date(), '{y}-{m}-{d}')}.xlsx`)
      this.parseLoading = false
    },
    exportExcelFile(array, sheetName = '表1', fileName = 'example.xlsx') {
      const jsonWorkSheet = xlsx.utils.json_to_sheet(array);
      const workBook = {
        SheetNames: [sheetName],
        Sheets: {
          [sheetName]: jsonWorkSheet,
        }
      };
      return xlsx.writeFile(workBook, fileName);
    }
  }
}
</script>