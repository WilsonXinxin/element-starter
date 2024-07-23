<template>
  <div>
    <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false"
      :on-change="handleChange" :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? '解析中' : '解析并下载' }}</el-button>
      <div slot="tip" v-if="tip" class="el-upload__tip" style="color:red">{{ tip }}</div>
    </el-upload>
  </div>
</template>

<script>
import xlsx from 'xlsx';
import { parseTime, extractNumberFromStart, extractStr, exportExcelFile, isAttr } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      xlsxData: [],
      parseLoading: false,
      tip: ''
    }
  },
  methods: {
    handleChange(file, fileList) {
      console.log(fileList);
      this.fileList = fileList
      if (this.fileList.length) this.tip = ''
    },
    async parseXlsx(xlsxFileArr) {
      return new Promise((resolve, reject) => {
        this.xlsxData = []
        xlsxFileArr.forEach(async ({ raw }) => {
          // 读取表格对象
          const buffer = await raw.arrayBuffer()
          const workbook = xlsx.read(buffer, {
            type: 'buffer',
            cellDates: true,//设为true，将天数的时间戳转为时间格式
          });
          const sheetNames = workbook.SheetNames;
          // 读取第一张表的内容
          const tableData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
          // 解析出账户名
          const accountIndex = tableData.findIndex(item => item['#支付宝收支明细'] === '#账户名') + 1
          const accountName = tableData[accountIndex]['#支付宝收支明细'].slice(1, tableData[accountIndex]['#支付宝收支明细'].length)
          // 解析出需要的表格数据
          const index = tableData.findIndex(item => item['#支付宝收支明细'] === '流水号') + 1
          if (index <= 0) reject('解析表格数据为空，请检查表格类型是否正确~')
          const list = [...tableData.slice(index, tableData.length - 1)]
          list.forEach(prod => {
            this.xlsxData.push({
              '日期': isAttr(prod, '__EMPTY'),
              '账户名': accountName,
              '订单号': (isAttr(prod, '__EMPTY_2') && isAttr(prod, '__EMPTY_2').indexOf('订单号') !== -1) ? extractNumberFromStart(isAttr(prod, '__EMPTY_2'), '订单号') : '',
              '收入': isAttr(prod, '__EMPTY_3'),
              '支出': isAttr(prod, '__EMPTY_4'),
              '摘要': (isAttr(prod, '__EMPTY_2') && isAttr(prod, '__EMPTY_2').indexOf('订单号') !== -1) ? extractStr(isAttr(prod, '__EMPTY_2'), '订单号') : isAttr(prod, '__EMPTY_2'),
              '支付方式': isAttr(prod, '__EMPTY_1'),
              '余额': isAttr(prod, '__EMPTY_5'),
              '资金渠道': isAttr(prod, '__EMPTY_6')
            })
          })
          this.xlsxData.length ? resolve() : reject('解析表格数据为空，请检查表格类型是否正确~')
        })
      })
    },
    async btnParse() {
      if (!this.fileList.length) {
        this.tip = '请先上传文件~'
        return
      }
      try {
        this.parseLoading = true
        await this.parseXlsx(this.fileList)
        exportExcelFile(this.xlsxData, 'table1', `example_${parseTime(new Date())}.xlsx`)
      } catch (error) {
        console.error(error)
        this.tip = error
      } finally {
        this.parseLoading = false
      }
    }
  }
}
</script>