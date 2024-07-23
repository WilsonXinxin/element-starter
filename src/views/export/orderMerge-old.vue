<template>
  <div>
    <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false"
      :on-change="handleChange" :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? "解析中" : "解析并下载" }}</el-button>
      <div slot="tip" class="el-upload__tip">
        需要上传两种表格
        <div v-if="showTip" style="color:red">请先上传文件~</div>
      </div>
    </el-upload>
  </div>
</template>

<script>
import xlsx from "xlsx"
import Papa from "papaparse"
import { parseTime, extractNumberFromStart, extractStr, exportExcelFile } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      csvData: [],
      xlsxData: [],
      showTip: false,
      parseLoading: false,
      lastIndex: 1,
      timer: null,
      csvTableHeader: {
        '流水号': 0,
        '时间': 1,
        '名称': 2,
        '备注': 3,
        '收入': 4,
        '支出': 5,
        '账户余额': 6,
        '资金渠道': 7,
        '账户名': 8,
        '订单号': 9,
        '摘要': 10
      },
      maxLength: 0
    }
  },
  methods: {
    handleChange(file, fileList) {
      console.log(fileList);
      this.fileList = fileList
      if (this.fileList.length) this.showTip = false
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
          const salesJson = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
          this.xlsxData.push(...salesJson)
          this.xlsxData.length ? resolve() : reject('xlsx table is empty')
        })
      })
    },
    async parseCsv(fileList) {
      return new Promise((resolve, reject) => {
        this.csvData = []
        // 解析csv表格
        fileList.forEach((file) => {
          Papa.parse(file.raw, {
            encoding: "GB2312", // 编码格式
            complete: ({ data }) => {
              const index = data.findIndex((item) => item[0] === "流水号") + 1
              const list = [...data.slice(index, data.length - 2)]
              // 添加账户名
              const accountIndex = data.findIndex((item) => item[0] === "#账户名") + 1
              const accountName = data[accountIndex][0].slice(1, data[accountIndex][0].length)
              list.forEach(item => {
                // 订单号
                const orderId = (item[3] && item[3].indexOf('订单号') !== -1) ? extractNumberFromStart(item[3], '订单号') : ''
                // 摘要
                const desc = (item[3] && item[3].indexOf('订单号') !== -1) ? extractStr(item[3], '订单号') : item[3]
                item.push(...[accountName, orderId, desc])
              })

              this.csvData.push(...list)
            },
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
        }, 200)
      })
    },
    isAttr(obj, name) {
      return (name in obj) ? obj[name] : ''
    },
    async btnParse() {
      // 需要解析三种类型的表格
      const csvFileArr = this.fileList.filter(item => item.name.indexOf('.csv') !== -1)
      const xlsxFileArr = this.fileList.filter(item => item.name.indexOf('.xlsx') !== -1)
      if (!csvFileArr.length || !xlsxFileArr.length) {
        this.showTip = true
        return
      }

      try {
        this.parseLoading = true
        await this.parseXlsx(xlsxFileArr)
        await this.parseCsv(csvFileArr)
        const exportData = []
        // 找出最大有多少笔重复订单扣款
        this.xlsxData.forEach(item => {
          const orderId = item['订单号']
          const csvArr = this.csvData.filter(item => item[this.csvTableHeader['订单号']] === orderId)
          this.maxLength = csvArr.length > this.maxLength ? csvArr.length : this.maxLength
        })
        this.xlsxData.forEach(item => {
          const orderId = item['订单号']
          const csvArr = this.csvData.filter(item => item[this.csvTableHeader['订单号']] === orderId)
          const list = {
            '订单创建时间': this.isAttr(item, '订单创建时间'),
            '订单号': this.isAttr(item, '订单号'),
            '订单状态': this.isAttr(item, '订单状态'),
            '买家公司名称': this.isAttr(item, '买家公司名称'),
            '买家会员': this.isAttr(item, '买家会员'),
            '实付款': this.isAttr(item, '实付款（元）'),
            // '支付宝扣款': '',
            // '支出摘要': '',
            '货品总价': this.isAttr(item, '货品总价'),
            '运费': this.isAttr(item, '运费（元）'),
            '商家改价': this.isAttr(item, '商家改价（元）'),
            '货品标题': this.isAttr(item, '货品标题'),
            '单价': this.isAttr(item, '单价'),
            '数量': this.isAttr(item, '数量'),
            '型号': this.isAttr(item, '型号'),
            '收货人姓名': this.isAttr(item, '收货人姓名'),
            '联系电话': this.isAttr(item, '联系电话'),
            '货运公司': this.isAttr(item, '货运公司'),
            '运单号': this.isAttr(item, '运单号'),
          }
          // 在指定位置插入支付宝扣款
          const entries = Object.entries(list);
          // 找到要插入的位置
          const insertIndex = entries.findIndex(entry => entry[0] === '实付款') + 1;
          for (let i = 0; i < this.maxLength; i++) {
            entries.splice(insertIndex + i, 0, [`支付宝扣款${i + 1}`, (csvArr.length && csvArr.length > i) ? csvArr[i][this.csvTableHeader['支出']] : '']);
          }
          entries.splice(insertIndex + this.maxLength, 0, ['支出摘要', (csvArr.length && csvArr.length >= this.maxLength - 1) ? csvArr[0][this.csvTableHeader['摘要']] : ''])
          exportData.push(Object.fromEntries(entries))
        })
        exportExcelFile(exportData, 'table1', `订单合并_${parseTime(new Date(), '{y}-{m}-{d} {h}:{i}:{s}')}.xlsx`)
      } catch (error) {
        console.error(error)
      } finally {
        this.parseLoading = false
      }
    }
  },
}
</script>
