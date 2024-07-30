<template>
  <div>
    <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false"
      :on-change="handleChange" :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? "解析中" : "解析并下载" }}</el-button>
      <div slot="tip" class="el-upload__tip">
        需要上传三种表格
        <div v-if="tip" style="color:red">{{ tip }}</div>
      </div>
    </el-upload>
  </div>
</template>

<script>
import xlsx from "xlsx"
import { parseTime, extractNumberFromStart, extractStr, exportExcelFile, isAttr } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      csvData: [],
      xlsxData: [],
      tip: false,
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
      maxLength: 0,
      tableData1: [],
      tableData2: [],
      tableData3: [],
      fileLength: 0,
      timer: null,
      lastIndex: 1
    }
  },
  methods: {
    handleChange(file, fileList) {
      // console.log(fileList);
      this.fileList = fileList
      this.fileLength = this.fileList.length
      if (this.fileList.length) this.tip = ''
    },
    async parseXlsx(xlsxFileArr) {
      return new Promise(async (resolve, reject) => {
        this.tableData1 = []
        this.tableData2 = []
        this.tableData3 = []
        await xlsxFileArr.forEach(async ({ raw }, index) => {
          this.lastIndex = index
          // 读取表格对象
          const buffer = await raw.arrayBuffer()
          const workbook = xlsx.read(buffer, {
            type: 'buffer',
            cellDates: true, //设为true，将天数的时间戳转为时间格式
          });
          const sheetNames = workbook.SheetNames;
          // 读取第一张表的内容
          const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
          const firstHeader = Object.keys(data[0])[0]
          switch (firstHeader) {
            case '订单号': // 主表 订单后台表
              this.tableData1.push(...data)
              break;
            case '#支付宝收支明细': // 支付宝流水表              
              const filterArr = data.filter(item => ('__EMPTY_2' in item) && item['__EMPTY_2'].toString().includes('订单号'))
              filterArr.forEach(prod => {
                this.tableData2.push({
                  '订单号': extractNumberFromStart(isAttr(prod, '__EMPTY_2'), '订单号'),
                  '支付宝扣款': isAttr(prod, '__EMPTY_4'),
                  '支出摘要': extractStr(isAttr(prod, '__EMPTY_2'), '订单号')
                })
              })
              break;
            default: // 后台账单表
              if (firstHeader.includes('#账号')) {
                const accountName = firstHeader.substring(4, firstHeader.length)             
                const filterArr = data.filter(item => ('__EMPTY_2' in item) && item['__EMPTY_2'].toString().includes('NP'))
                console.log(data);
                filterArr.forEach(prod => {
                  const index = prod['__EMPTY_2'].indexOf('NP') + 2
                  this.tableData3.push({
                    '订单号': prod['__EMPTY_2'].substring(index, prod['__EMPTY_2'].length),
                    '店铺': accountName,
                    '已到账金额': prod['__EMPTY_11']
                  })
                })
              }
              break;
          }
        })
        // 等待文件解析完毕
        this.timer = setInterval(() => {
          if (this.lastIndex === (this.fileLength - 1)) {
            if (this.timer !== null) clearInterval(this.timer)
            this.tableData1.length ? resolve() : reject('解析表格数据为空，请检查表格类型是否正确~')
          }
        }, 100);
      })
    },
    mergeTable() {
      this.tableData1.forEach(prod => {
        this.tableData3.forEach(item => {
          if (prod['订单号'] === item['订单号']) {
            prod['已到账金额'] = item['已到账金额']
            prod['店铺'] = item['店铺']
          }
        })
      })
    },
    async btnParse() {
      // 需要解析三种类型的表格
      if (this.fileLength < 3) {
        this.tip = '缺少表格~'
        return
      }
      try {
        this.parseLoading = true
        await this.parseXlsx(this.fileList)
        this.mergeTable()
        const exportData = []
        // 找出最大有多少笔重复订单扣款
        this.tableData1.forEach(item => {
          const csvArr = this.tableData2.filter(prod => prod['订单号'] === item['订单号'])
          this.maxLength = csvArr.length > this.maxLength ? csvArr.length : this.maxLength
        })
        this.tableData1.forEach(item => {
          const csvArr = this.tableData2.filter(prod => prod['订单号'] === item['订单号'])
          const title = isAttr(item, '货品标题')
          let color = ''
          let specs = ''
          if (title) {
            const colorIndex = title.indexOf('颜色:')
            const lastSpace = title.lastIndexOf('(')
            const colorEnd = lastSpace > colorIndex + 3 ? lastSpace : title.length
            color = colorIndex >= 0 ? title.substring(colorIndex + 3, colorEnd) : ''
            const specsIndex = title.indexOf('机身内存:')
            const specsEnd = title.indexOf('下单备注')
            specs = specsIndex >= 0 ? title.substring(specsIndex + 5, specsEnd) : ''
          }
          const list = {
            '订单创建时间': isAttr(item, '订单创建时间'),
            '店铺': isAttr(item, '店铺'),
            '订单号': isAttr(item, '订单号'),
            '订单状态': isAttr(item, '订单状态'),
            '买家公司名称': isAttr(item, '买家公司名称'),
            '买家会员': isAttr(item, '买家会员'),
            '实付款': isAttr(item, '实付款（元）'),
            '已到账金额': isAttr(item, '已到账金额'),
            // '支付宝扣款': '',
            // '支出摘要': '',
            '货品总价': isAttr(item, '货品总价'),
            '运费': isAttr(item, '运费（元）'),
            '商家改价': isAttr(item, '商家改价（元）'),
            '货品标题': title,
            '规格': specs,
            '颜色': color,
            '单价': isAttr(item, '单价'),
            '数量': isAttr(item, '数量'),
            '平台型号': isAttr(item, '型号'),
            '内部型号': isAttr(item, '货号'),
            '收货人姓名': isAttr(item, '收货人姓名'),
            '联系电话': isAttr(item, '联系电话'),
            '货运公司': isAttr(item, '货运公司'),
            '运单号': isAttr(item, '运单号'),
          }
          // 在指定位置插入支付宝扣款
          const entries = Object.entries(list);
          // 找到要插入的位置
          const insertIndex = entries.findIndex(entry => entry[0] === '已到账金额') + 1;
          for (let i = 0; i < this.maxLength; i++) {
            entries.splice(insertIndex + i, 0, [`支付宝扣款${i + 1}`, (csvArr.length && csvArr.length > i) ? csvArr[i]['支付宝扣款'] : '']);
          }
          entries.splice(insertIndex + this.maxLength, 0, ['支出摘要', (csvArr.length && csvArr.length >= this.maxLength - 1) ? csvArr[0]['支出摘要'] : ''])
          exportData.push(Object.fromEntries(entries))
        })
        exportExcelFile(exportData, 'table1', `订单合并_${parseTime(new Date(), '{y}-{m}-{d} {h}:{i}:{s}')}.xlsx`)
      } catch (error) {
        console.error(error)
        this.tip = error
      } finally {
        this.parseLoading = false
      }
    }
  },
}
</script>
