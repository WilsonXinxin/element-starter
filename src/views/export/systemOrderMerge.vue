<template>
<div>
  <el-upload class="upload-demo" :file-list="fileList" action="#" multiple :auto-upload="false"
      :on-change="handleChange" :on-remove="handleChange">
      <el-button slot="trigger" size="small" type="primary" icon="el-icon-upload2">选取表格</el-button>
      <el-button :loading="parseLoading" style="margin-left: 10px;" size="small" type="success" @click="btnParse">{{
      parseLoading ? "解析中" : "解析并下载" }}</el-button>
      <div slot="tip" class="el-upload__tip">
        需要上传系统和订单两种表格
        <div v-if="tip" style="color:red">{{ tip }}</div>
      </div>
    </el-upload>
</div>
</template>

<script>
import xlsx from "xlsx"
import { parseTime, exportExcelFile } from '@/utils'

export default {
  data() {
    return {
      fileList: [],
      tip: false,
      parseLoading: false,
      maxLength: 0,
      tableData1: [],
      tableData2: [],
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
          const firstHeader = Object.keys(data[0])[0] // 读取第一个字段          
          switch (firstHeader) {
            case '订单号': // 后台订单表
              this.tableData1.push(...data)
              break;
            case '销售明细表': // 系统订单表
              const filterArr = data.slice(2)
              const ids = []
              filterArr.forEach(prod => {
                const list = {
                  '单据编号': prod['销售明细表_1'],
                  '客户名称': prod['销售明细表_5'],
                  '旺旺名': prod['销售明细表_18'],
                  '成交金额': prod['销售明细表_15'] || 0,
                }
                if(list['旺旺名'] === '线下') {
                  if(ids.includes(list['单据编号'])) {
                    const i = this.tableData2.findIndex(data => data['单据编号'] === list['单据编号'])
                    if(i >= 0) this.tableData2[i]['成交金额'] += list['成交金额']
                    return
                  } else {
                    ids.push(list['单据编号'])
                  }
                }
                this.tableData2.push(list)
              })
              break;
            default:
              break;
          }
        })
        // 等待文件解析完毕
        this.timer = setInterval(() => {
          if (this.lastIndex === (this.fileLength - 1)) {
            if (this.timer !== null) clearInterval(this.timer)
            this.tableData1.length && this.tableData2.length ? resolve() : reject('缺少表格~')
          }
        }, 100);
      })
    },
    async btnParse() {
      // 需要解析三种类型的表格
      if (this.fileLength < 2) {
        this.tip = '缺少表格~'
        return
      }
      try {
        this.parseLoading = true
        await this.parseXlsx(this.fileList)
        const exportData = []
        this.tableData2.forEach(item => {
          const list = {
            '系统订单客户': item['客户名称'],
            '旺旺名': item['旺旺名'],
            '系统成交金额': item['成交金额']
          }
          this.tableData1.forEach(key => {
            if(key['买家会员'] === item['旺旺名']) {
              list['订单状态'] = key['订单状态']
            }
          })
          exportData.push(list)
        })
        exportExcelFile(exportData, 'table1', `系统和后台订单合并_${parseTime(new Date(), '{y}-{m}-{d} {h}:{i}:{s}')}.xlsx`)
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