<template>
  <div>
    <el-upload class="upload-demo" action="/" :http-request="(params) => handleRequest(params)">
      <el-button type="primary" icon="el-icon-upload2">点击上传表格</el-button>
    </el-upload>
  </div>
</template>

<script>
import xlsx from 'xlsx';
import { parseTime } from '@/utils'

export default {
  data() {
    return {
      importData: [],
      exportData: []
    }
  },
  methods: {
    extractOrderNumber(str) {

    },
    async handleRequest(params) {
      const { file } = params
      // 读取表格对象
      const buffer = await file.arrayBuffer()
      const workbook = xlsx.read(buffer, {
        type: 'buffer',
        cellDates: true,//设为true，将天数的时间戳转为时间格式
      });
      // 找到第一张表
      const sheetNames = workbook.SheetNames;
      const sheet1 = workbook.Sheets[sheetNames[0]];
      // 读取内容
      const jsonData = xlsx.utils.sheet_to_json(sheet1);
      const index = jsonData.findIndex(item => item['#支付宝收支明细'] === '流水号') + 1
      this.importData = jsonData.slice(index, jsonData.length - 1)
      console.log(this.importData);
      this.importData.forEach(item => {
        if ('__EMPTY_2' in item) {
          const str = item['__EMPTY_2']
          const startIndex = str.indexOf('订单号:') + 4
          const endIndex = str.lastIndexOf(")")
          console.log(str.slice(startIndex, endIndex));
        }


        this.exportData.push({
          '日期': parseTime(item['__EMPTY']),
          // '订单号': str.slice(startIndex, endIndex)
        })
      })
      console.log(this.exportData);

      // this.exportExcelFile([{
      //   "name": '里斯',
      //   "age": '12'
      // },{
      //   "name": '张三',
      //   'age': '999'
      // }])
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