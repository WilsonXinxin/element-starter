<template>
  <div>
    <!-- 备货设置 -->
    <el-form ref="form" :model="formData" label-width="120px" style="width: 350px;">
      <el-form-item label="备货时间：">
        {{ curDate }}
      </el-form-item>
      <el-form-item label="按销量备货(天)：">
        <!-- <div style="display: flex;"> -->
        <el-select v-model="dayValue" placeholder="请选择" size="mini" style="width: 80px;">
          <el-option v-for="item in dayOptions" :key="item.value" :label="item.label" :value="item.value">
          </el-option>
        </el-select>
        <el-button type="primary" @click="handleProdTable" size="mini" style="margin-left:10px">生成备货表</el-button>
        <!-- </div> -->
      </el-form-item>
    </el-form>

    <!-- 销量表 库存表 -->
    <!-- <el-upload class="upload-demo" action="/" :http-request="(params) => handleRequest(params, 'count')" :limit="1"
      :before-remove="handleRemove('count')" :file-list="fileList">
      <el-button size="small" type="primary">点击上传商品销量表</el-button>
    </el-upload>

    <div style="margin: 20px 0;">
      <el-table :data="tableData" border>
        <el-table-column v-for="(item, index) in tableHeader" align="center" :key="index" :prop="item" :label="item" />
      </el-table>
    </div>
    <el-upload class="upload-demo" action="/" :http-request="(params) => handleRequest(params, 'inventory')"
      :file-list="fileList2">
      <el-button size="small" type="primary">点击上传商品库存表</el-button>
    </el-upload>

    <div style="margin: 20px 0;">
      <el-table :data="tableInventoryData" border>
        <el-table-column v-for="(item, index) in tableInventoryHeader" align="center" :key="index" :prop="item"
          :label="item" />
      </el-table>
    </div> -->

    <!-- 备货表 -->
    <el-upload class="upload-demo" action="/" :http-request="(params) => handleRequest(params)">
      <el-button size="small" type="primary">点击上传表格</el-button>
    </el-upload>

    <div style="margin: 20px 0;">
      <el-table v-loading="loading" :data="prodTableData" border>
        <el-table-column v-for="(item, index) in prodTableHeader" align="center" :key="index" :prop="item"
          :label="item" />
      </el-table>
    </div>
  </div>
</template>

<script>
import xlsx from 'xlsx';
import moment from "moment";

export default {
  data() {
    return {
      loading: false,
      tableHeader: [],
      tableData: [
        //   {
        //   date: '2016-05-02',
        //   name: '王小虎',
        //   address: '上海市普陀区金沙江路 1518 弄'
        // }, {
        //   date: '2016-05-04',
        //   name: '王小虎',
        //   address: '上海市普陀区金沙江路 1517 弄'
        // }, {
        //   date: '2016-05-01',
        //   name: '王小虎',
        //   address: '上海市普陀区金沙江路 1519 弄'
        // }, {
        //   date: '2016-05-03',
        //   name: '王小虎',
        //   address: '上海市普陀区金沙江路 1516 弄'
        // }
      ],
      tableInventoryData: [],
      tableInventoryHeader: [],
      typeList: {
        count: 'table',
        inventory: 'tableInventory'
      },
      fileList: [],
      fileList2: [],
      formData: {},
      dayValue: 7,
      dayOptions: [
        {
          value: 7,
          label: '7天'
        },
        // {
        //   value: 15,
        //   label: '15天'
        // },
        // {
        //   value: 30,
        //   label: '30天'
        // },
        {
          value: 45,
          label: '45天'
        },
      ],
      workbook: {},
      prodTableData: [],
      prodTableHeader: []
    }
  },
  computed: {
    curDate() {
      return moment(new Date()).format('YYYY-MM-DD')
    }
  },
  methods: {
    excelDateFormat(row, columnName) {
      //日期转换
      const date = row[columnName]
      if (date === undefined || date === null || date === "") {
        return null;
      }
      //非时间格式问题  返回Invalid date
      let retFormat = moment(date).format('YYYY-MM-DD');
      if (retFormat === "Invalid date") {
        return retFormat;
      }
      return moment(date).add(1, 'days').format('YYYY-MM-DD')
    },
    // async handleRequest(params, type) {
    //   const { file } = params
    //   this[`${this.typeList[type]}Header`] = []

    //   // 读取表格对象
    //   const buffer = await file.arrayBuffer()
    //   const workbook = xlsx.read(buffer, {
    //     type: 'buffer',
    //     cellDates: true,//设为true，将天数的时间戳转为时间格式
    //   });
    //   // 找到第一张表
    //   const sheetNames = workbook.SheetNames;
    //   const sheet1 = workbook.Sheets[sheetNames[0]];
    //   console.log(workbook);
    //   // 读取内容
    //   const jsonData = xlsx.utils.sheet_to_json(sheet1);
    //   console.log(jsonData);
    //   this[`${this.typeList[type]}Data`] = jsonData

    //   if (Object.keys(jsonData).length) {
    //     for (let key in jsonData[0]) {
    //       this[`${this.typeList[type]}Header`].push(key)
    //     }
    //     jsonData.forEach(item => {
    //       for (let key in item) {
    //         if (key === '销售日期') item[key] = this.excelDateFormat(item, key)
    //       }
    //     })
    //   }
    // },

    async handleRequest(params) {
      const { file } = params
      // 读取表格对象
      const buffer = await file.arrayBuffer()
      this.workbook = xlsx.read(buffer, {
        type: 'buffer',
        cellDates: true,//设为true，将天数的时间戳转为时间格式
      });

      this.handleProdTable()
    },
    handleRemove(type) {
      // console.log(type);
      // this[`${this.typeList[type]}Header`] = []
      // this[`${this.typeList[type]}Data`] = []
    },
    handleProdTable() {
      try {
        if (!Object.keys(this.workbook).length) {
          this.$message.warning('请先上传表格')
          return
        }
        this.loading = true
        this.prodTableHeader = []
        this.prodTableData = []
        // 第一张表为45天销量，第二张表为7天销量，第三张表为库存表
        let index = this.dayValue === 45 ? 0 : 1
        const sheetNames = this.workbook.SheetNames;
        // 读取销量表内容
        const salesJson = xlsx.utils.sheet_to_json(this.workbook.Sheets[sheetNames[index]]);
        // 读取库存表内容
        const inventoryJson = xlsx.utils.sheet_to_json(this.workbook.Sheets[sheetNames[2]]);
        console.log(salesJson);
        console.log(inventoryJson);

        inventoryJson.forEach((item, index) => {
          const list = {
            '商品asin': item['asin'],
            // '商品名称': item['product-name'],
          }
          salesJson.forEach(prod => {
            if (prod['（子）ASIN'] === item['asin']) {
              list['销量'] = prod['已订购商品数量']
              list['库存数'] = item['afn-warehouse-quantity仓库数量']
              list['库存消耗天数'] = prod['已订购商品数量'] === 0 ? "~" : parseInt(item['afn-warehouse-quantity仓库数量'] / (prod['已订购商品数量'] / this.dayValue))
              list['是否备货'] = item['afn-warehouse-quantity仓库数量'] === 0 ? '是' : item['afn-warehouse-quantity仓库数量'] > prod['已订购商品数量'] ? '否' : '-'
              list['需求数'] = list['是否备货'] === '是' ? item['afn-warehouse-quantity仓库数量'] - prod['已订购商品数量'] : 0
            }
          })
          this.prodTableData[index] = list
        })

        console.log(this.prodTableData);
        if (this.prodTableData.length) {
          for (let key in this.prodTableData[0]) {
            this.prodTableHeader.push(key)
          }
        }
      } catch (error) {
        console.error(error)
      } finally {
        this.loading = false
      }
    }
  }
}
</script>