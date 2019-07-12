<template>
  <div class="error-page"  id="out-table">
    <table width="100%" border="1">
		<colgroup>
			<col name="col_1" width="12%">
			<col name="col_2" width="6%">
			<col name="col_3" width="10%" span="2">
		</colgroup>
		<colgroup name="col_4" :span="statementData.Columns.length-3" width="110px"></colgroup>
		<tr>
			<th>ISBN</th>
			<th>Title</th>
			<th>Price</th>
			<th>describe</th>
		</tr>
      .......
    </table>
    <el-button type="primary" @click="exportExcel">导出</el-button>
  </div>
</template>

<script>
import FileSaver from 'file-saver'
import XLSX from 'xlsx'
import {export_json_to_excel,export_table_to_excel,export_table_to_excel_custom} from '@/assets/js/Export2Excel'
export default {
  methods: {
    //生成表格
    exportExcel () {
        /* 导出样式表格网上方法 */
        const tHeader = ['船名', '船长', '货种', '载重吨', '净吨', '锚地', '预抵时间', '下锚时间', '预靠泊位'] //表头
        const title = ['锚地船舶', '', '', '', '', '', '', '', '']  //标题
        //表头对应字段
        const filterVal = ['NAME', 'VESSEL_LENGTH']
        const list = this.anchorTable 
        const data = this.formatJson(filterVal, list)
        data.map(item => {
          // console.log(item)
          item.map((i, index) => {
            if (!i) {
              item[index] = ''
            }
          })
        })
        const merges = ['A1:I1'] //合并单元格
        export_json_to_excel({
          title: title,
          header: tHeader,
          data,
          merges,
          filename: '锚地船舶',
          autoWidth: true,
          bookType: 'xlsx'
        })
	},
    formatJson(filterVal, jsonData) {
      return jsonData.map(v => filterVal.map(j => v[j]))
    }
	
	/* 导出样式表格自定义方法 */
    //生成表格
    exportExcel () {    
        var wb = XLSX.utils.table_to_book(document.querySelector('#out-table'))
        var dataInfo = wb.Sheets[wb.SheetNames[0]];
        export_table_to_excel_custom(dataInfo)
	}
	
	/* 导出无样式表格方法 */
    //生成表格
    exportExcel () {
        /* generate workbook object from table */
        var wb = XLSX.utils.table_to_book(document.querySelector('#out-table'))
        console.log(XLSX)

        /* get binary string as output */
        var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'array' })
        try {
            FileSaver.saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'sheetjs.xlsx')
        } catch (e) { if (typeof console !== 'undefined') console.log(e, wbout) }
        return wbout
	}
  }
}
</script>


<style scoped>
    ...
</style>
