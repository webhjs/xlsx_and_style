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
import {export_json_to_excel,export_json_to_excel_custom,export_table_to_excel,export_table_to_excel_custom} from '@/assets/js/Export2Excel'
export default {
  methods: {
    /* 导出样式表格网上方法 */
    exportExcel () {
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
        export_json_to_excel_custom({
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
    
	/* 导出无样式表格网上方法 */
    exportToExcel() {
      if(this.tableData.length){
        // console.log(this.tableData)
        require.ensure([], () => {
            const {
                export_json_to_excel
            } = require('@/assets/js/Export2Excel');
            const tHeader = ['报警ID','级别','IP地址','设备类型','设备描述','消息','状态','发生时刻','最近发生','恢复时刻','处理时刻','报警次数','处理者','是否处理','处理备注'];
            const filterVal = ['NewAlarmId','AlarmLevel','IPAddress','TrsTypeName','SysName','Message','AlarmStatusName','OccTime','OccLastTime','RecoverTime','TimeHandel','AlarmTimes','LoginName','IsHandel','HndMessage'];
            const list = this.tableData;
            const data = this.formatJson(filterVal, list);
            export_json_to_excel(tHeader, data, '历史告警'+this.selectDateRange[0].format("yyyy-MM-dd").toString()+'至'+this.selectDateRange[1].format("yyyy-MM-dd").toString());
        })
      }
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
