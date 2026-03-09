
#' 返回元数据信息
#'
#' @param erpToken ERP口令
#'
#' @return 返回值
#' @export
#'
#' @examples
#' TaxStatement_meta()
TaxStatement_meta <- function(erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039') {
  sql = paste0("select FName,FTableName,FCells from [rds_t_Tax_StatementConfiguration]  ")
  data = tsda::sql_select2(token = erpToken,sql = sql)
  return(data)

}
#' 将excel列转
#'
#' @param col_str 列名
#'
#' @return 返回值
#'
#' @examples
#' excel_TaxStatemen_to_num()
excel_TaxStatemen_to_num <- function(col_str) {
  col_str <- toupper(col_str)
  chars <- strsplit(col_str, "")[[1]]
  sum <- 0
  for (char in chars) {
    sum <- sum * 26 + which(LETTERS == char)
  }
  return(sum)
}

# 将Excel坐标转换为行列数字
#' 将数据左边进行处理
#'
#' @param coord 提供坐标
#'
#' @return 返回值
#' @export
#'
#' @examples
#' excel_coord_to_numeric()
excel_coord_to_numeric <- function(coord) {
  col_str <- gsub("[^A-Za-z]", "", coord)
  row_num <- as.integer(gsub("[^0-9]", "", coord))
  col_num <- excel_TaxStatemen_to_num(col_str)
  return(c(col = col_num, row = row_num))
}

#' 将汇总表数据写入EXCEL
#'
#' @param erpToken ERP口令
#' @param FYear
#' @param FMonth
#' @param FOrgNumber
#' @param outputDir
#'
#' @return 返回值
#' @import openxlsx
#' @export
#'
#' @examples
#' TaxStatement_excel()
TaxStatement_excel <-function (erpToken = 'C0426D23-1927-4314-8736-A74B2EF7A039',FYear,FMonth,FOrgNumber,outputDir)
{


  sql_exec =paste0("EXEC rds_proc_Tax_Statement  '",FYear,"','",FMonth,"','",FOrgNumber,"'       ")
  tsda::sql_update2(token = erpToken,sql_str =sql_exec )

      #进一步处理
      meta_head = TaxStatement_meta(erpToken = erpToken )
      ncount_meta_head = nrow(meta_head)
      fields_head = paste0(meta_head$FName,collapse = " , ")
      table_head = meta_head$FTableName[1]
      sql_head = paste0("select  ",fields_head,"   from  ",table_head,"  ")

      data_head =  tsda::sql_select2(token = erpToken,sql = sql_head)

        print(3)
        #表头存在数据
          #表体存在数据，进行相应的数据处理
          #获取完整的模板文件
        # 模板文件路径（使用 file.path 确保分隔符）
        templateFile = file.path(outputDir, "报表模板.xlsx")
        print(paste("模板文件路径:", templateFile))

        # 检查模板是否存在
        if (!file.exists(templateFile)) {
          stop("模板文件不存在：", templateFile)
        }

        excel_file <- openxlsx::loadWorkbook(templateFile)

          #写入表头数据
          for ( i in 1:ncount_meta_head) {
            #针对数据处理处理
            field_head = meta_head$FName[i]
            cell_head  = meta_head$FCells[i]
            print(cell_head)
            cellData_head = as.character(data_head[1,field_head])

            print(cellData_head)
            cellIndex_head =excel_coord_to_numeric(cell_head)
            indexCol = cellIndex_head['col']
            indexRow = cellIndex_head['row']

            header_style <- createStyle(
              fontName = "Calibri",
              fontSize = 10,
              halign = "center",       # 水平居中
              valign = "center",       # 垂直居中


            )


            openxlsx::writeData(wb = excel_file, sheet = "Sheet1", x = cellData_head,
                                startCol = indexCol, startRow = indexRow,

                                headerStyle = header_style )

          }

          #处理文件名生成EXCEL
          print(5)
          # 输出文件路径（使用 file.path 修正）
          outputFile = "税务报表.xlsx"
          xlsx_file_name = file.path(outputDir, outputFile)   # 关键修正
          print(paste("输出文件路径:", xlsx_file_name))

          # 保存文件
          res = saveWorkbook(excel_file, xlsx_file_name, overwrite = TRUE)

          # 验证文件是否生成
          if (file.exists(xlsx_file_name)) {
            print("文件生成成功")
          } else {
            print("文件生成失败")
          }








  return (res)

  }






