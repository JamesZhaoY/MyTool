package org.JamesZhao.Excel;


import cn.hutool.core.util.StrUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;

public class ExcelTest {
    private static final Log log = LogFactory.get(ExcelTest.class);

    /**
     * cellName在于借助poi的formula函数，给当前单元格设置函数，并利用计算器强制刷新计算出结果
     * @param excelStr 函数表达式
     * @param sheetName sheet页名称
     * @param filePath excel文件地址
     * @param cellName 需要用来计算的单元格
     * @return 函数表达式计算出来的结果
     */
    public BigDecimal excelMath(String excelStr, String sheetName, String filePath, String cellName)  {
        try{
            Sheet sheet =null;
            File file=new File(filePath);
            Cell cell;
            FormulaEvaluator evaluator;
            try (Workbook workbook = WorkbookFactory.create(file)) {
                if (StrUtil.isBlank(sheetName)) {
                    sheet = workbook.getSheetAt(0);
                } else {
                    sheet = workbook.getSheetAt(workbook.getSheetIndex(sheetName));
                }
                CellReference cellReference = new CellReference(cellName);
                Row row = sheet.getRow(cellReference.getRow());
                if (row == null) {
                    row = sheet.createRow(cellReference.getRow());
                }
                cell = row.getCell(cellReference.getCol());
                if (cell == null) {
                    cell = row.createCell(cellReference.getCol());
                }
//            新版本poi中已弃用该方法，根据cellValue自动设置单元格类型
//            cell.setCellType(CellType.FORMULA);
                cell.setCellFormula(excelStr);
                sheet.setForceFormulaRecalculation(true);
                evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            }
            CellValue evaluate = evaluator.evaluate(cell);
            if (evaluate.getCellType() == CellType.NUMERIC) {
                return BigDecimal.valueOf(evaluate.getNumberValue());
            }else {
                throw new RuntimeException("函数中存在非法值");
            }
        }catch (IOException e){
            log.error("文件读取异常");
        }catch (Exception e){
            log.error("函数计算失败："+e.getMessage());
        }
        return null;
}
}
