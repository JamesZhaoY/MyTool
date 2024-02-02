package org.JamesZhao.Excel;


import cn.hutool.core.util.StrUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {
    private static final Log log = LogFactory.get(ExcelUtil.class);

    /**
     * cellName在于借助poi的formula函数，给当前单元格设置函数，并利用计算器强制刷新计算出结果
     *
     * @param excelStr  函数表达式
     * @param sheetName sheet页名称
     * @param filePath  excel文件地址
     * @param cellName  需要用来计算的单元格
     * @return 函数表达式计算出来的结果
     */
    public BigDecimal excelMath(String excelStr, String sheetName, String filePath, String cellName) {
        try {
            Sheet sheet = null;
            File file = new File(filePath);
            Cell cell = null;
            FormulaEvaluator evaluator;
            try (Workbook workbook = WorkbookFactory.create(file)) {
                if (StrUtil.isBlank(sheetName)) {
                    sheet = workbook.getSheetAt(0);
                } else {
                    sheet = workbook.getSheetAt(workbook.getSheetIndex(sheetName));
                }
                CellReference cellReference = new CellReference(cellName);
                cell = getCell(sheet, cellReference, cell);
//            新版本poi中已弃用该方法，根据cellValue自动设置单元格类型
//            cell.setCellType(CellType.FORMULA);
                cell.setCellFormula(excelStr);
                sheet.setForceFormulaRecalculation(true);
                evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            }
            CellValue evaluate = evaluator.evaluate(cell);
            if (evaluate.getCellType() == CellType.NUMERIC) {
                return BigDecimal.valueOf(evaluate.getNumberValue());
            } else {
                throw new RuntimeException("函数中存在非法值");
            }
        } catch (IOException e) {
            log.error("文件读取异常");
        } catch (Exception e) {
            log.error("函数计算失败：" + e.getMessage());
        }
        return null;
    }


    /**
     * 使用刷新器将函数写入excel，并且调用FormulaParser来解析出函数中包含的单元格
     *
     * @param cellFormula 函数
     * @param cellName 设置函数的cell
     * @param filePath excel文件地址
     * @return 函数中包含的所有单元格
     */
    public static List<CellReference> getCellFormulaCellReferences(String cellFormula, String cellName, String filePath) {
        CellReference cellReference = new CellReference(cellName);
        try {
            FormulaParsingWorkbook evaluator = null;
            Sheet sheet = null;
            Cell cell = null;
            FormulaEvaluator formulaEvaluator;
            File file = new File(filePath);
            try (Workbook workbook = WorkbookFactory.create(file)) {
                if (file.getName().endsWith(".xls")) {
                    evaluator = HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook);
                } else if (file.getName().endsWith(".xlsx")) {
                    evaluator = XSSFEvaluationWorkbook.create((XSSFWorkbook) workbook);
                }
                sheet = workbook.getSheetAt(0);
                cell = getCell(sheet, cellReference, cell);
                cell.setCellFormula(cellFormula);
                sheet.setForceFormulaRecalculation(true);
                formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                formulaEvaluator.evaluateAll();
                String formula = cell.getCellFormula();
                List<CellReference> cellReferences = new ArrayList<>();
                Ptg[] parse = FormulaParser.parse(formula, evaluator, FormulaType.CELL, 0);
                for (Ptg ptg : parse) {
                    if (ptg instanceof Ref3DPxg) {
                        Ref3DPxg pxg = (Ref3DPxg) ptg;
                        String sheetName = pxg.getSheetName();
                        CellReference reference = new CellReference(sheetName, pxg.getRow(), pxg.getColumn(), Boolean.FALSE, Boolean.FALSE);
                        cellReferences.add(reference);
                    } else if (ptg instanceof RefPtgBase) {
                        RefPtgBase ptgBase = (RefPtgBase) ptg;
                        CellReference reference = new CellReference(ptgBase.getRow(), ptgBase.getColumn());
                        cellReferences.add(reference);
                    } else if (ptg instanceof AreaPtgBase) {
                        AreaPtgBase ptgBase = (AreaPtgBase) ptg;
                        CellReference first = new CellReference(ptgBase.getFirstRow(), ptgBase.getFirstColumn());
                        CellReference last = new CellReference(ptgBase.getLastRow(), ptgBase.getLastColumn());
                        for (int y = first.getRow(); y <= last.getRow(); y++) {
                            for (int x = first.getCol(); x <= last.getCol(); x++) {
                                cellReferences.add(new CellReference(y, x));
                            }
                        }
                    }
                }
                return cellReferences;
            }
        } catch (IOException e) {
            log.error("文件读取异常");
        }
        return new ArrayList<>();
    }

    private static Cell getCell(Sheet sheet, CellReference cellReference, Cell cell) {
        Row row = sheet.getRow(cellReference.getRow());
        if (row == null) {
            row = sheet.createRow(cellReference.getRow());
        }
        cell = row.getCell(cellReference.getCol());
        if (cell == null) {
            cell = row.createCell(cellReference.getCol());
        }
        return cell;
    }

}
