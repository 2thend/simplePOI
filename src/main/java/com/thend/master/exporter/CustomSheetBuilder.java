package com.thend.master.exporter;

import com.thend.master.model.Employee;
import com.thend.master.plugin.exporter.BaseSheetBuilder;
import com.thend.master.plugin.exporter.ExportExcelUtils;
import com.thend.master.plugin.exporter.ExportInfo;
import com.thend.master.plugin.exporter.ISheetBuilder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;

/**
 * 定制SheetBuilder
 * @author thend
 */
public class CustomSheetBuilder implements ISheetBuilder {

    private BaseSheetBuilder builder;
    private ExportInfo info;

    public CustomSheetBuilder(String sheetName){
        Sheet localSheet = ExportExcelUtils.getLocalWorkbook().createSheet();
        int index = ExportExcelUtils.getLocalWorkbook().getSheetIndex(localSheet);
        ExportExcelUtils.getLocalWorkbook().setSheetName(index, sheetName);
    }
    public <T> CustomSheetBuilder(String sheetName, Collection<String> titles, Collection<T> datas){
        info = new ExportInfo(sheetName, titles, datas);
        builder = new BaseSheetBuilder(sheetName, titles, datas);
    }
    public <T, K> CustomSheetBuilder(String sheetName, Collection<String> titles, Collection<T> datas, Collection<K> footDatas){
        info = new ExportInfo(sheetName, titles, datas, footDatas);
        builder = new BaseSheetBuilder(sheetName, titles, datas, footDatas);
    }

    @Override
    public ISheetBuilder buildTitle(Collection<String> titles) {
        Sheet localSheet = info.getLocalSheet();
        localSheet.setColumnWidth(0, 45 * 256);
        localSheet.setColumnWidth(1, 28 * 256);
        localSheet.setColumnWidth(2, 20 * 256);
        localSheet.setColumnWidth(3, 15 * 256);
        localSheet.setColumnWidth(4, 15 * 256);
        localSheet.setColumnWidth(5, 18 * 256);
        localSheet.setColumnWidth(6, 18 * 256);
        localSheet.setColumnWidth(7, 15 * 256);

        CellStyle headStyle = buildTitleStyle();
        Row localSheetheaderRow = localSheet.createRow(0);
        localSheetheaderRow.createCell(0).setCellValue("员工姓名");
        localSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        localSheetheaderRow.getCell(0).setCellStyle(headStyle);

        localSheetheaderRow.createCell(1).setCellValue("员工编号");
        localSheetheaderRow.getCell(1).setCellStyle(headStyle);
        localSheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));

        localSheetheaderRow.createCell(2).setCellValue("部门信息");
        localSheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 3));
        localSheetheaderRow.getCell(2).setCellStyle(headStyle);

        localSheetheaderRow.createCell(4).setCellValue("个人信息");
        localSheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 6));
        localSheetheaderRow.getCell(4).setCellStyle(headStyle);


        Row localSheetsecondRow = localSheet.createRow(1);
        localSheetsecondRow.createCell(2).setCellValue("部门ID");
        localSheetsecondRow.getCell(2).setCellStyle(headStyle);
        localSheetsecondRow.createCell(3).setCellValue("领导ID");
        localSheetsecondRow.getCell(3).setCellStyle(headStyle);
        localSheetsecondRow.createCell(4).setCellValue("手机号");
        localSheetsecondRow.getCell(4).setCellStyle(headStyle);
        localSheetsecondRow.createCell(5).setCellValue("住址");
        localSheetsecondRow.getCell(5).setCellStyle(headStyle);
        localSheetsecondRow.createCell(6).setCellValue("薪资");
        localSheetsecondRow.getCell(6).setCellStyle(headStyle);
        info.setCurRowIndex(2);
        return this;
    }

    @Override
    public <T> ISheetBuilder buildData(Collection<T> datas) {
        if (datas == null) {
            return this;
        }

        for (T rowData : datas) {
            CellStyle celltyle = buildDataStyle();
            Sheet localSheet = info.getLocalSheet();
            Employee item = (Employee) rowData;
            Row retailerrow = localSheet.createRow(info.getCurRowIndex());
            retailerrow.createCell(0).setCellValue(item.getEmpName());
            retailerrow.getCell(0).setCellStyle(celltyle);
            retailerrow.createCell(1).setCellValue(item.getEmpCode());
            retailerrow.getCell(1).setCellStyle(celltyle);
            retailerrow.createCell(2).setCellValue(item.getDepartId());
            retailerrow.getCell(2).setCellStyle(celltyle);
            retailerrow.createCell(3).setCellValue(item.getLeaderId());
            retailerrow.getCell(3).setCellStyle(celltyle);
            retailerrow.createCell(4).setCellValue(item.getMobile());
            retailerrow.getCell(4).setCellStyle(celltyle);
            retailerrow.createCell(5).setCellValue(item.getAddress());
            retailerrow.getCell(5).setCellStyle(celltyle);
            retailerrow.createCell(6).setCellValue(item.getSalary().doubleValue());
            retailerrow.getCell(6).setCellStyle(celltyle);
            info.setCurRowIndex(info.getCurRowIndex() + 1);
        }
        return this;
    }

    @Override
    public <T> ISheetBuilder buildFoot(Collection<T> datas) {
        return builder.buildFoot(datas);
    }

    @Override
    public CellStyle buildTitleStyle() {
        return builder.buildTitleStyle();
    }

    @Override
    public CellStyle buildDataStyle() {
        return builder.buildDataStyle();
    }

    @Override
    public CellStyle buildFootStyle() {
        return builder.buildFootStyle();
    }

    @Override
    public ExportInfo getExportInfo() {
        return info;
    }
}
