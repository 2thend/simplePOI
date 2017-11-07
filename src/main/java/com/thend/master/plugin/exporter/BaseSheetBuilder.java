package com.thend.master.plugin.exporter;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;

import java.util.Collection;
import java.util.Iterator;

/**
 * 自定义构建POI Sheet 基类
 * @author thend
 */
public class BaseSheetBuilder implements ISheetBuilder {

    private ExportInfo info;

    public BaseSheetBuilder(String sheetName){
        info = new ExportInfo(sheetName);
        Sheet localSheet = ExportExcelUtils.getLocalWorkbook().createSheet();
        info.setLocalSheet(localSheet);
        int index = ExportExcelUtils.getLocalWorkbook().getSheetIndex(localSheet);
        ExportExcelUtils.getLocalWorkbook().setSheetName(index, sheetName);
    }
    public <T> BaseSheetBuilder(String sheetName, Collection<String> titles, Collection<T> datas){
        info = new ExportInfo(sheetName, titles, datas);
    }
    public <T, K> BaseSheetBuilder(String sheetName, Collection<String> titles, Collection<T> datas, Collection<K> footDatas){
        info = new ExportInfo(sheetName, titles, datas, footDatas);
    }

    @Override
    public ExportInfo getExportInfo(){
        return info;
    }

    @Override
    public ISheetBuilder buildTitle(Collection<String> titles){
        if(titles == null || titles.size() == 0){
            return this;
        }
        Row headRow = info.getLocalSheet().createRow(info.getCurRowIndex());
        int colIndex = 0;
        for(Iterator list = titles.iterator(); list.hasNext();){
            Cell cellHead = headRow.createCell(colIndex);
            cellHead.setCellValue(list.next().toString());
            cellHead.setCellStyle(buildTitleStyle());
            colIndex++;
        }
        info.setCurRowIndex(info.getCurRowIndex() + 1);
        return this;
    }

    @Override
    public <T> ISheetBuilder buildData(Collection<T> datas) {
        if(datas == null || datas.size() == 0){
            return this;
        }
        Iterable curData = null;
        if(datas instanceof Collection){
            curData = (Collection) datas;
            Iterator rowList = curData.iterator();
            while(rowList.hasNext()){
                Iterable colList = (Iterable) rowList.next();
                if(colList == null){
                    return this;
                }
//            Row row = localSheet.createRow(dataIndex - ExcelExportUtil_bak.EXCEL_MAX_ROWNUM * curSheetIndex + 1);
                Row row = info.getLocalSheet().createRow(info.getCurRowIndex());
                int colIndex = 0;
                Iterator colIterator = colList.iterator();
                while(colIterator.hasNext()){
                    Object rowData = colIterator.next();
                    Cell cell = row.createCell(colIndex);
                    cell.setCellStyle(buildDataStyle());
                    cell.setCellValue(rowData.toString());
                    colIndex++;
                }
                info.setCurRowIndex(info.getCurRowIndex() + 1);
            }
        }
        return this;
    }

    @Override
    public <T> ISheetBuilder buildFoot(Collection<T> datas) {
        if(datas == null || datas.size() == 0){
            return this;
        }
        Row footRow = info.getLocalSheet().createRow(info.getCurRowIndex());
        int index = 0;
        for(Iterator list = datas.iterator(); list.hasNext();){
            Cell cellHead = footRow.createCell(index);
            cellHead.setCellValue(list.next().toString());
            cellHead.setCellStyle(buildTitleStyle());
            index++;
        }
        info.setCurRowIndex(1 + info.getCurRowIndex());
        return this;
    }

    @Override
    public CellStyle buildTitleStyle() {
        return buildSpecialStyle(ExportExcelUtils.getLocalWorkbook());
    }

    @Override
    public CellStyle buildDataStyle() {
        CellStyle titleStyle = ExportExcelUtils.getLocalWorkbook().createCellStyle();
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font font = ExportExcelUtils.getLocalWorkbook().createFont();
        titleStyle.setFont(font);
        return titleStyle;
    }

    @Override
    public CellStyle buildFootStyle() {
        return buildSpecialStyle(ExportExcelUtils.getLocalWorkbook());
    }


    private CellStyle buildSpecialStyle(Workbook workbook){
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        titleStyle.setFont(font);
        return titleStyle;
    }
}
