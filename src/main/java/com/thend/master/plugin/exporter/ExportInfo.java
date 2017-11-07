package com.thend.master.plugin.exporter;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.Collection;

/**
 * 导出Excel数据实体
 * @author thend
 * @param <T> 导出 Excel data 泛型
 * @param <K> 导出 Excel footData 泛型
 */
public class ExportInfo<T, K>{

    public static final String DEFAULT_SHEETNAME = "export";

    protected Sheet localSheet;
    private String sheetName;
    private Collection<String> titles = null;
    private Collection<T> datas = null;
    private Collection<K> footDatas = null;
    private int curRowIndex;

    public ExportInfo(String sheetName){
        this.sheetName = sheetName;
    }
    public ExportInfo(String sheetName, Collection<String> titles, Collection<T> datas){
        this.sheetName = sheetName;
        this.titles = titles;
        this.datas = datas;
    }
    public ExportInfo(String sheetName, Collection<String> titles, Collection<T> datas, Collection<K> footDatas){
        this(sheetName, titles, datas);
        this.footDatas = footDatas;
    }


    public Sheet getLocalSheet() {
        return localSheet;
    }

    public void setLocalSheet(Sheet localSheet) {
        this.localSheet = localSheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Collection<String> getTitles() {
        return titles;
    }

    public void setTitles(Collection<String> titles) {
        this.titles = titles;
    }

    public Collection<T> getDatas() {
        return datas;
    }

    public void setDatas(Collection<T> datas) {
        this.datas = datas;
    }

    public Collection<K> getFootDatas() {
        return footDatas;
    }

    public void setFootDatas(Collection<K> footDatas) {
        this.footDatas = footDatas;
    }

    public int getCurRowIndex() {
        return curRowIndex;
    }

    public void setCurRowIndex(int curRowIndex) {
        this.curRowIndex = curRowIndex;
    }
}
