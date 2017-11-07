package com.thend.master.plugin.exporter;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Collection;

/**
 * 自定义构建Sheet接口
 * @author thend
 */
public interface ISheetBuilder {

    CellStyle buildTitleStyle();

    CellStyle buildDataStyle();

    CellStyle buildFootStyle();

    ISheetBuilder buildTitle(Collection<String> titles);

    <T> ISheetBuilder buildData(Collection<T> datas);

    <T> ISheetBuilder buildFoot(Collection<T> datas);

    ExportInfo getExportInfo();
}
