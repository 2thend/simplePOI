package com.thend.master.plugin.exporter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

/**
 * 导出Excel插件
 * @author thend
 */
public class ExportExcelUtils {

    private static final Logger log = LoggerFactory.getLogger(ExportExcelUtils.class);

    private static ThreadLocal<Workbook> localWorkbook = new ThreadLocal<>();

    /** 导出excel单个sheet最大行数 */
    public static int EXCEL_MAX_ROWNUM = 50000;
    /** 导出excel内存缓存最大行数 */
    public static int EXCEL_CACHE_ROWNUM = 100;

    /**
     * 单Sheet的Excel导出
     * @param <T> 表数据泛型
     * @param <K> 表尾数据泛型
     * @param exportFile 出文件名
     * @param sheetName
     * @param titles 表头
     * @param datas 表数据
     * @param footDatas 表尾数据
     * @param response
     */
    public static <T extends Iterable, K extends Iterable> void export(String exportFile, String sheetName, Collection<String> titles, Collection<T> datas, Collection<K> footDatas, HttpServletResponse response){
        ISheetBuilder sheetBuilder = new BaseSheetBuilder(sheetName).buildTitle(titles).buildData(datas);
        if(footDatas != null){
            sheetBuilder.buildFoot(footDatas);
        }
        //导出excel文件
        outputExcel(exportFile, response);
    }
    public static <T extends Iterable, K extends Iterable> void export(String exportFile, String sheetName, Collection<String> titles, Collection<T> datas, Collection<K> footDatas){
        export(exportFile, sheetName, titles, datas, footDatas, null);
    }

    /**
     * 单Sheet的Excel导出 个性化定制
     * @param exportFile
     * @param sheetBuilder 传入sheetBuilder需初始化sheetName, titles, datas[, footDatas]
     * @param response
     */
    public static void export(String exportFile, ISheetBuilder sheetBuilder, HttpServletResponse response){
        if(sheetBuilder == null){
            return;
        }
        //定制Sheet
        buildSheet(sheetBuilder);
        //导出excel文件
        outputExcel(exportFile, response);
    }
    public static void export(String exportFile, ISheetBuilder sheetBuilder){
        export(exportFile, sheetBuilder, null);
    }

    /**
     * 多Sheet的Excel导出 个性化定制
     * @param exportFile
     * @param sheetBuilderList
     * @param response
     */
    public static void export(String exportFile, List<ISheetBuilder> sheetBuilderList, HttpServletResponse response){
        if(sheetBuilderList == null || sheetBuilderList.size() == 0){
            return;
        }
        for (ISheetBuilder builder : sheetBuilderList) {
            //定制Sheet
            buildSheet(builder);

        }
        //导出excel文件
        outputExcel(exportFile, response);
    }
    public static void export(String exportFile, List<ISheetBuilder> sheetBuilderList){
        export(exportFile, sheetBuilderList, null);
    }


    //输出Excel
    public static void outputExcel(String fileName, HttpServletResponse response) {
        if(response == null){
//            response = HttpKit.getResponse();
        }
        OutputStream ouputStream = null;
        try {
            // 设定输出文件头
            ouputStream =  response.getOutputStream();
            // 清空输出流
            response.reset();
            response.setHeader("Content-disposition", "attachment; filename=" + new String(fileName.getBytes(),"iso-8859-1"));
            response.setContentType("application/msexcel");// 定义输出类型
            localWorkbook.get().write(ouputStream);
//            System.out.println("====== >> use time = " + (endTime - startTime) + "ms");
            ouputStream.flush();
            ouputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                ouputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    public static void outputExcel(String fileName) {
        outputExcel(fileName, null);
    }

    public static Workbook getLocalWorkbook() {
        // 获取本地线程变量并强制转换为Task类型
        Workbook workbook = localWorkbook.get();
        // 线程首次执行此方法的时候，TaskLocal.get()肯定为null
        if (workbook == null) {
            // 创建一个Task对象，并保存到本地线程变量TaskLocal中
            workbook = new SXSSFWorkbook(EXCEL_CACHE_ROWNUM);
            localWorkbook.set(workbook);
        }
        return workbook;
    }

    /**
     * 辅助方法：定制Sheet
     * @param sheetBuilder
     */
    private static void buildSheet(ISheetBuilder sheetBuilder){
        Sheet localSheet = ExportExcelUtils.getLocalWorkbook().createSheet();
        ExportInfo info = sheetBuilder.getExportInfo();
        if(info == null){
            return;
        }
        info.setLocalSheet(localSheet);
        int index = ExportExcelUtils.getLocalWorkbook().getSheetIndex(localSheet);
        String sheetName = info.getSheetName() == null || "".equals(info.getSheetName()) ? ExportInfo.DEFAULT_SHEETNAME : info.getSheetName();
        ExportExcelUtils.getLocalWorkbook().setSheetName(index, sheetName);
        sheetBuilder.buildTitle(info.getTitles()).buildData(info.getDatas()).buildFoot(info.getFootDatas());
    }
}
