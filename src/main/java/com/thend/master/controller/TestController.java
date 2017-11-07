package com.thend.master.controller;

import com.thend.master.plugin.exporter.BaseSheetBuilder;
import com.thend.master.plugin.exporter.ExportExcelUtils;
import com.thend.master.plugin.exporter.ISheetBuilder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 测试controller
 * @author thend
 */
@Controller
public class TestController {

    @RequestMapping(value = "/")
    public String index() {
        return "index";
    }

    @RequestMapping(value = "/simpleExport", method = RequestMethod.GET)
    public void simpleExport(HttpServletResponse response) {
        Map<String, Object> exportData = buildExportData();
        List<String> titleList = (List<String>)exportData.get("titleList");
        List<List<String>> dataList = (List<List<String>>)exportData.get("dataList");
        //单Sheet文件导出 通用
        String exportFileName = exportData.get("exportFileName").toString();
        ExportExcelUtils.export(exportFileName, "sheet测试", titleList, dataList, null, response);
    }

    @RequestMapping(value = "/customExport", method = RequestMethod.GET)
    public void customExport(HttpServletResponse response) {
        Map<String, Object> exportData = buildExportData();
        List<String> titleList = (List<String>)exportData.get("titleList");
        List<List<String>> dataList = (List<List<String>>)exportData.get("dataList");
        String exportFileName = exportData.get("exportFileName").toString();
        //单Sheet文件导出 定制
        ISheetBuilder builder = new BaseSheetBuilder("sheet测试", titleList, dataList);
        ExportExcelUtils.export(exportFileName, builder, response);
    }

    @RequestMapping(value = "/multiExport", method = RequestMethod.GET)
    public void multiExport(HttpServletResponse response) {
        Map<String, Object> exportData = buildExportData();
        List<String> titleList = (List<String>)exportData.get("titleList");
        List<List<String>> dataList = (List<List<String>>)exportData.get("dataList");
        String exportFileName = exportData.get("exportFileName").toString();
        ISheetBuilder builder1 = new BaseSheetBuilder("sheet测试1", titleList, dataList);
        ISheetBuilder builder2 = new BaseSheetBuilder("sheet测试2", titleList, dataList);
        ISheetBuilder builder3 = new BaseSheetBuilder("sheet测试3", titleList, dataList);
        //多Sheet文件导出 定制
        List<ISheetBuilder> builderList = new ArrayList<>();
        builderList.add(builder1);
        builderList.add(builder2);
        builderList.add(builder3);
        ExportExcelUtils.export(exportFileName, builderList, response);
    }

    private Map<String, Object> buildExportData(){
        Map<String, Object> exportData = new HashMap();
        List<String> titleList = new ArrayList<>();
        titleList.add("aaa");
        titleList.add("bbb");
        titleList.add("ccc");
        titleList.add("ddd");
        titleList.add("eee");

        List<List<String>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++){
            List<String> temps = new ArrayList<>();
            temps.add("a" + i);
            temps.add("b" + i);
            temps.add("c" + i);
            temps.add("e" + i);
            temps.add("f" + i);
            dataList.add(temps);
        }
        String now = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(new Date());
        String exportFileName = "export测试_" + now + ".xlsx";
//        String exportFileName = "export_" + now;
        exportData.put("titleList", titleList);
        exportData.put("dataList", dataList);
        exportData.put("footDataList", null);
        exportData.put("exportFileName", exportFileName);
        return exportData;
    }
}
