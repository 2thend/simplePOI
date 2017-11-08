package com.thend.master.controller;

import com.thend.master.exporter.CustomSheetBuilder;
import com.thend.master.model.Employee;
import com.thend.master.plugin.exporter.BaseSheetBuilder;
import com.thend.master.plugin.exporter.ExportExcelUtils;
import com.thend.master.plugin.exporter.ISheetBuilder;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletResponse;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

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
        List<String> titleList = new ArrayList<>();
        titleList.add("文字aaa");
        titleList.add("文字bbb");
        titleList.add("文字ccc");
        titleList.add("文字ddd");
        titleList.add("文字eee");

        List<List<String>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++){
            List<String> temps = new ArrayList<>();
            temps.add("信息a" + i);
            temps.add("信息b" + i);
            temps.add("信息c" + i);
            temps.add("信息d" + i);
            temps.add("信息e" + i);
            dataList.add(temps);
        }
        String now = new SimpleDateFormat("yyyy-MM-dd_hh-mm-ss").format(new Date());
        String exportFileName = "export测试_" + now + ".xlsx";

        //单Sheet文件导出 通用
        ExportExcelUtils.export(exportFileName, "sheet测试", titleList, dataList, null, response);
    }

    @RequestMapping(value = "/customExport", method = RequestMethod.GET)
    public void customExport(HttpServletResponse response) {
        List<Employee> empList = buildEmployeeList();
        String now = new SimpleDateFormat("yyyy-MM-dd_hh-mm-ss").format(new Date());
        String exportFileName = "export测试_" + now + ".xlsx";
        //单Sheet文件导出 定制
        ISheetBuilder builder = new CustomSheetBuilder("sheet测试", null, empList);
        ExportExcelUtils.export(exportFileName, builder, response);
    }

    @RequestMapping(value = "/multiExport", method = RequestMethod.GET)
    public void multiExport(HttpServletResponse response) {
        List<String> titleList = new ArrayList<>();
        titleList.add("标题aaa");
        titleList.add("标题bbb");
        titleList.add("标题ccc");
        titleList.add("标题ddd");
        titleList.add("标题eee");

        List<List<String>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++){
            List<String> temps = new ArrayList<>();
            temps.add("信息aaa" + i);
            temps.add("信息bbb" + i);
            temps.add("信息ccc" + i);
            temps.add("信息ddd" + i);
            temps.add("信息eee" + i);
            dataList.add(temps);
        }

        List<Employee> empList = buildEmployeeList();
        String now = new SimpleDateFormat("yyyy-MM-dd_hh-mm-ss").format(new Date());
        String exportFileName = "export测试_" + now + ".xlsx";
        ISheetBuilder builder1 = new BaseSheetBuilder("普通Sheet", titleList, dataList);
        ISheetBuilder builder2 = new CustomSheetBuilder("定制Sheet", null, empList);
        //多Sheet文件导出 定制
        List<ISheetBuilder> builderList = new ArrayList<>();
        builderList.add(builder1);
        builderList.add(builder2);
        ExportExcelUtils.export(exportFileName, builderList, response);
    }

    private List<Employee> buildEmployeeList(){
        List<Employee> empList = new ArrayList<>();
        Random random = new Random();
        for(int i = 0; i < 10; i++){
            Employee employee = new Employee();
            employee.setEmpName("测试" + i);
            employee.setEmpCode("P0000" + i);
            employee.setDepartId(i);
            employee.setLeaderId(100+i);
            employee.setMobile("1311234567" + i);
            employee.setAddress("长安街 00" + i + "号");
            employee.setSalary(new BigDecimal(random.nextDouble() * 50000d));
            empList.add(employee);
        }
        return empList;
    }
}
