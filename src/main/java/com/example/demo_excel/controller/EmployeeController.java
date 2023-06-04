package com.example.demo_excel.controller;

import com.example.demo_excel.EmployeeExcelExporter;
import com.example.demo_excel.model.Employee;
import com.example.demo_excel.service.EmployeeServiceImpl;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@RestController
public class EmployeeController {
    @Autowired
    private EmployeeServiceImpl employeeServiceImpl;


    @GetMapping("export-to-excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=users_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        List<Employee> employeeList = employeeServiceImpl.listAll();

        EmployeeExcelExporter excelExporter = new EmployeeExcelExporter(employeeList);

        excelExporter.export(response);
    }
}
