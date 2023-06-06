package com.example.demo_excel.helper;

import com.example.demo_excel.model.Department;
import com.example.demo_excel.model.Employee;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;



public class ExcelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = { "id", "name", "address", "salary","department_id"};
    static String SHEET = "employee";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }

    public static List<Employee> excelToTutorials(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);
            Iterator<Row> rows = sheet.iterator();

            List<Employee> employeeList = new ArrayList<Employee>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Employee employee = new Employee();
                Department department = new Department();

                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                        case 0:
                            employee.setId((int) currentCell.getNumericCellValue());
                            break;

                        case 1:
                            employee.setName(currentCell.getStringCellValue());
                            break;

                        case 2:
                            employee.setAddress(currentCell.getStringCellValue());
                            break;

                        case 3:
                            employee.setSalary((int) currentCell.getNumericCellValue());
                            break;
//                        case 4:
//                            if(department==null){
//
//                            }
//                            break;


                        default:
                            break;
                    }

                    cellIdx++;
                }

                employeeList.add(employee);
            }

            workbook.close();

            return employeeList;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}
