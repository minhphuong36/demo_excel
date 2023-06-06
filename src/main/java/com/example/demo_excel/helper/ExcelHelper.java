package com.example.demo_excel.helper;

import com.example.demo_excel.dto.EmployeeDTO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

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

    public static List<EmployeeDTO> excelToEmployees(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);
            Iterator<Row> rows = sheet.iterator();

            List<EmployeeDTO> employeeList = new ArrayList<EmployeeDTO>();
            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                EmployeeDTO employeeDTO = new EmployeeDTO();


                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                        case 1:
                            employeeDTO.setName(currentCell.getStringCellValue());
                            break;

                        case 2:
                            employeeDTO.setAddress(currentCell.getStringCellValue());
                            break;

                        case 3:
                            employeeDTO.setSalary((int) currentCell.getNumericCellValue());
                            break;
                        case 4:
                            employeeDTO.setDepartmentName(currentCell.getStringCellValue());
                            break;

                        default:
                            break;
                    }

                    cellIdx++;
                }

                employeeList.add(employeeDTO);
            }

            workbook.close();

            return employeeList;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}
