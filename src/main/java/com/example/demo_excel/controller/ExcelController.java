package com.example.demo_excel.controller;

import com.example.demo_excel.helper.ExcelHelper;
import com.example.demo_excel.message.ResponseMessage;
import com.example.demo_excel.model.Employee;
import com.example.demo_excel.service.EmployeeServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

//@CrossOrigin("http://localhost:8080")
@RestController
public class ExcelController {
    @Autowired
    EmployeeServiceImpl employeeService;




    @PostMapping("/upload")
    public ResponseEntity<ResponseMessage> uploadFile(@RequestParam("file") MultipartFile file) {
        String message = "";

        if (ExcelHelper.hasExcelFormat(file)) {
            try {
                employeeService.save(file);

                message = "Uploaded the file successfully: " + file.getOriginalFilename();
                return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
            } catch (Exception e) {
                message = "Could not upload the file: " + file.getOriginalFilename() + "!";
                return ResponseEntity.status(HttpStatus.EXPECTATION_FAILED).body(new ResponseMessage(message));
            }
        }

        message = "Please upload an excel file!";
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
    }

    @GetMapping("/employee")
    public ResponseEntity<List<Employee>> getAllTutorials() {
        try {
            List<Employee> employeeList = employeeService.getAllEmployees();

            if (employeeList.isEmpty()) {
                return new ResponseEntity<>(HttpStatus.NO_CONTENT);
            }

            return new ResponseEntity<>(employeeList, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
