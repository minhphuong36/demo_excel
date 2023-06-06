package com.example.demo_excel.service;

import com.example.demo_excel.helper.ExcelHelper;
import com.example.demo_excel.model.Employee;
import com.example.demo_excel.repository.EmployeeRepository;
import jakarta.transaction.Transactional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@Service
@Transactional
public class EmployeeServiceImpl implements EmployeeService{

   @Autowired
   private EmployeeRepository employeeRepository;

    public List<Employee> listAll() {
        return employeeRepository.findAll(Sort.by("name").ascending());
    }

    public void save(MultipartFile file) {
        try {
            List<Employee> employeeList = ExcelHelper.excelToTutorials(file.getInputStream());
            employeeRepository.saveAll(employeeList);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }

    public List<Employee> getAllEmployees() {
        return employeeRepository.findAll();
    }
}

