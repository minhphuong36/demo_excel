package com.example.demo_excel.service;

import com.example.demo_excel.dto.EmployeeDTO;
import com.example.demo_excel.helper.ExcelHelper;
import com.example.demo_excel.model.Employee;
import com.example.demo_excel.repository.DepartmentRepository;
import com.example.demo_excel.repository.EmployeeRepository;
import jakarta.transaction.Transactional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

@Service
@Transactional
public class EmployeeServiceImpl implements EmployeeService{

   @Autowired
   private EmployeeRepository employeeRepository;

   @Autowired
    DepartmentRepository departmentRepository;

    public List<Employee> listAll() {
        return employeeRepository.findAll(Sort.by("name").ascending());
    }

    public void save(MultipartFile file) {
        try {
            List<EmployeeDTO> employeeDTOList = ExcelHelper.excelToEmployees(file.getInputStream());
            List<Employee> employees = employeeDTOList.stream().map( e ->{
                Employee employee = new Employee();
                employee.setName(e.getName());
                employee.setAddress(e.getAddress());
                employee.setSalary(e.getSalary());
                employee.setDepartment(departmentRepository.findByName(e.getDepartmentName()));
                return employee;
            }).collect(Collectors.toList());
            employeeRepository.saveAll(employees);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }

    public List<Employee> getAllEmployees() {
        return employeeRepository.findAll();
    }
}

