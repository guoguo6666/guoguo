package org.sang.controller;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.sang.bean.Category;
import org.sang.bean.RespBean;
import org.sang.service.CategoryService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * 超级管理员专属Controller
 */
@RestController
@RequestMapping("/admin/category")
public class CategoryController {
    @Autowired
    CategoryService categoryService;

    @RequestMapping(value = "/all", method = RequestMethod.GET)
    public List<Category> getAllCategories() {
        return categoryService.getAllCategories();
    }

    @RequestMapping(value = "/{ids}", method = RequestMethod.DELETE)
    public RespBean deleteById(@PathVariable String ids) {
        boolean result = categoryService.deleteCategoryByIds(ids);
        if (result) {
            return new RespBean("success", "删除成功!");
        }
        return new RespBean("error", "删除失败!");
    }

    @RequestMapping(value = "/{id}", method = RequestMethod.POST)
    public Category selectById(@PathVariable String id) {
        Category category = categoryService.selectCategoryById(id);
        return category;
    }

    @RequestMapping(value = "/export"+"/{id}", method = RequestMethod.POST)
    public RespBean export(@PathVariable String id) throws IOException {
        Category category = categoryService.selectCategoryById(id);

        //创建工作薄对象
        HSSFWorkbook workbook=new HSSFWorkbook();//这里也可以设置sheet的Name
        //创建工作表对象
        HSSFSheet sheet = workbook.createSheet();
        //创建工作表的第一行
        HSSFRow row = sheet.createRow(0);//设置第一行，从零开始
        row.createCell(0).setCellValue("编号");//第一行第一列
        row.createCell(1).setCellValue("栏目名称");//第一行第二列
        row.createCell(2).setCellValue("启用时间");//第一行第三列
        workbook.setSheetName(0,"sheet的Name");//设置sheet的Name
        //创建工作表的第二行
        HSSFRow row1 = sheet.createRow(1);//设置第一行，从零开始
        row1.createCell(0).setCellValue(category.getId());//第一行第一列
        row1.createCell(1).setCellValue(category.getCateName());//第一行第二列
        row1.createCell(2).setCellValue(category.getDate());//第一行第三列
        //文档输出

        FileOutputStream out = null;
        try {
            out = new FileOutputStream("D://category" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()).toString() +".xls");
            workbook.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return new RespBean("fail", "导出失败!");
        }
        finally {
            out.close();
        }
        return new RespBean("success", "导出成功!");
    }


    @RequestMapping(value = "/", method = RequestMethod.POST)
    public RespBean addNewCate(Category category) {

        if ("".equals(category.getCateName()) || category.getCateName() == null) {
            return new RespBean("error", "请输入栏目名称!");
        }

        int result = categoryService.addCategory(category);

        if (result == 1) {
            return new RespBean("success", "添加成功!");
        }
        return new RespBean("error", "添加失败!");
    }

    @RequestMapping(value = "/", method = RequestMethod.PUT)
    public RespBean updateCate(Category category) {
        int i = categoryService.updateCategoryById(category);
        if (i == 1) {
            return new RespBean("success", "修改成功!");
        }
        return new RespBean("error", "修改失败!");
    }
}
