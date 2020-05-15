package io.github.xinshepherd.excel;

import io.github.xinshepherd.excel.annotation.Excel;
import io.github.xinshepherd.excel.annotation.ExcelField;
import io.github.xinshepherd.excel.core.base.ImporterBase;
import lombok.Getter;
import lombok.Setter;
import org.junit.Assert;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @author donglin
 * @date 2020/05/15
 */
public class ImporterBaseTest {

    @Test
    public void testImport1() throws Exception {
        String filepath = getClass().getResource("/").getPath() + "/excel.xlsx";
        InputStream is = new FileInputStream(filepath);
        List<Student> students = ImporterBase.newInstance(is).resolve(Student.class);

        Assert.assertNotNull(students);
        Assert.assertEquals(3, students.size());

        SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd");

        Student student = students.get(0);
        Assert.assertEquals("张三", student.getName());
        Assert.assertEquals("男", student.getSex());
        Assert.assertEquals("1990-10-10", sdf.format(student.getDate()));
        Assert.assertEquals(29, student.getAge());
        Assert.assertEquals(90.5, student.getScore(), 0.00001);

        student = students.get(1);
        Assert.assertEquals("李四", student.getName());
        Assert.assertEquals("男", student.getSex());
        Assert.assertEquals("1999-01-01", sdf.format(student.getDate()));
        Assert.assertEquals(21, student.getAge());
        Assert.assertEquals(80.5, student.getScore(), 0.00001);

        student = students.get(2);
        Assert.assertEquals("小红", student.getName());
        Assert.assertEquals("女", student.getSex());
        Assert.assertEquals("2000-02-02", sdf.format(student.getDate()));
        Assert.assertEquals(0, student.getAge());
        Assert.assertEquals(85.5, student.getScore(), 0.00001);
    }


    @Test
    public void testImport2() throws Exception {
        String filepath = getClass().getResource("/").getPath() + "/excel.xls";
        InputStream is = new FileInputStream(filepath);
        List<Student> students = ImporterBase.newInstance(is).filename("excel.xls").titleRowIndex(2)
                .ignoreLastIndexes(1).resolve(Student.class);

        Assert.assertNotNull(students);
        Assert.assertEquals(3, students.size());

        SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd");

        Student student = students.get(0);
        Assert.assertEquals("张三", student.getName());
        Assert.assertEquals("男", student.getSex());
        Assert.assertEquals("1990-10-10", sdf.format(student.getDate()));
        Assert.assertEquals(29, student.getAge());
        Assert.assertEquals(90.5, student.getScore(), 0.00001);

        student = students.get(1);
        Assert.assertEquals("李四", student.getName());
        Assert.assertEquals("男", student.getSex());
        Assert.assertEquals("1999-01-01", sdf.format(student.getDate()));
        Assert.assertEquals(21, student.getAge());
        Assert.assertEquals(80.5, student.getScore(), 0.00001);

        student = students.get(2);
        Assert.assertEquals("小红", student.getName());
        Assert.assertEquals("女", student.getSex());
        Assert.assertEquals("2000-02-02", sdf.format(student.getDate()));
        Assert.assertEquals(0, student.getAge());
        Assert.assertEquals(85.5, student.getScore(), 0.00001);
    }

    @Getter
    @Setter
    @Excel
    public static class Student {
        @ExcelField("姓名")
        private String name;

        @ExcelField("性别")
        private String sex;

        @ExcelField("出生日期")
        private Date date;

        @ExcelField("年龄")
        private int age;

        @ExcelField("成绩")
        private double score;
    }
}
