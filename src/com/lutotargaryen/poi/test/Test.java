package com.lutotargaryen.poi.test;

import java.util.List;
import java.util.Map;

import com.lutotargaryen.poi.exception.RepeatCreateObject;
import com.lutotargaryen.poi.readexcel.ReadExcel;

public class Test {
	public static void main(String[] args) throws RepeatCreateObject {
		ReadExcel readExcel = ReadExcel.newInstance("com.mysql.jdbc.Driver","jdbc:mysql://127.0.0.1:3306/readexcel?userUnicode=true&characterEncoding=UTF-8");
		readExcel = ReadExcel.newInstance("com.mysql.jdbc.Driver","jdbc:mysql://127.0.0.1:3306/readexcel?userUnicode=true&characterEncoding=UTF-8","root","root");
		List<Map<String,Object>> list = readExcel.queryTest();
		System.out.println(list);
	}
}
