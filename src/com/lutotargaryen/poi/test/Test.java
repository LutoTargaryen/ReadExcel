package com.lutotargaryen.poi.test;

import java.io.IOException;

import com.lutotargaryen.poi.exception.RepeatCreateObject;
import com.lutotargaryen.poi.readexcel.ReadExcel;

public class Test {
	public static void main(String[] args) throws RepeatCreateObject, IOException {
		ReadExcel readExcel = ReadExcel.newInstance("com.mysql.jdbc.Driver","jdbc:mysql://127.0.0.1:3306/readexcel?userUnicode=true&characterEncoding=UTF-8");
		
		int s = readExcel.readExcelToMysql("D://ChoiceInfo.XLS");
		System.out.println(s);
	
	}
}
