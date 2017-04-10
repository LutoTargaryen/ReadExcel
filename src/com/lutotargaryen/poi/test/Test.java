package com.lutotargaryen.poi.test;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import com.lutotargaryen.poi.exception.RepeatCreateObject;
import com.lutotargaryen.poi.readexcel.ReadExcel;


public class Test {
	public static void main(String[] args){
		try {
			Class.forName("com.mysql.jdbc.Driver");
			Connection connection = DriverManager.getConnection("jdbc:mysql://127.0.0.1:3306/readexcel?userUnicode=true&characterEncoding=UTF-8","root","root");
			ReadExcel readExcel = ReadExcel.newInstance(connection);
			int i = readExcel.readExcelToMysql("D://ChoiceInfo.XLS");
			System.out.println(i);
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (RepeatCreateObject e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
