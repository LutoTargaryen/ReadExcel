package com.lutotargaryen.poi.readexcel;

import java.util.List;
import java.util.Map;

public class ReadExcel {
	
	private static ReadExcel readExcel = null;
	/**
	 * 创建ReadExcel对象
	 * 如果不设置UserName和Password则默认数据库连接用户为root,数据库连接密码也为root
	 * @param DBUrl 数据库连接url
	 * @return
	 * @throws Exception 
	 */
	public static ReadExcel newInstance(String DBUrl) throws Exception{
		if(readExcel == null){
			readExcel = new ReadExcel();
			JdbcUtil.setDRIVERCLASS("com.mysql.jdbc.Driver");
			JdbcUtil.setURL(DBUrl);
			JdbcUtil.setUSERNAME("root");
			JdbcUtil.setPASSWROD("root");
		}else{
			throw new Exception("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	/**
	 * 创建ReadExcel对象
	 * @param DBUrl	数据库连接url
	 * @param UserName	数据库连接用户名
	 * @param PassWord	数据库连接密码
	 * @return
	 * @throws Exception 
	 */
	public static ReadExcel newInstance(String DBUrl,String UserName,String PassWord) throws Exception{
		if(readExcel == null){
			readExcel = new ReadExcel();
			JdbcUtil.setDRIVERCLASS("com.mysql.jdbc.Driver");
			JdbcUtil.setURL(DBUrl);
			JdbcUtil.setUSERNAME(UserName);
			JdbcUtil.setPASSWROD(PassWord);
		}else{
			throw new Exception("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	public List<Map<String,Object>> queryTest(){
		List<Map<String,Object>> list = JdbcUtil.executeQuery("select * from user");
		return list;
	}
}
