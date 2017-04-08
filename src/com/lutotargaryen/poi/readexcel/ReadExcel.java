package com.lutotargaryen.poi.readexcel;

import java.util.List;
import java.util.Map;

import com.lutotargaryen.poi.exception.RepeatCreateObject;

public class ReadExcel {
	
	// 使用volatile关键字保证多线程访问时readExcel的可见性
	private volatile static ReadExcel readExcel;
	/**
	 * 创建ReadExcel对象
	 * 如果不设置UserName和Password则默认数据库连接用户为root,数据库连接密码也为root
	 * @param Driver 数据库驱动类路径
	 * @param DBUrl 数据库连接url
	 * @return
	 * @throws RepeatCreateObject 
	 */
	public static ReadExcel newInstance(String Driver,String DBUrl) throws RepeatCreateObject{
		if(readExcel == null){
			synchronized (ReadExcel.class) {
				if(readExcel == null){
					readExcel = new ReadExcel();
					JdbcUtil.setDRIVERCLASS(Driver);
					JdbcUtil.setURL(DBUrl);
					JdbcUtil.setUSERNAME("root");
					JdbcUtil.setPASSWROD("root");
				}
			}
			
		}else{
			throw new RepeatCreateObject("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	/**
	 * 创建ReadExcel对象
	 * @param Driver 数据库驱动类路径
	 * @param DBUrl	数据库连接url
	 * @param UserName	数据库连接用户名
	 * @param PassWord	数据库连接密码
	 * @return
	 * @throws RepeatCreateObject 
	 */
	public static ReadExcel newInstance(String Driver,String DBUrl,String UserName,String PassWord) throws RepeatCreateObject{
		if(readExcel == null){
			synchronized (ReadExcel.class) {
				if(readExcel == null){
					readExcel = new ReadExcel();
					JdbcUtil.setDRIVERCLASS(Driver);
					JdbcUtil.setURL(DBUrl);
					JdbcUtil.setUSERNAME(UserName);
					JdbcUtil.setPASSWROD(PassWord);
				}
			}
		}else{
			throw new RepeatCreateObject("This class can have only one instance object, and you have created an instance object");
		}
		return readExcel;
	}
	public List<Map<String,Object>> queryTest(){
		List<Map<String,Object>> list = JdbcUtil.executeQuery("select * from user");
		return list;
	}
}
