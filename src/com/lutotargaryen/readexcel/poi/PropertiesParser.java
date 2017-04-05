package com.lutotargaryen.readexcel.poi;

import java.io.IOException;
import java.util.Properties;

/**
 * 资源文件解析类
 * @author luto
 *
 */
public class PropertiesParser extends Properties {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private static PropertiesParser parser;
	
	/**
	 * 为使用 单例模式，所以要私有化构造方法，
	 */
	private PropertiesParser(){
	}
	/**
	 * 装载属性文件
	 */
	{
		try {
			this.load(this.getClass().getClassLoader().getResourceAsStream("jdbc.properties"));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 创建当前实例的静态方法
	 */
	public static PropertiesParser newInstance(){
		if(parser == null){
			parser = new PropertiesParser();
		}
		return parser;
	}
	@Override
	public String getProperty(String key) {
		return super.getProperty(key);
	}
}
