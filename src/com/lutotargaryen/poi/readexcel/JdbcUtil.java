package com.lutotargaryen.poi.readexcel;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 数据库连接类
 * @author luto
 *
 */
public class JdbcUtil {
	/*//资源文件解析器
	private final static PropertiesParser PARSER = PropertiesParser.newInstance();
	//数据库连接驱动
	final static String DRIVERCLASS = PARSER.getProperty("jdbc.driver_class");
	//数据库连接url
	final static String URL = PARSER.getProperty("jdbc.url");
	//数据库连接用户名
	final static String USERNAME = PARSER.getProperty("jdbc.username");
	//数据库连接密码
	final static String PASSWROD = PARSER.getProperty("jdbc.password");*/
	
	//数据库连接驱动
	private static String DRIVERCLASS;
	//数据库连接url
	private static String URL;
	//数据库连接用户名
	private static String USERNAME;
	//数据库连接密码
	private  static String PASSWROD;
	
	static void setDRIVERCLASS(String dRIVERCLASS) {
		DRIVERCLASS = dRIVERCLASS;
	}
	static void setURL(String uRL) {
		URL = uRL;
	}
	static void setUSERNAME(String uSERNAME) {
		USERNAME = uSERNAME;
	}
	static void setPASSWROD(String pASSWROD) {
		PASSWROD = pASSWROD;
	}
	/**
	 * 获取数据库连接的方法
	 */
	private static Connection getConnection(){
		java.sql.Connection connection = null;
		try {
			//加载数据库连接驱动
			Class.forName(DRIVERCLASS);
			connection = DriverManager.getConnection(URL, USERNAME, PASSWROD);
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return connection;
	}
	/**
	 * 释放资源方法
	 * @param parameters 需要释放的参数列表
	 */
	private static void closeObject(Object...parameters){
		//判断是否有参数
		if(parameters != null && parameters.length > 0){
			try {
				for (Object object : parameters) {
					if(object instanceof ResultSet){
						//如果是ResultSet的实例对象
						((ResultSet) object).close();
					}
					if(object instanceof Statement){
						//如果是Statement的实例对象
						((Statement) object).close();
					}
					if(object instanceof Connection){
						//如果是Connection的实例对象
						Connection connection = (Connection) object;
						if(connection != null && connection.isClosed()){
							connection.close();
							//释放内存中资源对象
							connection = null;
						}
					}
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	/**
	 * 设置参数方法
	 * @param statement
	 * @param parameters
	 */
	private static void setParameter(PreparedStatement statement,Object...parameters){
		if(statement != null && parameters != null && parameters.length > 0){
			try {
				for(int i = 0;i < parameters.length;i++){
					statement.setObject(i + 1, parameters[0]);
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	/**
	 * 执行DQL(数据库查询语句(select))方法
	 * @param sql	sql语句
	 * @param parameters	sql语句中需要的参数
	 * @return	从数据库中查询出的结果集合
	 */
	static List<Map<String,Object>> executeQuery(String sql,Object...parameters){
		//查询的结果列表
		List<Map<String,Object>> talbe = null;
		//连接对象
		Connection connection = null;
		//预处理对象
		PreparedStatement statement = null;
		//从数据库中查询出的结果集
		ResultSet resultSet = null;
		try {
			//获取连接
			connection = getConnection();
			//创建预编译对象
			statement = connection.prepareStatement(sql);
			//设置参数
			setParameter(statement, parameters);
			//执行executeQuery方法,返回ResultSet结果集
			resultSet = statement.executeQuery();
			if(resultSet != null){
				//把reusltSet结果集转换为一张虚拟的表
				ResultSetMetaData rsd = resultSet.getMetaData();
				//获取该虚拟表中的列数
				int columnCount = rsd.getColumnCount();
				//实例化talbe
				talbe = new ArrayList<>();
				//将结果集中的类容存储到list列表中
				//遍历result中的每一行
				while(resultSet.next()){
					//定义存储每一行数据的Map集合
					Map<String,Object> row = new HashMap<>();
					//遍历一行每一列
					for(int i = 0;i<columnCount;i++){
						//获取列名
						String columnName = rsd.getColumnName(i+1);
						//获取值
						String columnValue = resultSet.getString(columnName);
						//把列名和值存储到map集合
						row.put(columnName, columnValue);
					}
					//把每一行的数据添加到table中
					talbe.add(row);
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			//在finally中释放资源
			closeObject(resultSet,statement,connection);
		}
		return talbe;
	}
	/**
	 * 执行DML(数据库操纵语句(update,delect,insert))
	 * @param sql	sql语句
	 * @param parameters	sql语句中需要的参数
	 * @return	SQL 数据操作语言 (DML) 语句的行数
	 */
	static int executeUpdate(String sql,Object...parameters){
		//SQL 数据操作语言 (DML) 语句的行数
		int row = -1;
		//创建连接对象
		Connection connection = null;
		//预编译对象
		PreparedStatement statement = null;
		
		try {
			//获取连接对象
			connection = getConnection();
			//创建预编译对象
			statement = connection.prepareStatement(sql);
			//设置参数
			setParameter(statement, parameters);
			//执行预编译对象
			row = statement.executeUpdate();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			//在finally中释放资源
			closeObject(statement,connection);
		}
		return row;
	}
	/**
	 * 执行DDL(数据库定义语句(create,alter,drop),此处用于创建表) 还没写完
	 * @param table		表名
	 * @param parameters	表中所需字段
	 * @return
	 */
	@SuppressWarnings("unused")
	private static String executeDefinition(String table,Object...parameters){
		String tableName = null;
		//创建连接对象
		Connection connection = null;
		//预编译对象
		PreparedStatement statement = null;
		try {
			//获取连接对象
			connection = getConnection();
			//创建预编译对象
			statement = connection.prepareStatement("");
			
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			//在finally中释放资源
			closeObject(statement,connection);
		}
		
		return tableName;
	}
	

}
