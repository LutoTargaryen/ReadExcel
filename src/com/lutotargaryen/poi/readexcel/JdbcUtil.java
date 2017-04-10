package com.lutotargaryen.poi.readexcel;

import java.io.Serializable;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
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
public class JdbcUtil implements Serializable{
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	private static Connection conn;
	
	static void setConnnection(Connection connection){
		conn = connection;
	}
	
	/**
	 * 获取数据库连接的方法
	 */
	private static Connection getConnection(){
		return conn;
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
					statement.setObject(i + 1, parameters[i]);
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	/**
	 * 执行DML(数据库操纵语句(update,delect,insert))
	 * @param sql	sql语句
	 * @param parameters	sql语句中需要的参数
	 * @return	SQL 数据操作语言 (DML) 语句的行数
	 */
	static int executeUpdate(String sql,List<Object[]> parameters){
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
			for (Object[] objects : parameters) {
				setParameter(statement, objects);
				//添加到批处理集合
				statement.addBatch();
			}
			//执行批处理集合
			row = statement.executeBatch().length;
			
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
	/**
	 * 获取表中的列
	 * @param tableName
	 * @return
	 */
	static List<Map<String,Object>> getColumns(String tableName){
		//查询的结果列表
		List<Map<String,Object>> talbe = null;
		//连接对象
		Connection connection = null;
		ResultSet resultSet = null;
		try {
			//获取连接
			connection = getConnection();
			DatabaseMetaData dbmd = connection.getMetaData();
			resultSet = dbmd.getColumns(connection.getCatalog(),dbmd.getUserName(),tableName,null);
		
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
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			//在finally中释放资源
			closeObject(resultSet,connection);
		}
		return talbe;
	}
	/**
	 * 获取数据库中的所有表名
	 * @return
	 */
	static List<String> getTable(){
		//查询的结果列表
		List<String> list = null;
		//连接对象
		Connection connection = null;
		ResultSet resultSet = null;
		try {
			//获取连接
			connection = getConnection();
			DatabaseMetaData dbmd = connection.getMetaData();
			resultSet = dbmd.getTables(connection.getCatalog(),dbmd.getUserName(),null,null);
		
			if(resultSet != null){
				//实例化talbe
				list = new ArrayList<>();
				//将结果集中的类容存储到list列表中
				//遍历result中的每一行
				while(resultSet.next()){
					//获取表名
					String columnValue = resultSet.getString("TABLE_NAME");
					list.add(columnValue);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			//在finally中释放资源
			closeObject(resultSet,connection);
		}
		return list;
	}

}
