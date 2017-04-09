package com.lutotargaryen.poi.readexcel;

import java.io.Serializable;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtil implements Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	static Date StringToDate(String str) throws ParseException{
		Date date = null;
		String pattern = null;
		
		if(str.contains("-")){
			pattern = "yyyy-MM-dd";
		}else if(str.contains("年")){
			pattern = "yyyy年MM月dd";
		}else if(str.contains("/")){
			pattern = "yyyy/MM/dd";
		}
		SimpleDateFormat sdf = new SimpleDateFormat(pattern);
		date = sdf.parse(str);
		return date;
	}
	
}
