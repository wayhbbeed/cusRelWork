package main.java.dao;

import java.util.ArrayList;
import java.util.Iterator;

import main.java.Cus;
import main.java.utils.Words;

public class Dao {
	private Words words=null;
	private ArrayList<Cus> al=null;
	private Cus cus=null;

	//创建数据库
	public void init() {
		DBAccess db = new DBAccess();
		if(db.createConn()) {
		   String sql = " create table CUSTOM ( " +
		   		        " id varchar ( 30 ) primary key, " +
						" Province varchar ( 50 ), "+
						" Province_manager varchar, ( 50 )" +
						" Custom varchar ( 50 ), "+
						" Business varchar ( 50 ), "+
						" Receiver varchar ( 50 ), "+
						" Bill_info varchar ( 200), "+
						" Invoice_info varchar ( 200 ), "+
						" Cus_info varchar ( 50 ), "+
						" Goods varchar ( 50 ), "+
						" Requirement varchar ( 100 ))";
			db.update(sql);
			db.closeStm();
			db.closeConn();
		}
	}
	
	//批量将WORD表格插入数据库
	public void insertWordToDb() throws Exception{
		DBAccess db = new DBAccess();
		words=new Words();
		al=new ArrayList<Cus>();
		al=words.getFileContent();
		Iterator<Cus> it=al.iterator();
		while (it.hasNext()) {
			cus=new Cus();
			cus=it.next();
			if(db.createConn()) {
				    String sql = "insert into custom values( ' "
				    	          + cus.getId() + "','"
				    	          + cus.getProvince() + "','"
				    	          + cus.getProvince_manager() + "','"
				    	          + cus.getCustom() + "','"
				    	          + cus.getBusiness() + "','"
				    	          + cus.getReceiver() + "','"
				    	          + cus.getBill_info() + "','"
				    	          + cus.getInvoice_info() + "','"
				    	          + cus.getCus_info() + "','"
				    	          + cus.getGoods() + "','"
				    	          + cus.getRequirement() + "')";
					db.update(sql);
					db.closeStm();
					db.closeConn();
				}
		}
		
	}
}
