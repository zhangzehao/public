package com.sfpay.excel.bean;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import org.antlr.stringtemplate.StringTemplate;
import org.antlr.stringtemplate.StringTemplateGroup;



class Student{
	private String name;
	
	private int age;
	
	private List<String> hobbys = new ArrayList<String>();

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public List<String> getHobbys() {
		return hobbys;
	}

	public void setHobbys(List<String> hobbys) {
		this.hobbys = hobbys;
	}
}

class Row{
	private String name1;
	
	private String name2;
	
	private String name3;

	public String getName1() {
		return name1;
	}

	public void setName1(String name1) {
		this.name1 = name1;
	}

	public String getName2() {
		return name2;
	}

	public void setName2(String name2) {
		this.name2 = name2;
	}

	public String getName3() {
		return name3;
	}

	public void setName3(String name3) {
		this.name3 = name3;
	}
	
	
}

class Worksheet{
	private String sheet;
	
	private int columnNum;
	
	private int rowNum;
	
	private List<Row> rows;

	public String getSheet() {
		return sheet;
	}

	public void setSheet(String sheet) {
		this.sheet = sheet;
	}

	public List<Row> getRows() {
		return rows;
	}

	public void setRows(List<Row> rows) {
		this.rows = rows;
	}

	public int getColumnNum() {
		return columnNum;
	}

	public void setColumnNum(int columnNum) {
		this.columnNum = columnNum;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}
	
}

/**
 * 
 * @author 顺丰支付  张泽豪
 * 
 * 2013-8-5 上午11:25:56 
 * 
 */
public class TestTemplate{
	
	public static void main(String[] args) throws FileNotFoundException{
		TestTemplate template = new TestTemplate();
		template.output2();
	}

	/**
	 * @param args
	 * @throws FileNotFoundException 
	 */
	public void template() throws FileNotFoundException {
		StringTemplate st = new StringTemplate("hello,$name$");
		st.setAttribute("name", "china");
		System.out.println(st.toString());
		
		StringTemplate st2 = new StringTemplate("select $columns:{<i>$it$</i>\n}$ from users"); 
		List<String> columns = new ArrayList<String>();
		columns.add("a");
		columns.add("b");
		columns.add("c");
		columns.add("d");
		columns.add("e");
		st2.setAttribute("columns",columns);
		System.out.println(st2.toString());
		
		StringTemplate st3 = new StringTemplate("$students:{" +
				"$it.name$," +
				"$it.age$," +
				"$it.hobbys:{$it$,}$" +
			"}$"); 
		List<Student> students = new ArrayList<Student>();
		Student student = new Student();
		student.setName("hunter");
		student.setAge(24);
		List<String> hobbyList = new ArrayList<String>();
		hobbyList.add("sports");
		hobbyList.add("grils");
		hobbyList.add("money");
		student.setHobbys(hobbyList);
		students.add(student);
		
		student = new Student();
		student.setName("zhangzehao");
		student.setAge(25);
		hobbyList = new ArrayList<String>();
		hobbyList.add("movie");
		hobbyList.add("coding");
		student.setHobbys(hobbyList);
		students.add(student);
		
		st3.setAttribute("students", students);
		
		System.out.println(st3.toString());
		
	}
	
	/**
	 * 生成数据量大的时候，该方法会出现内存溢出
	 * @throws FileNotFoundException
	 */
	public void output1() throws FileNotFoundException{
		StringTemplateGroup stGroup = new StringTemplateGroup("stringTemplate");         
		StringTemplate st4 =  stGroup.getInstanceOf("com/sfpay/excel/bean/test");
		List<Worksheet> worksheets = new ArrayList<Worksheet>();

		File file = new File("D:/output.xls");
		PrintWriter writer = new PrintWriter(new BufferedOutputStream(new FileOutputStream(file)));
		
		for(int i=0;i<30;i++){
			Worksheet worksheet = new Worksheet();
			worksheet.setSheet("第"+(i+1)+"页");
			List<Row> rows = new ArrayList<Row>();
			for(int j=0;j<6000;j++){
				Row row = new Row();
				row.setName1("zhangzehao");
				row.setName2(""+j);
				row.setName3(i+" "+j);
				rows.add(row);
			}
			worksheet.setRows(rows);
			worksheets.add(worksheet);
		}
		
		st4.setAttribute("worksheets", worksheets);
		writer.write(st4.toString());
		writer.flush();
		writer.close();
		System.out.println("生成excel完成");
	}
	
	/**
	 * 该方法不管生成多大的数据量，都不会出现内存溢出，只是时间的长短
	 * 
	 * 经测试，生成1800万数据，6~10分钟之间，3G大的文件，打开大文件就看内存是否足够大了
	 * 
	 * 数据量小的时候，推荐用jxls的模板技术生成excel文件，谁用谁知道，大数据量可以结合该方法使用
	 * 
	 * @throws FileNotFoundException
	 */
	public void output2() throws FileNotFoundException{
		long startTimne = System.currentTimeMillis();
		
		StringTemplateGroup stGroup = new StringTemplateGroup("stringTemplate");    
		
		//写入excel文件头部信息
		StringTemplate head =  stGroup.getInstanceOf("com/sfpay/excel/bean/head");
		File file = new File("D:/output.xls");
		PrintWriter writer = new PrintWriter(new BufferedOutputStream(new FileOutputStream(file)));
		writer.print(head.toString());
		writer.flush();
		
		int sheets = 300;
		//excel单表最大行数是65535
		int maxRowNum = 60000;
		
		//写入excel文件数据信息
		for(int i=0;i<sheets;i++){
			StringTemplate body =  stGroup.getInstanceOf("com/sfpay/excel/bean/body");
			Worksheet worksheet = new Worksheet();
			worksheet.setSheet(" "+(i+1)+" ");
			worksheet.setColumnNum(3);
			worksheet.setRowNum(maxRowNum);
			List<Row> rows = new ArrayList<Row>();
			for(int j=0;j<maxRowNum;j++){
				Row row = new Row();
				row.setName1(""+new Random().nextInt(100000));
				row.setName2(""+j);
				row.setName3(i+""+j);
				rows.add(row);
			}
			worksheet.setRows(rows);
			body.setAttribute("worksheet", worksheet);
			writer.print(body.toString());
			writer.flush();
			rows.clear();
			rows = null;
			worksheet = null;
			body = null;
			Runtime.getRuntime().gc();
			System.out.println("正在生成excel文件的 sheet"+(i+1));
		}
		
		//写入excel文件尾部
		writer.print("</Workbook>");
		writer.flush();
		writer.close();
		System.out.println("生成excel文件完成");
		long endTime = System.currentTimeMillis();
		System.out.println("用时="+((endTime-startTimne)/1000)+"秒");
	}

}
