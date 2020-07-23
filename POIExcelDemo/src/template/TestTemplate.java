package template;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

//参考地址：http://mylfd.iteye.com/blog/1982101
public class TestTemplate {
	public static void main(String[] args) {
		
		ExcelTemplate instance = ExcelTemplate.getInstance();
		
		ExcelTemplate excel = instance.readTemplatePath(
				"D:\\workspaceGit\\testdd2\\WebContent\\docs\\收费通知单模板.xlsx");
		
		
		String money = "至通知单生成之日，本期累计应缴为：20元，您仍需缴费：30元，预存余额为：40元。";												
		
		Sheet sheet = instance.getSheet();
		Cell row40 = sheet.getRow(4).getCell(0);
		row40.setCellValue(money);
		
		String startDate = "计费起止期间：2020-07-01";
		Cell row60 = sheet.getRow(6).getCell(0);
		row60.setCellValue(startDate);
		
		String endDate = "通知单生成日期：2020-07-31";
		Cell row67 = sheet.getRow(6).getCell(7);
		row67.setCellValue(endDate);
		
		instance.setStartIndex(9, 0);
	
		excel.creatNewRowMerginCol();
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRowMerginCol();
		excel.createNewCol("bbb");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRowMerginCol();
		excel.createNewCol("ccc");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRowMerginCol();
		excel.createNewCol("ddd");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRowMerginCol();
		excel.createNewCol("eee");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		
		excel.creatNewRowMerginCol();
		excel.createNewCol("ffff");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		int curRowIndex = instance.getCurRowIndex();
		
		int nowRowIndex = curRowIndex + 3;
		instance.setStartIndex(nowRowIndex, 0);
		
		excel.creatNewRow();
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("bbb");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("ccc");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("ddd");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("eee");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		
		
//		Map<String, String> datas = new HashMap<String, String>();
//		datas.put("title", "拉斯维加斯");
//		datas.put("date", new Date().toString());
//		datas.put("department", "百合科技人事部");
//		excel.replaceFind(datas);
//		excel.insertSer();
		excel.writeToFile("D:\\workspaceGit\\testdd2\\WebContent\\docs\\poi24.xlsx");
	}
}
