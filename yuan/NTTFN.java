package beanUtils.beanUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;
import java.io.BufferedReader;
import java.io.InputStreamReader;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import beanUtils.beanUtils.ExcelUtils1.CallBackExcelSingleRow1;

public class NTTFN {

	static String mdb = "OM";
	static String number = "0014";
	static String bid = "105" + number;
	static String isData = "da";
	static boolean reqntt = false;
	static boolean resntt = true;
	static String fileName = "活动信息查询";
	static int reqBegin = 11;
	static int reqEnd = 17;
	static int resBegin = 24;
	static int resEnd = 29;
	static boolean page = false;
	static boolean resList = false;
	static String className = "ActivityInfo";
	static boolean select = true;
	static boolean select_list=true;
	static boolean insert = false;
	static boolean update = false;
	static boolean delete = false;

	static String obName = className.substring(0, 1).toLowerCase() + className.substring(1);
	static String component = "FN" + number;
	static String nttFilePath = "D:/JavaBean";
	static String fnFilePath = "D:/JavaBean";
	static String packageName = mdb.toLowerCase();
	static String fid = mdb + component;
	static String az = "";
	static String bz = "";
	static String cz = "";
	static String dz = "";
	static String ez = "";

	public static void main(String[] args) throws Exception {
		File a = new File(nttFilePath);
		File b = new File(fnFilePath);
		if (!a.exists()) {
			a.mkdirs();
		}
		if (!b.exists()) {
			b.mkdirs();
		}
		String path = "D:\\platform-document\\02_系统设计\\20_业务设计\\105_在线商城\\02_接口定义\\" + number + "_" + fileName
				+ "\\接口定义书" + "（" + fileName + "）.xlsx";
		showBean(path, reqBegin, reqEnd, resBegin, resEnd, 1);
		System.out.println("------------------------------------------------");
	}

	public static void showBean(String path, final int reqBegin, final int reqEnd, final int resBegin, final int resEnd,
			final int beginRow) throws Exception {

		if (reqntt) {
			InputStream input = new FileInputStream(new File(path));
			ExcelUtils1.readExcel(input, ExcelUtils1.EXCEL_FORMAT_XLSX, 1, new CallBackExcelSingleRow1() {

				ArrayList<NttCreateCondition> conditions = new ArrayList<NttCreateCondition>();

				public void readRow(List<String> rowContent, int rowIndex) {
					if (rowIndex >= reqBegin && rowIndex <= reqEnd) {
						String name = rowContent.get(1).trim();// 全称
						String shortName = rowContent.get(2).trim();// 缩写
						String desc = rowContent.get(3).trim();// 描述
						conditions.add(getCondition(name, shortName, desc));
					}
					if (rowIndex == (reqEnd + 1)) {
						conditions.trimToSize();

						String str = "";
						for (NttCreateCondition condition : conditions) {
							str += ShortDesc(condition);
						}
						for (NttCreateCondition condition : conditions) {
							str += Area(condition);
						}
						REQ();
						reqData(str);

						String check = "";
						String ss = "";
						int i = 1;
						String error = "";
						String edit = "";

						if (delete || insert || update) {
							edit += "Query query = new Query(Criteria.where(" + className + ").is(\"\"));\r\n";
						}
						if (insert || update) {
							edit += className + " " + obName + "=null;\r\n";
						}
						String insert1 = "";
						if (insert) {
							insert1 = "	if (" + mdb + "Const.INSERT.equals(operateType)) {// 插入操作\r\n" + obName
									+ "= mdbTemplateOM.findOne(query, " + className + ".class);\r\n" + "if (" + obName
									+ " != null) {\r\n"
									+ "setFailureReturnCode(res, ReturnCodeCategory.FAILURE_FUNCTION_DATE_OPERATION, BUSINESS_FUNCTION_ID,"
									+ className + "_IS_EXIST);// 数据已存在，无法插入\r\n" + "return res;\r\n" + "}\r\n" + obName
									+ "=new " + className + "();\r\n";

						}
						String update1 = "";
						if (update) {
							update1 = "if (" + mdb + "Const.UPDATE.equals(operateType)) {// 更新操作\r\n" + obName
									+ "= mdbTemplateOM.findOne(query, " + className + ".class);\r\n" + "if (" + obName
									+ " == null) {\r\n"
									+ "setFailureReturnCode(res, ReturnCodeCategory.FAILURE_FUNCTION_DATE_OPERATION, BUSINESS_FUNCTION_ID,"
									+ className + "_IS_NOT_EXIST);// 数据不存在，无法更新\r\n" + "return res;\r\n" + "}\r\n"
									+ "Update update = new Update();";
						}
						String update2 = "";
						for (NttCreateCondition condition : conditions) {
							ss += "private static final String " + Check2(condition) + "_IS_NULL=\"E0" + i + "\";\r\n";
							i++;
							String aa = "Condition.";
							if ("da".equals(isData)) {
								aa = "Data.";
							}

							error += "if (StringUtils.isNullOrSpace(" + condition.getName() + ")) {\r\n"
									+ "addErrorCode(res, BUSINESS_FUNCTION_ID," + Check2(condition) + "_IS_NULL, " + fid
									+ "RequestBody" + aa + Check2(condition) + ");\r\n" + "}\r\n";

							update2 += "if (StringUtils.isNotEmpty(" + condition.getName() + ")) {\r\n" + "update.set("
									+ obName + "." + Check2(condition) + ", condition.getName());\r\n" + "}\r\n";
						}

						String insert2 = "";

						for (NttCreateCondition condition : conditions) {
							check += Check(condition);
							String a = condition.getName();
							String b = a.substring(0, 1).toUpperCase() + a.substring(1);
							insert2 += obName + ".set" + b + "(" + a + ");\r\n";

						}
						String insert666 = "";
						if (insert) {
							insert666 = insert1 + insert2 + "mdbTemplate" + mdb + ".insert(" + obName + ");\r\n}";
							edit += insert666;
						}
						String update666 = "";
						if (update) {
							update666 = update1 + update2 + "mdbTemplate" + mdb + ".updateFirst(query, update, "
									+ className + ".class);\r\n}";
							edit += "else " + update666;
						}
						String delete666 = "";
						if (delete) {
							delete666 = "if (OMConst.DELETE.equals(operateType)) {// 删除\r\nmdbTemplate" + mdb
									+ ".remove(query, " + className + ".class);\r\n}\r\n";
							edit += "else " + delete666;
						}

						// FN(check, ss, error, edit);
						az = check;
						bz = ss;
						cz = error;
						dz = edit;

					}
				}
			});
		}
		if (resntt)

		{
			InputStream input = new FileInputStream(new File(path));
			ExcelUtils1.readExcel(input, ExcelUtils1.EXCEL_FORMAT_XLSX, 1, new CallBackExcelSingleRow1() {

				ArrayList<NttCreateCondition> conditions = new ArrayList<NttCreateCondition>();

				public void readRow(List<String> rowContent, int rowIndex) {
					if (rowIndex >= resBegin && rowIndex <= resEnd) {
						String name = rowContent.get(1).trim();// 全称
						String shortName = rowContent.get(2).trim();// 缩写
						String desc = rowContent.get(3).trim();// 描述
						conditions.add(getCondition(name, shortName, desc));
					}
					if (rowIndex == (resEnd + 1)) {
						conditions.trimToSize();

						String str = "";
						for (NttCreateCondition condition : conditions) {
							str += ShortDesc(condition);
						}
						for (NttCreateCondition condition : conditions) {
							str += Area(condition);
						}
					
						String rr="";
						for (NttCreateCondition condition : conditions) {
							String name=condition.getName().substring(0, 1).toUpperCase()+condition.getName().substring(1);
							rr+="data.set"+name+"("+obName+".get"+name+"());\r\n";
						}
					
						String asd="";
						asd+="Query query=new Query();\r\n";
						if(select_list){
							asd+="List<"+className+"> list="+"MDBTemplate" + mdb+".find(query,"+className+".class);\r\n";
							asd+="List<"+fid+"ResponseData> dataList=new ArrayList<>();\r\n";
							asd+="for("+fid+"ResponseData "+obName+":list){\r\n"
									+ fid+"ResponseData data=new "+fid+"ResponseData();\r\n"
									+rr+"dataList.add(data);"
									+ "}\r\n"+"res.setData(dataList);\r\n";
						}else{
							asd+=className+" "+obName+"="+"MDBTemplate" + mdb+".findOne(query,"+className+".class);";
						    asd+=fid+"ResponseData data=new "+fid+"ResponseData();";
						    asd+=rr+"res.setData(data);\r\n";
						}
						
						if(select){
							ez=asd;
						}
						
						res();
						resData(str);
					}
				}
			});
		}

		FN(az, bz, cz, dz,ez);
	}

	private static NttCreateCondition getCondition(String name, String shortName, String desc) {
		return new NttCreateCondition(name, shortName, desc);
	}

	// ---
	private static String Check2(NttCreateCondition condition) {
		String ss = condition.getName();
		List<Integer> list = new ArrayList<Integer>();
		for (int i = 0; i < ss.length(); i++) {
			char ch = ss.charAt(i);
			if (ch >= 'A' && ch <= 'Z') {
				list.add(i);
			}
		}
		for (Integer i : list) {
			ss = ss.substring(0, i) + "_" + ss.substring(i);
		}
		String sss = ss.toUpperCase();
		return sss;
	}

	// 参数检查
	private static String Check(NttCreateCondition condition) {
		String ss = condition.getName();
		String sss = ss.substring(0, 1).toUpperCase() + ss.substring(1);
		if ("da".equals(isData)) {
			return "String " + ss + "=req.getData().get" + sss + "();" + "//" + condition.getDesc() + "\r\n";
		} else {
			return "String " + ss + "=req.getCondition().get" + sss + "();" + "//" + condition.getDesc() + "\r\n";
		}
	}

	private static String ShortDesc(NttCreateCondition condition) {
		return "public static final String " + condition.getBigName() + " = \"" + condition.getShortName() + "\";\r\n";
	}

	private static String Area(NttCreateCondition condition) {
		return "/**\r\n*" + condition.getDesc() + "\r\n*/@JsonProperty(" + condition.getBigName()
				+ ")\r\nprivate String " + condition.getName() + ";\r\n";
	}

	public static void resData(String str) {
		String result = "package cn.sh.changxing.entity." + packageName
				+ ";\r\nimport org.codehaus.jackson.annotate.JsonProperty;\r\npublic class " + fid
				+ "ResponseBodyData {\r\n" + str;
		result += "}";

		try {
			File file = new File(nttFilePath, fid + "ResponseBodyData.java");
			@SuppressWarnings("resource")
			PrintStream ps = new PrintStream(new FileOutputStream(file));
			ps.println(result);// 往文件里写入字符串
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	public static void res() {
		String result = "";
		String result2 = "";
		String result3 = "";
		if (resList) {
			result3 = "import java.util.List;\r\n";
			result2 = "@JsonProperty(\"da\")\r\nprivate List<" + fid
					+ "ResponseBodyData> data;\r\npublic void setData(List<" + fid
					+ "ResponseBodyData> data) {\r\nthis.data = data;\r\n}\r\npublic List<" + fid
					+ "ResponseBodyData> getData() {\r\nreturn data;\r\n}\r\n";
		} else {
			result2 = "@JsonProperty(\"da\")\r\nprivate " + fid + "ResponseBodyData data;\r\npublic void setData(" + fid
					+ "ResponseBodyData data) {\r\nthis.data = data;\r\n}\r\npublic " + fid
					+ "ResponseBodyData getData() {\r\nreturn data;\r\n}\r\n";
		}
		if (page) {
			result = "package cn.sh.changxing.entity." + packageName + ";\r\n" + result3
					+ "import org.codehaus.jackson.annotate.JsonProperty;\r\nimport cn.sh.changxing.platform.EntityPageableResponseBody;\r\npublic class "
					+ fid
					+ "ResponseBody extends EntityPageableResponseBody {\r\nprivate static final long serialVersionUID = 1L;\r\n"
					+ result2 + "}";
		} else {
			result = "package cn.sh.changxing.entity." + packageName + ";\r\n" + result3
					+ "import org.codehaus.jackson.annotate.JsonProperty;\r\nimport cn.sh.changxing.platform.EntityResponseBody;\r\npublic class "
					+ fid
					+ "ResponseBody extends EntityResponseBody {\r\nprivate static final long serialVersionUID = 1L;\r\n"
					+ result2 + "}";
		}

		try {
			File file = new File(nttFilePath, fid + "ResponseBody.java");
			@SuppressWarnings("resource")
			PrintStream ps = new PrintStream(new FileOutputStream(file));
			ps.println(result);// 往文件里写入字符串
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	public static void reqData(String str) {
		String result = "package cn.sh.changxing.entity." + packageName
				+ ";\r\nimport org.codehaus.jackson.annotate.JsonProperty;\r\npublic class " + fid
				+ "RequestBodyData{\r\n" + str;
		result += "}";

		try {
			File file = new File(nttFilePath, fid + "RequestBodyData.java");
			@SuppressWarnings("resource")
			PrintStream ps = new PrintStream(new FileOutputStream(file));
			ps.println(result);// 往文件里写入字符串
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	public static void REQ() {
		String result1 = "";
		String result2 = "";
		if (page) {
			result2 = "import cn.sh.changxing.platform.EntityPageableRequestBody;\r\npublic class " + fid
					+ "RequestBody extends EntityPageableRequestBody ";
		} else {
			result2 = "import cn.sh.changxing.platform.EntityRequestBody;\r\npublic class " + fid
					+ "RequestBody extends EntityRequestBody ";
		}
		if ("da".equals(isData)) {
			result1 = "package cn.sh.changxing.entity." + packageName
					+ ";\r\nimport org.codehaus.jackson.annotate.JsonProperty;\r\n" + result2
					+ "{\r\nprivate static final long serialVersionUID = 1L;\r\n@JsonProperty(\"da\")\r\nprivate " + fid
					+ "RequestBodyData data=new " + fid + "RequestBodyData();\r\npublic void setData(" + fid
					+ "RequestBodyData data) {\r\nthis.data = data;\r\n}\r\npublic " + fid
					+ "RequestBodyData getData() {\r\nreturn data;\r\n}\r\n}";
		} else {
			result1 = "package cn.sh.changxing.entity." + packageName
					+ ";\r\nimport org.codehaus.jackson.annotate.JsonProperty;\r\n" + result2
					+ "{\r\nprivate static final long serialVersionUID = 1L;\r\n@JsonProperty(\"co\")\r\nprivate " + fid
					+ "RequestBodyData condition=new " + fid + "RequestBodyData();\r\npublic void setCondition(" + fid
					+ "RequestBodyData condition) {\r\nthis.condition = condition;\r\n}\r\npublic " + fid
					+ "RequestBodyData getCondition() {\r\nreturn condition;\r\n}\r\n}";
		}

		try {
			File file = new File(nttFilePath, fid + "RequestBody.java");
			@SuppressWarnings("resource")
			PrintStream ps = new PrintStream(new FileOutputStream(file));
			ps.println(result1);// 往文件里写入字符串
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	private static void FN(String check, String ss, String error, String edit,String ez) {
		String req = "Entity";
		String res = "Entity";
		String result1 = "";
		String result2 = "";
		if (!(insert || update || delete)) {
			edit = "";
		}
		if (page) {
			result1 += "import cn.sh.changxing.platform.RequestBodyPage;import org.springframework.data.mongodb.core.query.Query;";
			result2 = "RequestBodyPage requestBodyPage = req.getPage();\r\nint pageSkip = requestBodyPage.getSkipCount();// 从第几条记录开始\r\nint pageSize = requestBodyPage.getPageSize();// 默认每页记录条数\r\nint changePageFlag = requestBodyPage.getChangePageFlag();// 0：否（查询）1：是（仅换页)\r\n// 设置总页数\r\nif (changePageFlag == 0) {\r\nres.setRecordTotal((int) (mdbTemplate"
					+ mdb + ".count(new Query(), \"\")));\r\n}";
		}
		if (reqntt) {
			req = fid;
			result1 += "import cn.sh.changxing.entity." + packageName + "." + req + "RequestBody;\r\n";
		} else {
			result1 += "import cn.sh.changxing.platform.EntityRequestBody;\r\n";
		}
		if (resntt) {
			res = fid;
			result1 += "import cn.sh.changxing.entity." + packageName + "." + res + "ResponseBody;\r\n";
		} else {
			result1 += "import cn.sh.changxing.platform.EntityResponseBody;\r\n";
		}
		String result = "package cn.sh.changxing.bs." + packageName + ";\r\n"
				+ "import org.slf4j.Logger;import org.slf4j.LoggerFactory;\r\n"
				+ "import cn.sh.changxing.common.utils.StringUtils;\r\n"
				+ "import cn.sh.changxing.platform.ReturnCodeCategory;\r\n"
				+ "import org.springframework.beans.factory.annotation.Autowired;\r\n"
				+ "import org.springframework.stereotype.Component;\r\n" + "import cn.sh.changxing.mdb." + packageName
				+ ".MDBTemplate" + mdb + ";\r\n" + result1
				+ "import cn.sh.changxing.platform.service.provide.EntityFunctionProvide;\r\n"
				+ "import cn.sh.changxing.platform.service.provide.ProvideException;\r\n" + "/**\r\n" + "*\r\n" + "* "
				+ fileName + "\r\n" + "* @author niushunyuan\r\n" + "*/\r\n" + "@Component(\"" + component + "\")\r\n"
				+ "public class " + fid + " extends EntityFunctionProvide<" + req + "RequestBody, " + res
				+ "ResponseBody>{\r\n" + "private static final Logger LOG=LoggerFactory.getLogger(" + fid
				+ ".class);\r\n" + "@Autowired private MDBTemplate" + mdb + " mdbTemplate" + mdb + ";\r\n"
				+ "private static final String BUSINESS_FUNCTION_ID =\"" + bid + "\";\r\n" + ss + "@Override\r\n"
				+ " protected " + res + "ResponseBody onRequestBody(" + req
				+ "RequestBody req) throws ProvideException {\r\n" + "LOG.info(\"" + fileName + "开始\");\r\n" + "" + res
				+ "ResponseBody res=new " + res + "ResponseBody();\r\n"
				+ "res.setSerialNumber(req.getCommon().getSerialNumber());\r\n" + check + result2 + "\r\n" + error
				+ "if (res.getCommon().getError().size() > 0) {\r\nreturn res;\r\n}\r\n" + edit
				+ "//setFailureReturnCode(res, ReturnCodeCategory.FAILURE_FUNCTION_DATE_OPERATION, BUSINESS_FUNCTION_ID,\");\r\n"+ez+"return res;}}";

		try {
			File file = new File(fnFilePath, fid + ".java");
			@SuppressWarnings("resource")
			PrintStream ps = new PrintStream(new FileOutputStream(file));
			ps.println(result);// 往文件里写入字符串
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}

class NttCreateCondition {
	private String name = "";
	private String shortName = "";
	private String desc = "";

	public NttCreateCondition(String name, String shortName, String desc) {
		this.name = name;
		this.shortName = shortName;
		this.desc = desc;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getShortName() {
		return shortName;
	}

	public void setShortName(String shortName) {
		this.shortName = shortName;
	}

	public String getDesc() {
		return desc;
	}

	public void setDesc(String desc) {
		this.desc = desc;
	}

	public String getBigName() {
		String bigName = name;
		char[] chars = bigName.toCharArray();
		StringBuilder sBuild = new StringBuilder();
		for (char index : chars) {
			if (Character.isUpperCase(index)) {
				sBuild.append("_");
			}
			sBuild.append(index);
		}
		return sBuild.toString().toUpperCase();
	}
}

class ExcelUtils1 {

	public static final int READ_ALL_SHEET = -1;

	public static final int EXCEL_FORMAT_XLS = 0;
	public static final int EXCEL_FORMAT_XLSX = 1;
	public static final int EXCEL_FORMAT_CSV = 2;

	public static final int ERROR_EXCEL_FORMAT = -1;

	private static final String EMPTY = "";
	private static final String DEFAULT_CHARSET = "utf-8";
	private static final String POINT_CODE = ",";

	/**
	 * 该方法是用来解析Excel表格的,每读取一行便调用CallBackExcelSingleRow回调函数，
	 * 返回当前行所有列的数据和当前行在excel中属于第几行<br/>
	 * 对于csv格式的文件,返回的数据不会经过任何改变<br/>
	 * 1.对于excel中的公式,该接口当前返回的该公式的文档,不会进行相应的计算<br/>
	 * 2.对于boolean类型的数据,返回"TRUE"或者"FALSE"<br/>
	 * 3.对于时间类型的数据,接口当前返回的是yyyyMMdd格式的时间数据<br/>
	 * 
	 * @param excelInputStream
	 *            所要解析文件的输入流
	 * @param excelFormat
	 *            所要解析文件的格式,目前只能解析xls,xlsx和csv格式的文件
	 * @param readWhatSheet
	 *            如果解析的是xls和xlsx格式的文件,该参数表示读取文件中的哪一个Sheet表格,默认读取所有Sheet的数据
	 * @param singleRow
	 *            readRow(currentRowContent, rowIndex)
	 *            currentRowContent返回当前行的所有内容,rowIndex是当前所在Sheet的行数
	 */
	public static void readExcel(InputStream excelInputStream, int excelFormat, int readWhatSheet,
			CallBackExcelSingleRow1 singleRow) throws Exception {
		int needReadSheet = READ_ALL_SHEET;
		if (readWhatSheet >= 0) {
			needReadSheet = readWhatSheet;
		}

		if (excelFormat == EXCEL_FORMAT_XLS) {
			readXls(excelInputStream, singleRow, needReadSheet);
			return;
		}
		if (excelFormat == EXCEL_FORMAT_XLSX) {
			readXlsx(excelInputStream, singleRow, needReadSheet);
			return;
		}
		if (excelFormat == EXCEL_FORMAT_CSV) {
			readCsv(excelInputStream, singleRow);
			return;
		}

		singleRow.readRow(null, ERROR_EXCEL_FORMAT);
	}

	private static void readXls(InputStream excelInputStream, CallBackExcelSingleRow1 singleRow, int readWhatSheet)
			throws Exception {
		HSSFWorkbook excelDoc = new HSSFWorkbook(excelInputStream);
		int beginSheet = 0;
		int endSheet = excelDoc.getNumberOfSheets();
		if (readWhatSheet != READ_ALL_SHEET) {
			beginSheet = readWhatSheet - 1;
			endSheet = readWhatSheet;
		}

		for (int currentSheet = beginSheet; currentSheet < endSheet; currentSheet++) {
			HSSFSheet currentSheetDoc = excelDoc.getSheetAt(currentSheet);
			if (currentSheetDoc == null) {
				continue;
			}
			for (int rowIndex = 0; rowIndex <= currentSheetDoc.getLastRowNum(); rowIndex++) {
				HSSFRow currentRowDoc = currentSheetDoc.getRow(rowIndex);

				int currentRowFirstCol = currentRowDoc.getFirstCellNum();
				int currentRowLastCol = currentRowDoc.getLastCellNum();

				List<String> currentRowContent = new ArrayList<String>();

				for (int currentCol = currentRowFirstCol; currentCol < currentRowLastCol; currentCol++) {
					HSSFCell currentCellDoc = currentRowDoc.getCell(currentCol);
					if (currentCellDoc == null) {
						currentRowContent.add(EMPTY);
						continue;
					}
					currentRowContent.add(getHSSFWordbookCellContent(currentCellDoc));
				}
				singleRow.readRow(currentRowContent, rowIndex);
			}
		}
	}

	private static void readXlsx(InputStream excelInputStream, CallBackExcelSingleRow1 singleRow, int readWhatSheet)
			throws Exception {
		XSSFWorkbook excelDoc = new XSSFWorkbook(excelInputStream);
		int beginSheet = 0;
		int endSheet = excelDoc.getNumberOfSheets();
		if (readWhatSheet != READ_ALL_SHEET) {
			beginSheet = readWhatSheet - 1;
			endSheet = readWhatSheet;
		}

		for (int currentSheet = beginSheet; currentSheet < endSheet; currentSheet++) {
			XSSFSheet currentSheetDoc = excelDoc.getSheetAt(currentSheet);
			if (currentSheetDoc == null) {
				continue;
			}
			for (int rowIndex = 0; rowIndex <= currentSheetDoc.getLastRowNum(); rowIndex++) {
				XSSFRow currentRowDoc = currentSheetDoc.getRow(rowIndex);
				if (currentRowDoc == null) {
					continue;
				}

				int currentRowFirstCol = currentRowDoc.getFirstCellNum();
				int currentRowLastCol = currentRowDoc.getLastCellNum();
				List<String> currentRowContent = new ArrayList<String>();

				for (int currentCol = currentRowFirstCol; currentCol < currentRowLastCol; currentCol++) {
					XSSFCell currentCellDoc = currentRowDoc.getCell(currentCol);
					if (currentCellDoc == null) {
						currentRowContent.add(EMPTY);
						continue;
					}
					currentRowContent.add(getXSSFWordbookCellContent(currentCellDoc));
				}
				singleRow.readRow(currentRowContent, rowIndex);
			}
		}
	}

	private static void readCsv(InputStream excelInputStream, CallBackExcelSingleRow1 singleRow) throws Exception {
		BufferedReader csvReader = new BufferedReader(new InputStreamReader(excelInputStream, DEFAULT_CHARSET));
		String readLineContent;
		int rowIndex = 0;
		while ((readLineContent = csvReader.readLine()) != null) {
			String[] rowCodes = StringUtils.split(readLineContent, POINT_CODE);
			List<String> currentRowContent = new ArrayList<String>();
			for (String rowSingleCode : rowCodes) {
				currentRowContent.add(rowSingleCode);
			}
			singleRow.readRow(currentRowContent, rowIndex);
			rowIndex++;
		}
	}

	private static String getHSSFWordbookCellContent(HSSFCell xlsCell) {
		switch (xlsCell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			return xlsCell.getBooleanCellValue() ? "TRUE" : "FALSE";
		case Cell.CELL_TYPE_ERROR:
			return ErrorEval.getText(xlsCell.getErrorCellValue());
		case Cell.CELL_TYPE_FORMULA:
			return xlsCell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(xlsCell)) {
				return DateUtils.Date2String("yyyyMMdd", xlsCell.getDateCellValue());
			}
			xlsCell.setCellType(Cell.CELL_TYPE_STRING);
			return xlsCell.getStringCellValue();
		case Cell.CELL_TYPE_STRING:
			return xlsCell.getStringCellValue();
		default:
			return EMPTY;
		}
	}

	private static String getXSSFWordbookCellContent(XSSFCell xlsxCell) {
		switch (xlsxCell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			return xlsxCell.getBooleanCellValue() ? "TRUE" : "FALSE";
		case Cell.CELL_TYPE_ERROR:
			return ErrorEval.getText(xlsxCell.getErrorCellValue());
		case Cell.CELL_TYPE_FORMULA:
			try {
				return String.valueOf(xlsxCell.getNumericCellValue());
			} catch (IllegalStateException e) {
				return String.valueOf(xlsxCell.getRichStringCellValue());
			}
			// return xlsxCell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(xlsxCell)) {
				return DateUtils.Date2String("yyyyMMdd", xlsxCell.getDateCellValue());
			}
			xlsxCell.setCellType(Cell.CELL_TYPE_STRING);
			return xlsxCell.getStringCellValue();
		case Cell.CELL_TYPE_STRING:
			return xlsxCell.getStringCellValue();
		default:
			return EMPTY;
		}
	}

	public interface CallBackExcelSingleRow1 {
		void readRow(List<String> rowContent, int rowIndex);
	}
}
