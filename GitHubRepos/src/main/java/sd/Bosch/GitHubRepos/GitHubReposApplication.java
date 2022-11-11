package sd.Bosch.GitHubRepos;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.web.client.RestTemplate;

import java.io.*;
import java.util.Iterator;

@SpringBootApplication
public class GitHubReposApplication {
	private static final String gitHubUrl = "https://api.github.com/orgs/bosch-io/repos";
	private static final String [] header = {"ID", "Full Name", "Description", "Programming Language", "Repository URL"};

	public static void main(String[] args) {
		ConfigurableApplicationContext ctx = SpringApplication.run(GitHubReposApplication.class, args);
		saveDataToJsonFile(getBoschRepositories());
		convertJsonFileToExcel();
		ctx.close();
	}
	public static Object[] getBoschRepositories(){
		RestTemplateBuilder builder = new RestTemplateBuilder();
		RestTemplate restTemplate = builder.build();

		int i = 1;
		int length = 0;
		Object[] partialObjects;
		Object[] fullObjects = new Object[0];

		//Traverse the GitHub pages and fetch repos until landing on an empty page, storing the repos as objects in a linked hashmap
		do {
			//Get data from GitHub and store in a partial map
			partialObjects = restTemplate.getForObject(gitHubUrl + String.format("?page=%d", i++), Object[].class);
			assert partialObjects != null;

			//Exit condition
			if (partialObjects.length == 0) break;

			//Increase length for the next map
			length += partialObjects.length;

			//Create temporary map to store the existing object of the fullObjects one
			Object[] tempObjects = new Object[fullObjects.length];
			System.arraycopy(fullObjects, 0, tempObjects, 0, fullObjects.length);

			//Create new fullObjects map with increased length and transfer everything from the temp and partial maps
			fullObjects = new Object[length];
			System.arraycopy(tempObjects, 0, fullObjects, 0, tempObjects.length);
			System.arraycopy(partialObjects, 0, fullObjects, length - partialObjects.length, partialObjects.length);
		} while (true);

		return fullObjects;
	}

	public static void saveDataToJsonFile(Object[] object) {
		try(FileWriter file = new FileWriter("Bosch Git Repos.json")){

			//Create the opening statement for the body
			file.write("{ \"body\" : [ ");
			for (int i = 0; i < object.length; i++) {
				file.write(new Gson().toJson(object[i]));

				//Add separator between objects until the last one is reached
				if (i != object.length - 1) {
					file.write(",");
				}
			}

			//Close the body
			file.write("]}");
			file.flush();
		}catch (IOException ex){
			ex.getCause();
			ex.printStackTrace();
		}
	}

	public static void convertJsonFileToExcel() {
		ObjectMapper om = new ObjectMapper();
		try {
			//Deserialize the Json file
			JsonNode node = om.readTree(new File("Bosch Git Repos.json"));

			//Create & populate the Excel workbook
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("BoschGitRepos");
			createHeader(sheet, wb);
			createBody(node, sheet, wb);

			//Create the Excel file itself
			String excelFilePath = System.getProperty("user.home") + "/Desktop/Bosch.IO Public GitHub Repositories.xlsx";
			FileOutputStream outPutStream = new FileOutputStream(excelFilePath);
			wb.write(outPutStream);
			wb.close();
			outPutStream.close();
		} catch (IOException ex) {
			ex.getCause();
			ex.printStackTrace();
		}
	}

	private static void createHeader(XSSFSheet sheet, XSSFWorkbook wb) {
		//Create the header of the table
		CellStyle headerStyle = getHeaderStyle(wb, sheet);
		Row row = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(header[i]);
			cell.setCellStyle(headerStyle);
		}

		//Apply the filter & freeze the first row & adjust the widths of the columns
		sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, 4));
		sheet.createFreezePane(0, 1);
		sheet.setColumnWidth(0, 2800);
		sheet.setColumnWidth(1, 9000);
		sheet.setColumnWidth(2, 16000);
		sheet.setColumnWidth(3, 7000);
		sheet.setColumnWidth(4, 19000);
	}

	private static void createBody(JsonNode node, XSSFSheet sheet, XSSFWorkbook wb) {
		//Create the body
		JsonNode body = node.get("body");
		int rowNum = 1;
		int i = 0;
		JsonNode rowNode;

		//Get body styles
		CellStyle bodyStyleOdd = getBodyStyleOdd(wb);
		CellStyle bodyStyleEven = getBodyStyleEven(wb);

		//Iterate through the node, creating new rows for each object and adding them to the table
		while (i < body.size()) {

			//Get next repository (object) & create a row for it
			rowNode = body.get(i++);
			Row bodyRow = sheet.createRow(rowNum++);

			int colNum = 0;

			//Create the cells
			Cell idCell = bodyRow.createCell(colNum++);
			Cell nameCell = bodyRow.createCell(colNum++);
			Cell descriptionCell = bodyRow.createCell(colNum++);
			Cell languageCell = bodyRow.createCell(colNum++);
			Cell urlCell = bodyRow.createCell(colNum);

			//Populate the cells
			idCell.setCellValue(rowNode.get("id").asInt());
			nameCell.setCellValue(rowNode.get("full_name").asText());
			descriptionCell.setCellValue(rowNode.get("description") == null ? "" : rowNode.get("description").asText());
			languageCell.setCellValue(rowNode.get("language") == null ? "" : rowNode.get("language").asText());
			urlCell.setCellValue(rowNode.get("url").asText());

			//Set style for each cell
			Iterator<Cell> cellIterator = bodyRow.cellIterator();
			while (cellIterator.hasNext()) {
				if (rowNum % 2 == 0) {
					cellIterator.next().setCellStyle(bodyStyleEven);
				} else {
					cellIterator.next().setCellStyle(bodyStyleOdd);
				}
			}

		}
	}

	private static CellStyle getHeaderStyle(XSSFWorkbook wb, XSSFSheet sheet) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		XSSFFont font = wb.createFont();
		font.setBold(true);
		cellStyle.setFont(font);
		cellStyle.setBorderBottom(BorderStyle.THICK);
		cellStyle.setBorderRight(BorderStyle.THICK);
		cellStyle.setFillForegroundColor(new XSSFColor(new byte[] {(byte) 155, (byte) 188, (byte) 255}));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		sheet.setDisplayGridlines(false);
		return cellStyle;
	}

	private static CellStyle getBodyStyleOdd(XSSFWorkbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(new XSSFColor(new byte[] {(byte) 193, (byte) 208, (byte) 255}));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return cellStyle;
	}

	private static CellStyle getBodyStyleEven(XSSFWorkbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setWrapText(true);
		cellStyle.setAlignment(HorizontalAlignment.LEFT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(new XSSFColor(new byte[] {(byte) 234, (byte) 234, (byte) 234}));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return cellStyle;
	}
}

