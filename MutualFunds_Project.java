package com.mutualfunds;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class MutualFunds_Project {
	static WebDriver driver;
	static XSSFWorkbook wb;

	public static void main(String[] args) throws InterruptedException {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.morningstar.in/funds.aspx");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		String FundName = "Parag Parikh Flexi Cap Regular Growth";

		driver.findElement(By.xpath("//*[@id=\"ctl00_ucHeader_txtQuote_txtAutoComplete\"]")).sendKeys(FundName);

		Thread.sleep(2000);
		driver.findElement(By.xpath("//strong[normalize-space()='" + FundName + "']")).click();

		WebElement FundNameElement = driver
				.findElement(By.xpath("//*[@id=\"page-section\"]/div[2]/div/div/div[1]/div/div[1]/div/div[2]/h1"));
		String FundNameText = FundNameElement.getText();

		WebElement TotalAssetElement = driver.findElement(By.xpath(
				"//*[@id=\"ctl00_ContentPlaceHolder1_ucQuoteHeader_mipQuoteDiv\"]/div/div/div/div[2]/div/div/div/div[2]/div[2]/div/div[2]"));
		String TotalAsset = TotalAssetElement.getText();

		WebElement CategoryElement = driver.findElement(By.xpath(
				"//*[@id=\"ctl00_ContentPlaceHolder1_ucQuoteHeader_mipQuoteDiv\"]/div/div/div/div[2]/div/div/div/div[2]/div[7]/div/div[2]"));
		String Category = CategoryElement.getText();

		WebElement InvestmentElement = driver.findElement(By.xpath(
				"//*[@id=\"ctl00_ContentPlaceHolder1_ucQuoteHeader_mipQuoteDiv\"]/div/div/div/div[2]/div/div/div/div[2]/div[8]/div/div[2]"));
		String InvestmentStyle = InvestmentElement.getText();

		wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Headers");

		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Mutual_FundName");
		header.createCell(1).setCellValue("TotalAsset");
		header.createCell(2).setCellValue("Category");
		header.createCell(3).setCellValue("InvestmentStyle");

		Row datarow = sheet.createRow(1);
		datarow.createCell(0).setCellValue(FundNameText);
		datarow.createCell(1).setCellValue(TotalAsset);
		datarow.createCell(2).setCellValue(Category);
		datarow.createCell(3).setCellValue(InvestmentStyle);

		for (int i = 0; i < 4; i++) {
			sheet.autoSizeColumn(i);
		}

		getAssetAllocation();
		getStyleMeasures();
		getRiskAndVolatilityfor3yr();
		getMarketVolatility3yr();
		getRiskAndVolatilityfor5yr();
		getMarketVolatility5yr();
		getRiskAndVolatilityfor10yr();
		getMarketVolatility10yr();
		try {
			File file = new File("MutualFunds.xlsx");
			FileOutputStream fos = new FileOutputStream(file);
			wb.write(fos);
			fos.close();
			wb.close();
		} catch (Exception e) {
			e.getMessage();
		}
		driver.close();
	}

	public static void getAssetAllocation() {

		driver.findElement(By.xpath("//*[@id=\"ctl00_ContentPlaceHolder1_ucNavigation_rptNavigation_ctl01_lnkTab\"]"))
				.click();

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id='quotePageContent']/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/thead/tr/th"));
		System.out.println(TableRowHeader.size());
		Sheet sheet1 = wb.createSheet("Asset Allocation");
		Row Header1 = sheet1.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/thead/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);

		}

		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id='quotePageContent']/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/tbody/tr/th"));
		Row Header2;
		for (int i = 0; i < BodyHeaders.size(); i++) {
			String RowHeader = driver.findElement(By.xpath(
					"//*[@id='quotePageContent']/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/tbody/tr["
							+ (i + 1) + "]/th"))
					.getText();
			Header2 = sheet1.createRow(i + 1);
			Header2.createCell(0).setCellValue(RowHeader);
		}

		List<WebElement> ColumnDataElements = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/tbody/tr/td"));
		System.out.println(ColumnDataElements.size());

		List<WebElement> EachRowColumnData = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/tbody/tr[1]/td"));

		for (int i = 0; i < BodyHeaders.size(); i++) {
			Header2 = sheet1.getRow(i + 1);

			for (int j = 0; j < EachRowColumnData.size(); j++) {
				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[3]/div[1]/div/div/div/div/div[2]/div/div[1]/div/div/div[1]/div/div[2]/table/tbody/tr["
								+ (i + 1) + "]/td[" + (j + 1) + "]"))
						.getText();
				Header2.createCell(j + 1).setCellValue(ColumnData);

			}

		}

		for (int i = 0; i < TableRowHeader.size(); i++) {
			sheet1.autoSizeColumn(i);
		}

	}

	public static void getStyleMeasures() {

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/thead/tr/th"));
		Sheet sheet2 = wb.createSheet("StyleMeasures");
		Row Header1 = sheet2.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/thead/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);

		}

		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/tbody/tr/th"));
		Row Header2;
		for (int i = 0; i < BodyHeaders.size(); i++) {
			String RowHeader = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/tbody/tr["
							+ (i + 1) + "]/th"))
					.getText();
			Header2 = sheet2.createRow(i + 1);
			Header2.createCell(0).setCellValue(RowHeader);
		}

		List<WebElement> ColumnDataElements = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/tbody/tr/td"));
		System.out.println(ColumnDataElements.size());

		List<WebElement> EachRowColumnData = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/tbody/tr[1]/td"));

		for (int i = 0; i < BodyHeaders.size(); i++) {
			Header2 = sheet2.getRow(i + 1);

			for (int j = 0; j < EachRowColumnData.size(); j++) {
				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div[1]/div/div/div/div/div[4]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/div[2]/div/div/div/table/tbody/tr["
								+ (i + 1) + "]/td[" + (j + 1) + "]"))
						.getText();
				Header2.createCell(j + 1).setCellValue(ColumnData);

			}

		}

		for (int i = 0; i < TableRowHeader.size(); i++) {
			sheet2.autoSizeColumn(i);
		}

	}

	public static void getRiskAndVolatilityfor3yr() {

		driver.findElement(By.xpath("//*[@id=\"ctl00_ContentPlaceHolder1_ucNavigation_rptNavigation_ctl04_lnkTab\"]"))
				.click();

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th"));
		Sheet sheet3 = wb.createSheet("Risk & Volatility for 3yr");
		Row Header1 = sheet3.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);

		}
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/th"));
		Row Header2;
		for (int i = 0; i < BodyHeaders.size(); i++) {
			String RowHeader = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
							+ (i + 1) + "]/th"))
					.getText();
			Header2 = sheet3.createRow(i + 1);
			Header2.createCell(0).setCellValue(RowHeader);
		}

		List<WebElement> ColumnDataElements = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/td"));
		System.out.println(ColumnDataElements.size());

		List<WebElement> EachRowColumnData = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[1]/td"));

		for (int i = 0; i < BodyHeaders.size(); i++) {
			Header2 = sheet3.getRow(i + 1);

			for (int j = 0; j < EachRowColumnData.size(); j++) {
				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
								+ (i + 1) + "]/td[" + (j + 1) + "]"))
						.getText();
				Header2.createCell(j + 1).setCellValue(ColumnData);

			}

		}

		for (int i = 0; i < TableRowHeader.size(); i++) {
			sheet3.autoSizeColumn(i);
		}

	}

	public static void getMarketVolatility3yr() {

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th"));
		Sheet sheet4 = wb.createSheet("Market Volatility for 3yr");
		Row Header1 = sheet4.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);
		}

		Row Header2;
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr/th[1]"));

		for (int l = 0; l < BodyHeaders.size(); l++) {

			Header2 = sheet4.createRow(l + 1);
			String BodyHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr["
							+ (l + 1) + "]/th[1]"))
					.getText();
			Header2.createCell(0).setCellValue(BodyHeaderData);
		}
		List<WebElement> BodyColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr/td"));
		System.out.println(BodyColumn.size());
		List<WebElement> EachRowColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr[1]/td"));

		for (int m = 0; m < BodyHeaders.size(); m++) {
			Header2 = sheet4.getRow(m + 1);
			for (int n = 0; n < EachRowColumn.size(); n++) {

				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr["
								+ (m + 1) + "]/td[" + (n + 1) + "]"))
						.getText();

				Header2.createCell(n + 1).setCellValue(ColumnData);

			}
		}
		List<WebElement> RowHeader2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th"));

		Row Header3 = sheet4.createRow(3);
		for (int o = 0; o < RowHeader2.size(); o++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th["
							+ (o + 1) + "]"))
					.getText();

			Header3.createCell(o).setCellValue(RowHeaderData);
		}

		String RowHeader2Data = driver.findElement(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/th"))
				.getText();
		Row Header4 = sheet4.createRow(4);

		Header4.createCell(0).setCellValue(RowHeader2Data);

		List<WebElement> BodyColumn2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td"));

		Header4 = sheet4.getRow(4);
		for (int p = 0; p < BodyColumn2.size(); p++) {
			String BodyColumnData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td["
							+ (p + 1) + "]"))
					.getText();
			Header4.createCell(p + 1).setCellValue(BodyColumnData);

		}
		List<WebElement> RowHeader3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th"));

		Row Header5 = sheet4.createRow(5);
		for (int q = 0; q < RowHeader3.size(); q++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th["
							+ (q + 1) + "]"))
					.getText();

			Header5.createCell(q).setCellValue(RowHeaderData);
		}

		List<WebElement> BodyColumn3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td"));

		Row Header6 = sheet4.createRow(6);

		for (int r = 0; r < BodyColumn3.size(); r++) {
			String BodyColumnData3 = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td["
							+ (r + 1) + "]"))
					.getText();
			System.out.println(BodyColumnData3);
			Header6.createCell(r).setCellValue(BodyColumnData3);
		}

		for (int i = 0; i < 4; i++) {
			sheet4.autoSizeColumn(i);
		}

	}

	public static void getRiskAndVolatilityfor5yr() {

		driver.findElement(By.id("for5Year")).click();

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th"));
		Sheet sheet5 = wb.createSheet("RiskAndVolatility 5yr");
		Row Header1 = sheet5.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);

		}
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/th"));
		Row Header2;
		for (int i = 0; i < BodyHeaders.size(); i++) {
			String RowHeader = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
							+ (i + 1) + "]/th"))
					.getText();
			Header2 = sheet5.createRow(i + 1);
			Header2.createCell(0).setCellValue(RowHeader);
		}

		List<WebElement> ColumnDataElements = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/td"));
		System.out.println(ColumnDataElements.size());

		List<WebElement> EachRowColumnData = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[1]/td"));

		for (int i = 0; i < BodyHeaders.size(); i++) {
			Header2 = sheet5.getRow(i + 1);

			for (int j = 0; j < EachRowColumnData.size(); j++) {
				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
								+ (i + 1) + "]/td[" + (j + 1) + "]"))
						.getText();
				Header2.createCell(j + 1).setCellValue(ColumnData);

			}

		}

		for (int i = 0; i < TableRowHeader.size(); i++) {
			sheet5.autoSizeColumn(i);
		}

	}

	public static void getMarketVolatility5yr() {
		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th"));
		Sheet sheet6 = wb.createSheet("Market Volatility for 5yr");
		Row Header1 = sheet6.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);
		}

		Row Header2;
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr/th[1]"));

		for (int l = 0; l < BodyHeaders.size(); l++) {

			Header2 = sheet6.createRow(l + 1);
			String BodyHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr["
							+ (l + 1) + "]/th[1]"))
					.getText();
			Header2.createCell(0).setCellValue(BodyHeaderData);
		}
		List<WebElement> BodyColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr/td"));
		System.out.println(BodyColumn.size());
		List<WebElement> EachRowColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr[1]/td"));

		for (int m = 0; m < BodyHeaders.size(); m++) {
			Header2 = sheet6.getRow(m + 1);
			for (int n = 0; n < EachRowColumn.size(); n++) {

				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr["
								+ (m + 1) + "]/td[" + (n + 1) + "]"))
						.getText();

				Header2.createCell(n + 1).setCellValue(ColumnData);

			}
		}
		List<WebElement> RowHeader2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th"));

		Row Header3 = sheet6.createRow(3);
		for (int o = 0; o < RowHeader2.size(); o++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th["
							+ (o + 1) + "]"))
					.getText();

			Header3.createCell(o).setCellValue(RowHeaderData);
		}

		String RowHeader2Data = driver.findElement(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/th"))
				.getText();
		Row Header4 = sheet6.createRow(4);

		Header4.createCell(0).setCellValue(RowHeader2Data);

		List<WebElement> BodyColumn2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td"));

		Header4 = sheet6.getRow(4);
		for (int p = 0; p < BodyColumn2.size(); p++) {
			String BodyColumnData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td["
							+ (p + 1) + "]"))
					.getText();
			Header4.createCell(p + 1).setCellValue(BodyColumnData);

		}
		List<WebElement> RowHeader3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th"));

		Row Header5 = sheet6.createRow(5);
		for (int q = 0; q < RowHeader3.size(); q++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th["
							+ (q + 1) + "]"))
					.getText();

			Header5.createCell(q).setCellValue(RowHeaderData);
		}

		List<WebElement> BodyColumn3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td"));

		Row Header6 = sheet6.createRow(6);

		for (int r = 0; r < BodyColumn3.size(); r++) {
			String BodyColumnData3 = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td["
							+ (r + 1) + "]"))
					.getText();
			System.out.println(BodyColumnData3);
			Header6.createCell(r).setCellValue(BodyColumnData3);
		}

		for (int i = 0; i < 4; i++) {
			sheet6.autoSizeColumn(i);
		}

	}

	public static void getRiskAndVolatilityfor10yr() {

		driver.findElement(By.id("for10Year")).click();

		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th"));
		Sheet sheet7 = wb.createSheet("RiskAndVolatility 10yr");
		Row Header1 = sheet7.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/thead/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);

		}
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/th"));
		Row Header2;
		for (int i = 0; i < BodyHeaders.size(); i++) {
			String RowHeader = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
							+ (i + 1) + "]/th"))
					.getText();
			Header2 = sheet7.createRow(i + 1);
			Header2.createCell(0).setCellValue(RowHeader);
		}

		List<WebElement> ColumnDataElements = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/td"));
		System.out.println(ColumnDataElements.size());

		List<WebElement> EachRowColumnData = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[1]/td"));

		for (int i = 0; i < BodyHeaders.size(); i++) {
			Header2 = sheet7.getRow(i + 1);

			for (int j = 0; j < EachRowColumnData.size(); j++) {
				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr["
								+ (i + 1) + "]/td[" + (j + 1) + "]"))
						.getText();
				Header2.createCell(j + 1).setCellValue(ColumnData);

			}
		}

		for (int i = 0; i < TableRowHeader.size(); i++) {
			sheet7.autoSizeColumn(i);
		}

	}

	public static void getMarketVolatility10yr() {
		List<WebElement> TableRowHeader = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th"));
		Sheet sheet8 = wb.createSheet("Market Volatility for 10yr");
		Row Header1 = sheet8.createRow(0);

		for (int i = 0; i < TableRowHeader.size(); i++) {
			String Headers = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[1]/tr/th["
							+ (i + 1) + "]"))
					.getText();
			Header1.createCell(i).setCellValue(Headers);
		}

		Row Header2;
		List<WebElement> BodyHeaders = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr/th[1]"));

		for (int l = 0; l < BodyHeaders.size(); l++) {

			Header2 = sheet8.createRow(l + 1);
			String BodyHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table/tbody[1]/tr["
							+ (l + 1) + "]/th[1]"))
					.getText();
			Header2.createCell(0).setCellValue(BodyHeaderData);
		}
		List<WebElement> BodyColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr/td"));
		System.out.println(BodyColumn.size());
		List<WebElement> EachRowColumn = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr[1]/td"));

		for (int m = 0; m < BodyHeaders.size(); m++) {
			Header2 = sheet8.getRow(m + 1);
			for (int n = 0; n < EachRowColumn.size(); n++) {

				String ColumnData = driver.findElement(By.xpath(
						"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[1]/tr["
								+ (m + 1) + "]/td[" + (n + 1) + "]"))
						.getText();

				Header2.createCell(n + 1).setCellValue(ColumnData);

			}
		}
		List<WebElement> RowHeader2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th"));

		Row Header3 = sheet8.createRow(3);
		for (int o = 0; o < RowHeader2.size(); o++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/thead[2]/tr/th["
							+ (o + 1) + "]"))
					.getText();

			Header3.createCell(o).setCellValue(RowHeaderData);
		}

		String RowHeader2Data = driver.findElement(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/th"))
				.getText();
		Row Header4 = sheet8.createRow(4);

		Header4.createCell(0).setCellValue(RowHeader2Data);

		List<WebElement> BodyColumn2 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td"));

		Header4 = sheet8.getRow(4);
		for (int p = 0; p < BodyColumn2.size(); p++) {
			String BodyColumnData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[1]/tbody[2]/tr/td["
							+ (p + 1) + "]"))
					.getText();
			Header4.createCell(p + 1).setCellValue(BodyColumnData);

		}
		List<WebElement> RowHeader3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th"));

		Row Header5 = sheet8.createRow(5);
		for (int q = 0; q < RowHeader3.size(); q++) {

			String RowHeaderData = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/thead/tr/th["
							+ (q + 1) + "]"))
					.getText();

			Header5.createCell(q).setCellValue(RowHeaderData);
		}

		List<WebElement> BodyColumn3 = driver.findElements(By.xpath(
				"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td"));

		Row Header6 = sheet8.createRow(6);

		for (int r = 0; r < BodyColumn3.size(); r++) {
			String BodyColumnData3 = driver.findElement(By.xpath(
					"//*[@id=\"quotePageContent\"]/div/div/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/div/div/table[2]/tbody/tr/td["
							+ (r + 1) + "]"))
					.getText();
			System.out.println(BodyColumnData3);
			Header6.createCell(r).setCellValue(BodyColumnData3);
		}

		for (int i = 0; i < 4; i++) {
			sheet8.autoSizeColumn(i);
		}

	}

}

