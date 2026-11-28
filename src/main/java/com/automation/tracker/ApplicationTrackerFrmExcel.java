package com.automation.tracker;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class ApplicationTrackerFrmExcel {
	
	private static final String TELEGRAM_TOKEN = System.getenv("TELEGRAM_TOKEN");
    private static final String TELEGRAM_CHAT_ID = System.getenv("TELEGRAM_CHAT_ID");
    private static final String EXCEL_FILE_PATH = "watchlist.xlsx"; // Root directory pointer

 // Updated Object: Encapsulates individual URL, Condition, AND XPATH mapping
    public static class TargetMonitor {
        String pageLabel;
        String webpageUrl;
        String expectedTargetCondition;
        String targetXpath; // Added metadata field

        public TargetMonitor(String pageLabel, String webpageUrl, String expectedTargetCondition, String targetXpath) {
            this.pageLabel = pageLabel;
            this.webpageUrl = webpageUrl;
            this.expectedTargetCondition = expectedTargetCondition;
            this.targetXpath = targetXpath;
        }
    }

    public static void main(String[] args) {
    	System.out.println("NEW BUILD - TEST 001 - NOVEMBER 28");
        System.out.println("[*] Initialising Excel-Driven Automation Scraper...");

            // Step A: Dynamically read requirements directly from the spreadsheet file (Runs once per execution)
            List<TargetMonitor> monitoringWatchlist = readWatchlistFromExcel();

            if (monitoringWatchlist.isEmpty()) {
                System.err.println("[-] Warning: Watchlist Excel sheet is empty or missing.");
                System.exit(1); // Exits with a failure code so GitHub log flags it clearly
            } 
            
            System.out.println("[*] Watchlist successfully loaded. Total monitors to evaluate: " + monitoringWatchlist.size());

                // Step B: Loop through all rows captured out of the spreadsheet layout
                for (TargetMonitor target : monitoringWatchlist) {
                	System.out.println("                                                                       ");
                	System.out.println("=======================================================================");
                    System.out.println("[*] Processing Suite: " + target.pageLabel);
                    System.out.println("=======================================================================");
                    boolean conditionMet = verifyPageCondition(target);
                    
                    if (conditionMet) {
						/*
						 * String alertMsg = String.format("🚨 *DATA-DRIVEN EXCEL ALERT!*\n\n" +
						 * "🖥️ *Monitor:* %s\n" + "📅 *Target Found:* %s\n" + "🔗 *Direct Link:* %s",
						 * target.pageLabel, target.expectedTargetCondition, target.webpageUrl);
						 * 
						 * sendTelegramNotification(alertMsg);
						 */
                    	String alertMsg = String.format(
                    		    "🚨 <b>DATA-DRIVEN EXCEL ALERT!</b>\n\n" +
                    		    "🖥️ <b>Monitor:</b> %s\n" +
                    		    "📅 <b>Target Found:</b> %s\n" +
                    		    "🔗 <b>Direct Link:</b> <a href=\"%s\">Book Tickets</a>",

                    		    escapeHtml(target.pageLabel),
                    		    escapeHtml(target.expectedTargetCondition),
                    		    escapeHtml(target.webpageUrl)
                    		);

                    		sendTelegramNotification(alertMsg);
                    }
                }

                System.out.println("[+] Evaluation step complete. Terminating virtual process cleanly.");
                // No while loop, no sleep. The virtual environment shuts down naturally here until the next cron wake-up.
            }

    // 📊 Data Provider Method: Parses the physical Excel Workbook rows cleanly into memory objects
    private static List<TargetMonitor> readWatchlistFromExcel() {
        List<TargetMonitor> list = new ArrayList<>();
        
        System.out.println("[Debug] Looking for Excel file at absolute path: " + new java.io.File(EXCEL_FILE_PATH).getAbsolutePath());
        
        try (FileInputStream file = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(file)) {
            
            Sheet sheet = workbook.getSheetAt(0); // Focus on the first active sheet
            DataFormatter formatter = new DataFormatter(); // Safe conversion tool for numerical cells

            for (int r = 1; r <= sheet.getLastRowNum(); r++) { // Start at 1 to skip headers row
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Safely extract string data from each column index
                String label = formatter.formatCellValue(row.getCell(0)).trim();
                String url = formatter.formatCellValue(row.getCell(1)).trim();
                String condition = formatter.formatCellValue(row.getCell(2)).trim();
                String xpath = formatter.formatCellValue(row.getCell(3)).trim();

                // Skip any partial/blank records in the sheet layouts
                if (!label.isEmpty() && !url.isEmpty() && !condition.isEmpty() && !xpath.isEmpty()) {
                    list.add(new TargetMonitor(label, url, condition, xpath));
                }
            }
        } catch (Exception e) {
            System.err.println("[-] Error parsing Excel Sheet source: " + e.getMessage());
        }
        return list;
    }

    private static boolean verifyPageCondition(TargetMonitor target) {
        WebDriverManager.chromedriver().setup();

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless=new");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

        WebDriver driver = new ChromeDriver(options);
        boolean targetFoundOnPage = false;

        try {
            driver.get(target.webpageUrl);

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(12));

            // CRITICAL FIX: The XPATH is no longer hardcoded! It reads straight from the current row
            By dynamicLocator = By.xpath(target.targetXpath);
            
            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(dynamicLocator));
            List<WebElement> domelementsList = driver.findElements(dynamicLocator);
            int size = domelementsList.size();
            System.out.println("Total no of shows =" + size);
//          //use below for matching condition 
//            for (WebElement element : domelementsList) {
//                String cleanText = element.getText().trim();

//                if (cleanText.contains(target.expectedTargetCondition)) {
//                    System.out.println("[+] Match Found! Text [" + target.expectedTargetCondition + "] visible on " + target.pageLabel);
//                    targetFoundOnPage = true;
//                    break;
//                }
            
         // Convert the String from Excel into a primitive int for mathematical comparison
            int targetCount = Integer.parseInt(target.expectedTargetCondition.trim());
                
                if (size >= targetCount) {
                	System.out.println("[+] Success! Visible elements count (" + size + ") met or exceeded target (" + targetCount + ")");
                  targetFoundOnPage = true;
              }
                else {
                	System.out.println("[+]  Did not match! Visible elements count (" + size + ") is less than target (" + targetCount + ")");
                	System.out.println("The URL: " + driver.getCurrentUrl());
                }
        }
        
            
         catch (Exception e) {
            System.err.println("[-] Script Timeout / Element missing on: " + target.pageLabel +  " using xpath: " + target.targetXpath);
        } finally {
            driver.quit();
        }
        return targetFoundOnPage;
    }
    

	private static void sendTelegramNotification(String message) {

		try {
			String apiUrl = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";

			String payload = "chat_id=" + java.net.URLEncoder.encode(TELEGRAM_CHAT_ID, "UTF-8") + "&text="
					+ java.net.URLEncoder.encode(message, "UTF-8") + "&parse_mode=HTML";

			java.net.URI uri = java.net.URI.create(apiUrl);
			java.net.URL url = uri.toURL();

			java.net.HttpURLConnection conn = (java.net.HttpURLConnection) url.openConnection();

			conn.setRequestMethod("POST");
			conn.setDoOutput(true);

			conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");

			try (java.io.OutputStream os = conn.getOutputStream()) {
				os.write(payload.getBytes(java.nio.charset.StandardCharsets.UTF_8));
			}

			int responseCode = conn.getResponseCode();

			java.io.InputStream stream = responseCode >= 200 && responseCode < 300 ? conn.getInputStream()
					: conn.getErrorStream();

			String response;

			try (java.util.Scanner scanner = new java.util.Scanner(stream, "UTF-8").useDelimiter("\\A")) {
				response = scanner.hasNext() ? scanner.next() : "";
			}

			if (responseCode == 200) {
				System.out.println("[🎉 SUCCESS] Telegram message successfully pushed!");
			} else {
				System.err.println("[❌ ERROR] Telegram HTTP Code: " + responseCode);
				System.err.println("[Telegram Response] " + response);
			}

			conn.disconnect();

		} catch (Exception e) {
			System.err.println("[-] Alert transmission failure: " + e.getMessage());
			e.printStackTrace();
		}
	}
	
	private static String escapeHtml(String text) {
	    if (text == null) return "";

	    return text.replace("&", "&amp;")
	               .replace("<", "&lt;")
	               .replace(">", "&gt;")
	               .replace("\"", "&quot;");
	}

    
}

