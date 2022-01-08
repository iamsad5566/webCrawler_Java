import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.TimeUnit;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import java.lang.*;
import java.io.FileOutputStream;
import org.jsoup.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Crawler {
        // Create a Workbook first so that all methods below can use
        static Workbook wb = new HSSFWorkbook();

    public static void main(String[] args) throws Exception {
        // Target month and last month
        String[] yearAndMonth = thisMonth();
        String thisMonth = yearAndMonth[0]+yearAndMonth[1];
        String lastMonth = lastMonth(yearAndMonth[0], yearAndMonth[1]);
        System.out.println(thisMonth);
        System.out.println(lastMonth);

        // Set the pathway of the chromedriver.exe and create a driver
        String path = System.getProperty("user.dir");
        System.setProperty("webdriver.chrome.driver", path + "/chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        // List 存放受試者費代墊
        var list_experiment = browser_experiment(driver, thisMonth, lastMonth);

        // Add the sheets
        Sheet first = wb.createSheet("各類所得（受試者費）");
        Sheet second = wb.createSheet("各計畫當月總支出");
        Sheet third = wb.createSheet("受試者費以外的支出");
        Sheet fourth = wb.createSheet("個人當月代墊總和");

        // Build the first sheet
        buildSheet1(list_experiment, first);

        // 移到報帳管理，建立sheet2
        var list_funding = browser_funding(driver, thisMonth, lastMonth);

        // Build the second sheet
        buildSheet2(list_funding, second);

        // Build the third sheet，計算實驗室該月除了受試者費外的支出
        buildSheet3(first, second, third);
        // 計算每個人當月代墊
        buildSheet4(first, third, fourth);


        StringBuilder sb = new StringBuilder();
        String[] date_arr = new Date().toString().split(" ");
        sb.append(date_arr[1]).append('_').append(date_arr[2]);

        FileOutputStream fileOut = new FileOutputStream(path + sb.toString() + ".xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static String[] thisMonth() {
        Scanner scanner = new Scanner(System.in);
        System.out.print("請輸入結帳年度：");
        String year = scanner.next();
        System.out.printf("請輸入結帳月份：");
        String month = scanner.next();

        // 開頭加上年份
        StringBuilder sb = new StringBuilder(year);
        if(Integer.valueOf(month) < 10)
            month = "0" + month;

        String[] res = {year, month};

        return  res;
    }

    public static String lastMonth(String y, String thisMonth) {
        String year = y;
        String month = "";
        int prev = Integer.valueOf(thisMonth) - 1;
        if(prev == 0) {
            prev = 12;
            month = Integer.toString(prev);
            year = Integer.toString(Integer.valueOf(y)-1);
        }

        else if(prev < 10)
            month = "0" + Integer.toString(prev);

        else
            month = Integer.toString(prev);

        return year+month;
    }

    public static List<String> browser_experiment(WebDriver driver, String thisMoth, String lastMonth) {
        // Direct to NTU web
        driver.get("https://ntuacc.cc.ntu.edu.tw/acc/index.asp?campno=m&idtype=3");
        // Let's Login
        driver.findElement(new By.ByXPath("//*[@id=\"bossid\"]")).sendKeys("");
        driver.findElement(new By.ByXPath("//*[@id=\"assid\"]")).sendKeys("");
        driver.findElement(new By.ByXPath("//*[@id=\"asspwd\"]")).sendKeys("");
        driver.findElement(new By.ByXPath("//*[@id=\"vsub\"]/td/input")).click();

        // Navigate to 各類所得
        driver.navigate().to("https://ntuacc.cc.ntu.edu.tw/acc/salary/variousal.asp");

        // List to return
        var list = new ArrayList<String>();

        // Get all in a month
        outer:for(;;) {
            var tr_tag = driver.findElements(By.tagName("tr"));
            for(WebElement e:tr_tag) {
                if(e.getText().startsWith("00") && e.getText().split(" ")[3].startsWith("受試者費")) {

                    // 如果找到到帳日為上個月份，跳出迴圈
                    if(e.getText().split(" ")[5].startsWith(lastMonth))
                        break outer;

                    else if(!e.getText().split(" ")[5].startsWith(thisMoth) || (e.getText().split(" ").length < 10))
                        continue;

                    list.add(e.getText());
                    System.out.println(e.getText());
                }
            }
            WebElement nextPage = driver.findElement(By.linkText("下一頁"));
            nextPage.click();
        }

        return list;
    }

    public static List<String> browser_funding(WebDriver driver, String thisMonth, String lastMonth) throws InterruptedException, IOException {
        //  移動到報帳管理/計畫經費報帳
        driver.navigate().to("https://ntuacc.cc.ntu.edu.tw/acc/apply/list.asp");

        // List to return
        var list = new ArrayList<String>();

        // 直到抓到上月份的經費，break
        // 換網頁(網域)後會刪除原先儲存的資料，所以每換一次要重抓一次
        outter:for(; ;) {
            var plans = driver.findElements(By.tagName("tr"));
            for(int i = 0; i < plans.size(); i++) {
//                if(plans.get(i).getText().startsWith("110T217C570"))
//                    continue;
                if(plans.get(i).getText().startsWith("110T")) {
                    StringBuilder sb = new StringBuilder();

                    for(String s:plans.get(i).getText().split(" ")) {
                        if(!s.equals("各類所得") && !s.equals("勞健保月薪(勞退新制)")) {
                            sb.append(s).append(" ");
                        }
                    }

                    if(sb.toString().split(" ")[5].startsWith(lastMonth))
                        break outter;

                    if(sb.toString().split(" ").length < 11 || !sb.toString().split(" ")[5].startsWith(thisMonth))
                        continue;

                    // 點進去找 name
                    Actions act = new Actions(driver);
                    act.doubleClick(plans.get(i)).perform();
                    String nameInfo = driver.findElement(By.xpath("/html/body/center/center/table[1]/tbody/tr[2]/td[2]")).getText();
                    sb.append(nameInfo);
                    // 以上為與 js 互動的 part

                    // 存入 list
                    list.add(sb.toString());
                    System.out.println(sb);
                    driver.findElement(By.name("back")).click();

                    // 重抓一次所有計畫
                    plans = driver.findElements(By.tagName("tr"));
                }
            }

            WebElement nextPage = driver.findElement(By.linkText("下一頁"));
            nextPage.click();
        }

        Thread.sleep(1000);
        driver.quit();
        return list;
    }

    // First sheet
    public static void buildSheet1(List<String> list, Sheet first) throws Exception{
        //顯示標題
        Row title_row = first.createRow(0);
        title_row.setHeight((short)(40*20));
        Cell title_cell = title_row.createCell(0);

        // Set up the header
        String headers[] = new String[]{"出納組編號","報帳條碼","科目代碼","說明", "金額", "到帳日", "目前狀態", "傳票號碼", "付款資料", "列印次數"};
        Row header_row = first.createRow(1);
        header_row.setHeight((short)(20*24));

        //建立單元格的 顯示樣式
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER); //水平方向上的對其方式
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);	//垂直方向上的對其方式

        title_cell.setCellStyle(style);
        title_cell.setCellValue("各類所得（受試者費）");

        first.addMergedRegion(new CellRangeAddress(0,0,0,headers.length-1));

        // 把 header 的每一個 string 加上
        for(int i=0;i<headers.length;i++){
            //設定列寬   基數為256
            first.setColumnWidth(i, 30*256);
            Cell cell = header_row.createCell(i);
            //應用樣式到  單元格上
            cell.setCellStyle(style);
            cell.setCellValue(headers[i]);
        }

        // Fill in the data
        for(int i=0;i< list.size();i++){
            Row row = first.createRow(i+2);
            row.setHeight((short)(20*20)); //設定行高  基數為20
            for(int j=0;j<list.get(i).split(" ").length;j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(list.get(i).split(" ")[j]);
            }
        }
    }

    // Second sheet
    public static void buildSheet2(List<String> list, Sheet second) {
        // Sheet2 title
        Row title_row = second.createRow(0);
        title_row.setHeight((short)(40*20));
        Cell title_cell = title_row.createCell(0);

        String headers[] = new String[]{"報帳條碼","經費或計畫名稱","計畫代碼", "經費別", "金額", "報帳日", "傳票號碼", "付款資料", "報帳ID", "列印次數", "受款人", "備註"};
        Row header_row = second.createRow(1);
        header_row.setHeight((short)(20*24));

        //建立單元格的 顯示樣式
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER); //水平方向上的對其方式
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);	//垂直方向上的對其方式
        title_cell.setCellStyle(style);
        title_cell.setCellValue("計畫經費報帳");
        second.addMergedRegion(new CellRangeAddress(0,0,0,headers.length-1));

        for(int i=0;i<headers.length;i++){
            //設定列寬   基數為256
            second.setColumnWidth(i, 30*256);
            Cell cell = header_row.createCell(i);
            //應用樣式到  單元格上
            cell.setCellStyle(style);
            cell.setCellValue(headers[i]);
        }

        for(int i=0;i< list.size();i++){
            Row row = second.createRow(i+2);
            row.setHeight((short)(20*20)); //設定行高  基數為20
            for(int j=0;j<list.get(i).split(" ").length;j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(list.get(i).split(" ")[j]);
            }
        }
    }

    public static void buildSheet3(Sheet first, Sheet second, Sheet third) {
        String headers[] = new String[]{"報帳條碼","經費或計畫名稱","計畫代碼", "經費別", "金額", "報帳日", "傳票號碼", "付款資料", "報帳ID", "列印次數", "受款人", "備註"};
        Row header_row = third.createRow(0);
        header_row.setHeight((short)(20*24));

        //建立單元格的 顯示樣式
        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER); //水平方向上的對其方式
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);	//垂直方向上的對其方式

        for(int i=0;i<headers.length;i++){
            //設定列寬   基數為256
            third.setColumnWidth(i, 30*256);
            Cell cell = header_row.createCell(i);
            //應用樣式到  單元格上
            cell.setCellValue(headers[i]);
            cell.setCellStyle(style);
        }

        var barCode = new ArrayList<String>();
        for(int i = 2; first.getRow(i) != null; i++)
            barCode.add(first.getRow(i).getCell(1).getStringCellValue());

        int targetRow = 1;
        for(int i = 2; second.getRow(i) != null; i++) {
            if(barCode.contains(second.getRow(i).getCell(0).getStringCellValue()))
                continue;

            Row row = third.createRow(targetRow);
            row.setHeight((short)(20*20));
            for(int j = 0; j < headers.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(second.getRow(i).getCell(j).getStringCellValue());
                cell.setCellStyle(style);
            }
            targetRow++;
        }
    }

    public static void buildSheet4(Sheet first, Sheet third, Sheet fourth) {
        String[] header = {"彥茹", "欣蓉", "彥匡", "宇安", "昀安", "宇昕", "林懿", "品程"};
        Row header_row = fourth.createRow(0);
        header_row.setHeight((short)(20*24));

        for(int i = 0; i < header.length; i++) {
            Cell cell = header_row.createCell(i);
            cell.setCellValue(header[i]);
        }

        // 要算每個人代墊筆數的 array
        int[] rowNum = new int[header.length];
        Arrays.fill(rowNum, 1);

        int index = 1;
        for(int i = 2; first.getRow(i) != null; i++, index++) {
            Row row = fourth.createRow(index);
            row.setHeight((short)(20*20));
            for(int j = 0; j < header.length; j++)
                row.createCell(j);
        }

        for(int i = 1; third.getRow(i) != null; i++, index++) {
            Row row = fourth.createRow(index);
            row.setHeight((short)(20*20));
            for(int j = 0; j < header.length; j++)
                row.createCell(j);
        }

        //  從第一個 column 開始
        for(int i = 0; i < header.length; i++) {

            // 抓 first sheet，第二個 row 以後的資料
            for(int j = 2; first.getRow(j) != null; j++) {
                Cell target = fourth.getRow(rowNum[i]).getCell(i);
                String name = first.getRow(j).getCell(3).getStringCellValue();
                String money = first.getRow(j).getCell(4).getStringCellValue();
                if(name.contains(header[i])) {
                    target.setCellValue(money);
                    rowNum[i]++;
                }
            }

            // 抓 third sheet 第一個 row 以後的資料
            for(int j = 1; third.getRow(j) != null; j++) {
                Cell target = fourth.getRow(rowNum[i]).getCell(i);
                String name = third.getRow(j).getCell(11).getStringCellValue();
                String money = third.getRow(j).getCell(4).getStringCellValue();
                if(name.contains(header[i])) {
                    target.setCellValue(money);
                    rowNum[i]++;
                }
            }
        }

        // 計算每個人代墊總金額
        Row row = fourth.createRow(index);
        row.setHeight((short)(20*20));

        // 從第一個 column 開始
        for(int i = 0; i < header.length; i++) {
            int sum = 0;

            // rowNum 記錄每一個人代墊幾筆
            for(int j = 1; j <= rowNum[i]; j++) {
                String sValue = fourth.getRow(j).getCell(i).getStringCellValue();
                StringBuilder sb = new StringBuilder();
                for(char c:sValue.toCharArray()) {
                    if (Character.isDigit(c))
                        sb.append(c);
                }
                if(sb.length()!= 0)
                    sum += Integer.valueOf(sb.toString());
            }
            Cell cell = fourth.getRow(index).createCell(i);
            cell.setCellValue(sum);
        }
    }
}