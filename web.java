package sel;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.ProgressMonitor;
import javax.swing.UIManager;
import java.awt.Desktop;
import java.io.File;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.*;
import java.io.IOException;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.biff.RowsExceededException;
public class web extends JFrame implements ActionListener {
static WebDriver driver = new ChromeDriver();
static ProgressMonitor pbar;
static int counter = 0;
static float total = 0, inc = 0;
static JButton b,e;  
static Label label;
private enum Actions {b,e}
public web() {pbar = new ProgressMonitor(null, "Monitoring Progress","Initializing . . .", 0, 100);
b = new JButton();
b.addActionListener(this);
b.setActionCommand(Actions.b.name());
e = new JButton();
e.addActionListener(this);
e.setActionCommand(Actions.e.name());}
public static void main(String[] args) throws WriteException, RowsExceededException, BiffException, IOException{
String inp[][] = new String[100][10], rfn, t;
int i,j,k,l,r=0;		
Workbook rb = Workbook.getWorkbook(new File("Book1.xls"));
Sheet sh = rb.getSheet("Sheet1");
int totalNoOfRows = sh.getRows();
int totalNoOfCols = sh.getColumns();
for (int row = 0; row < totalNoOfRows; row++) {
for (int col = 0; col < totalNoOfCols; col++) inp[row][col] = sh.getCell(col, row).getContents();}
System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
driver.manage().window().setPosition(new Point(-2000, 0));
UIManager.put("ProgressMonitor.progressText", "Excecuting");
UIManager.put("OptionPane.cancelButtonText", "Stop");
new web();
if(inp[0][1].equalsIgnoreCase(""))
driver.get("https://ipindiaonline.gov.in/tmrpublicsearch/frmmain.aspx");
else
driver.get(inp[0][1]);
final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HH-mm");
Timestamp timestamp = new Timestamp(System.currentTimeMillis());
if(inp[1][1].equalsIgnoreCase(""))
rfn = "Results"+sdf.format(timestamp)+".xls";
else
rfn = inp[1][1]+sdf.format(timestamp)+".xls";
if(!inp[2][1].equalsIgnoreCase(""))
rfn = inp[2][1]+"\\"+rfn;
File exlFile = new File(rfn);
WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);
WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
inc = 100/(float)((totalNoOfRows-5)*45);
for(i=5 ; i<totalNoOfRows ; i++){
if(!(inp[i][0].matches("[0-9]+") || inp[i][0].matches("[a-z]+") || inp[i][0].matches("[A-Z]+"))){
label = new Label(0, r++, "Blank Search Key");
writableSheet.addCell(label);
label = new Label(0, r++, "");
writableSheet.addCell(label);
label = new Label(0, r++, "");
writableSheet.addCell(label);
total+=100/(float)((totalNoOfRows-5));
continue;}
label = new Label(0, r++, inp[i][0]);
writableSheet.addCell(label);
if(inp[i][1].equalsIgnoreCase("p")){
if(inp[i][0].length()>2){
for(j=1 ; j<46 ; j++){
label = new Label(0, r++, "Class"+Integer.toString(j));
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_DDLSearchType")).click();
new Select(driver.findElement(By.id("ContentPlaceHolder1_DDLSearchType"))).selectByVisibleText("Phonetic");
driver.findElement(By.xpath("//option[@value='PH']")).click();
driver.findElement(By.xpath("//option[@value='PH']")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBPhonetic")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBPhonetic")).sendKeys(inp[i][0]);
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).sendKeys(Integer.toString(j));
driver.findElement(By.id("ContentPlaceHolder1_BtnSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
String results[] = t.split("\n");
if(!(results[0].contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
for(k=4; k<results.length-5; k++){
if(results[k].contains("International Non-Proprietary Names(INN)")){
label = new Label(0, r++, "International Non-Proprietary Names(INN)");
writableSheet.addCell(label);}
else if(results[k].contains("No Record found")){
label = new Label(0, r++, "No Record found");
writableSheet.addCell(label);}
else if(results[k].contains("Matching Trademark(s)")){
label = new Label(0, r++, "Matching Trademark(s)");
writableSheet.addCell(label);}
else if(results[k].contains("Wordmark:")){
for(l=0 ; l<5 ; l++){
label = new Label(l, r, results[k++]);
writableSheet.addCell(label);}
r++;}}
label = new Label(0, r++, "");
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_LnkNextSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
total += inc;
counter = (int)total;
b.doClick();}}
else{
label = new Label(0, r++, "Invalid Phonetic");
writableSheet.addCell(label);
label = new Label(0, r++, "Enter atleast 3 characters");
writableSheet.addCell(label);
total+=100/(float)((totalNoOfRows-5));
label = new Label(0, r++, "");
writableSheet.addCell(label);
label = new Label(0, r++, "");
writableSheet.addCell(label);}}
else if(inp[i][1].equalsIgnoreCase("v")){
if(inp[i][0].length()>5 && inp[i][0].matches("[0-9]+")){
for(j=1 ; j<46 ; j++){
label = new Label(0, r++, "Class"+Integer.toString(j));
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_DDLSearchType")).click();
new Select(driver.findElement(By.id("ContentPlaceHolder1_DDLSearchType"))).selectByVisibleText("Vienna Code");
driver.findElement(By.xpath("//option[@value='VC']")).click();
driver.findElement(By.xpath("//option[@value='VC']")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBVienna")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBVienna")).sendKeys(inp[i][0]);
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).sendKeys(Integer.toString(j));
driver.findElement(By.id("ContentPlaceHolder1_BtnSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
String results[] = t.split("\n");
if(!(results[0].contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
for(k=4; k<results.length-5; k++){
if(results[k].contains("International Non-Proprietary Names(INN)")){
label = new Label(0, r++, "International Non-Proprietary Names(INN)");
writableSheet.addCell(label);}
else if(results[k].contains("No Record found")){
label = new Label(0, r++, "No Record found");
writableSheet.addCell(label);}
else if(results[k].contains("Matching Trademark(s)")){
label = new Label(0, r++, "Matching Trademark(s)");
writableSheet.addCell(label);}
else if(results[k].contains("Wordmark:")){
for(l=0 ; l<5 ; l++){
label = new Label(l, r, results[k++]);
writableSheet.addCell(label);}
r++;}}
label = new Label(0, r++, "");
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_LnkNextSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
total += inc;
counter = (int)total;
b.doClick();}}
else{
label = new Label(0, r++, "Invalid Vienna Code");
writableSheet.addCell(label);
if(!inp[i][0].matches("[0-9]+")){
label = new Label(0, r++, "Enter only Numbers");
writableSheet.addCell(label);}
else{
Label label2 = new Label(0, r++, "Enter atleast 6 characters");
writableSheet.addCell(label2);}
label = new Label(0, r++, "");
writableSheet.addCell(label);
label = new Label(0, r++, "");
writableSheet.addCell(label);
total+=100/(float)((totalNoOfRows-5));}}
else{
if(inp[i][0].length()>2){
label = new Label(0, r++, "Start With");
writableSheet.addCell(label);
for(j=1 ; j<46 ; j++){
label = new Label(0, r++, "Class"+Integer.toString(j));
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).sendKeys(inp[i][0]);
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).sendKeys(Integer.toString(j));
driver.findElement(By.id("ContentPlaceHolder1_BtnSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
String results[] = t.split("\n");
if(!(results[0].contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
for(k=4; k<results.length-5; k++){
if(results[k].contains("International Non-Proprietary Names(INN)")){
label = new Label(0, r++, "International Non-Proprietary Names(INN)");
writableSheet.addCell(label);}
else if(results[k].contains("No Record found")){
label = new Label(0, r++, "No Record found");
writableSheet.addCell(label);}
else if(results[k].contains("Matching Trademark(s)")){
label = new Label(0, r++, "Matching Trademark(s)");
writableSheet.addCell(label);}
else if(results[k].contains("Wordmark:")){
for(l=0 ; l<5 ; l++){
label = new Label(l, r, results[k++]);
writableSheet.addCell(label);}
r++;}}
label = new Label(0, r++, "");
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_LnkNextSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
total += inc/3;
counter = (int)total;
b.doClick();}
label = new Label(0, r++, "Contains");
writableSheet.addCell(label);
for(j=1 ; j<46 ; j++){
label = new Label(0, r++, "Class"+Integer.toString(j));
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_DDLFilter")).click();
new Select(driver.findElement(By.id("ContentPlaceHolder1_DDLFilter"))).selectByVisibleText("Contains");
driver.findElement(By.xpath("//option[@value='1']")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).sendKeys(inp[i][0]);
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).sendKeys(Integer.toString(j));
driver.findElement(By.id("ContentPlaceHolder1_BtnSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
String results[] = t.split("\n");
if(!(results[0].contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
for(k=4; k<results.length-5; k++){
if(results[k].contains("International Non-Proprietary Names(INN)")){
label = new Label(0, r++, "International Non-Proprietary Names(INN)");
writableSheet.addCell(label);}
else if(results[k].contains("No Record found")){
label = new Label(0, r++, "No Record found");
writableSheet.addCell(label);}
else if(results[k].contains("Matching Trademark(s)")){
label = new Label(0, r++, "Matching Trademark(s)");
writableSheet.addCell(label);}
else if(results[k].contains("Wordmark:")){
for(l=0 ; l<5 ; l++){
label = new Label(l, r, results[k++]);
writableSheet.addCell(label);}
r++;}}
label = new Label(0, r++, "");
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_LnkNextSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
total += inc/3;
counter = (int)total;
b.doClick();}
label = new Label(0, r++, "Match With");
writableSheet.addCell(label);
for(j=1 ; j<46 ; j++){
label = new Label(0, r++, "Class"+Integer.toString(j));
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_DDLFilter")).click();
new Select(driver.findElement(By.id("ContentPlaceHolder1_DDLFilter"))).selectByVisibleText("Match With");
driver.findElement(By.xpath("//option[@value='2']")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBWordmark")).sendKeys(inp[i][0]);
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).click();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).clear();
driver.findElement(By.id("ContentPlaceHolder1_TBClass")).sendKeys(Integer.toString(j));
driver.findElement(By.id("ContentPlaceHolder1_BtnSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
String results[] = t.split("\n");
if(!(results[0].contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
for(k=4; k<results.length-5; k++){
if(results[k].contains("International Non-Proprietary Names(INN)")){
label = new Label(0, r++, "International Non-Proprietary Names(INN)");
writableSheet.addCell(label);}
else if(results[k].contains("No Record found")){
label = new Label(0, r++, "No Record found");
writableSheet.addCell(label);}
else if(results[k].contains("Matching Trademark(s)")){
label = new Label(0, r++, "Matching Trademark(s)");
writableSheet.addCell(label);}
else if(results[k].contains("Wordmark:")){
for(l=0 ; l<5 ; l++){
label = new Label(l, r, results[k++]);
writableSheet.addCell(label);}
r++;}}
label = new Label(0, r++, "");
writableSheet.addCell(label);
driver.findElement(By.id("ContentPlaceHolder1_LnkNextSearch")).click();
t = driver.findElement(By.tagName("body")).getText();
if(!(t.contains("Class Details |  Well Known Marks |  Prohibited Marks |  Vienna Code Classification |  International Non-Proprietary Names(INN) |  Help |"))) {
e.doClick();
driver.close();
label = new Label(0, r++, "Error");
writableSheet.addCell(label);
writableWorkbook.write();
writableWorkbook.close();
System.exit(1);}
total += inc/3;
counter = (int)total;
b.doClick();}}
else{
label = new Label(0, r++, "Invalid Wordmark");
writableSheet.addCell(label);
label = new Label(0, r++, "Enter atleast 3 characters");
writableSheet.addCell(label);
total+=100/(float)((totalNoOfRows-5));
label = new Label(0, r++, "");
writableSheet.addCell(label);
label = new Label(0, r++, "");
writableSheet.addCell(label);}}
label = new Label(0, r++, "");
writableSheet.addCell(label);}
if(counter<99){
counter = (int)total;
b.doClick();}
writableWorkbook.write();
writableWorkbook.close();
driver.close();
Desktop.getDesktop().open(new File(rfn));}
public void actionPerformed(ActionEvent evt) {
if (pbar.isCanceled()) {
pbar.close();
driver.close();
System.exit(1);}
else if (counter>=99)
pbar.close();
else if (evt.getActionCommand() == Actions.b.name()) {
pbar.setProgress(counter);
pbar.setNote(counter + "% complete");}
else if (evt.getActionCommand() == Actions.e.name()) {
pbar.setProgress(100);
pbar.setNote("Error.");}}}