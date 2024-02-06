using Microsoft.Office.Interop.Excel;
using MST;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V102.CSS;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace MST
{
    public partial class Form1 : Form
    {
        public static IWebDriver driver1;
        public static IList<IWebElement> listcty;
        public static IList<IWebElement> listnghe, list1, list2;
        public Application excel = new Application();
        public Workbook wb;
        public Worksheet ws, ws1;
        public Range cell1, cell2, cell3, cell4, cell5, cell6;
        public static long fullRow, lastRow, lastRowsheet2;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public static string nganhnghe, ChromePath, pathString, mst, data, cty;
        public static string AppPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public static string link1 = "https://www.trangvangvietnam.com";
        public static string link2 = "https://masothue.com/";
        private void btnConfirm_Click(object sender, EventArgs e)
        {
            driver1 = new ChromeDriver(ChromePath);
            driver1.Navigate().GoToUrl(link2);
            driver1.Manage().Window.Maximize();
            excel.DisplayAlerts = false;
            wb = excel.Workbooks.Open(pathString + "\\FileMST.xlsx", 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = wb.Worksheets[1];
            ws1 = wb.Worksheets[2];
            fullRow = ws.Rows.Count;
            lastRow = ws.Cells[fullRow, 2].End(XlDirection.xlUp).Row;
            for (int i = 2; i <= lastRow; i++)
            {
                lastRowsheet2 = ws1.Cells[fullRow, 2].End(XlDirection.xlUp).Row;
                nganhnghe = ws.Cells[i, 1].Value;
                mst = ws.Cells[i, 2].Value;
                cty = ws.Cells[i, 3].Value;
                try
                {
                    //
                    driver1.FindElement(By.XPath("//*[@id='search']")).Clear();
                    driver1.FindElement(By.XPath("//*[@id='search']")).SendKeys(mst + OpenQA.Selenium.Keys.Enter);

                    //Thread.Sleep(1000);
                    //đoạn này tương tự việc xác định vị trí lưu dữ liệu
                    cell1 = ws1.Cells[lastRowsheet2 + 1, 1];
                    cell1 = ws1.Cells[lastRowsheet2 + 1, 2]; //mst
                    cell2 = ws1.Cells[lastRowsheet2 + 1, 3]; //cty
                    cell3 = ws1.Cells[lastRowsheet2 + 1, 4]; //daidien
                    cell4 = ws1.Cells[lastRowsheet2 + 1, 5]; //trangthai
                    cell5 = ws1.Cells[lastRowsheet2 + 1, 6]; //tel
                    cell6 = ws1.Cells[lastRowsheet2 + 1, 1];

                    cell1.Value = mst;
                    cell2.Value = cty;
                    cell6.Value = nganhnghe;
                    Thread.Sleep(1000);
                    list1 = driver1.FindElement(By.TagName("table")).FindElements(By.TagName("tr"));
                    list2 = driver1.FindElements(By.ClassName("table-taxinfo"))[0].FindElements(By.TagName("span"));

                    for (int j = 1; j < list1.Count -1; j++)
                    {
                        //chỗ này kiểm tra toàn bộ element, nếu element nào có thông tin là các case bên dưới thì giá trị của nó chính là thông tin mình cần
                       
                            data = driver1.FindElement(By.XPath("//*[@id='main']/section[1]/div/table[1]/tbody/tr[" + j + "]/td[1]")).Text.Trim();
                            switch (data)
                            {
                                case "Người đại diện":
                                    cell3.Value = list2[j].Text;
                                    break;
                                case "Điện thoại":
                                    cell5.Value = list2[j].Text;
                                    break;
                                case "Tình trạng":
                                    cell4.Value = driver1.FindElement(By.XPath("//*[@id='main']/section[1]/div/table[1]/tbody/tr[" + j + "]/td[2]/a")).Text;
                                    break;
                                default:
                                    break;
                            }
                        

                    }
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    driver1.Navigate().Refresh();
                    cell3.Value = "";
                    cell4.Value = "Không có kết quả";
                    cell5.Value = "";
                    continue;
                }
                
            }
            wb.SaveAs(pathString + "\\FileMST_final.xlsx", AccessMode: XlSaveAsAccessMode.xlNoChange);
            wb.Close();
            excel.Quit();
            driver1.Close();
            driver1.Quit();
        }

        private void btnGet_Click(object sender, EventArgs e)
        {
            driver1 = new ChromeDriver(ChromePath);
            driver1.Navigate().GoToUrl(link1);
            driver1.Manage().Window.Maximize();
            driver1.FindElement(By.XPath("//*[@id='myBtns']")).Click();
            listnghe = driver1.FindElement(By.Id("box_niengiamnganh")).FindElements(By.CssSelector("#niengiam25 > div.cell_niengiam_txt > h2 > a"));
            excel.DisplayAlerts = false;
            wb = excel.Workbooks.Open(pathString + "\\FileTemplate.xlsx", 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = wb.Worksheets[1];
            fullRow = ws.Rows.Count;

            //25 nganh hien thi
            for (int i = 12; i < 13; i++)
            {
                nganhnghe = listnghe[i].Text;
                listnghe[i].Click();
                //gọi vào hàm ngành tiêu điểm
                NganhTieuDiem();
                driver1.Navigate().GoToUrl(link1);
                listnghe = driver1.FindElement(By.Id("box_niengiamnganh")).FindElements(By.CssSelector("#niengiam25 > div.cell_niengiam_txt > h2 > a"));
            }
            /*
            driver1.FindElement(By.XPath("//*[@id='myBtns']")).Click();
            listnghe = driver1.FindElement(By.Id("main_box_niengiamnganh")).FindElements(By.CssSelector("#niengiam > div.cell_niengiam_txt > h2 > a"));
            
            //9 nganh thuoc more
            for (int i = 0; i < listnghe.Count; i++)
            {
                nganhnghe = listnghe[i].Text;
                listnghe[i].Click();
                NganhTieuDiem();
                driver1.Navigate().GoToUrl(link1);
                driver1.FindElement(By.XPath("//*[@id='myBtns']")).Click();
                listnghe = driver1.FindElement(By.Id("box_niengiamnganh")).FindElements(By.CssSelector("#niengiam > div.cell_niengiam_txt > h2 > a"));
            }*/
            wb.SaveAs(pathString + "\\FileMST.xlsx", AccessMode: XlSaveAsAccessMode.xlNoChange);
            wb.Close();
            excel.Quit();


            driver1.Close();
            driver1.Quit();
        }



        public Form1()
        {
            InitializeComponent();
            //kiem tra folder Data có ton tại trong đường dẫn của chương trình không
            //nếu chưa có thì tao mowí folder
            pathString = AppPath + "\\Data";
            if (!System.IO.File.Exists(pathString))
            {
                System.IO.Directory.CreateDirectory(pathString);
            }
            //đọc file ini để lấy thông tin đường dẫn của chrome driver
            var MyIni = new IniFile(AppPath + "\\Config.ini");
            ChromePath = MyIni.Read("ChromeDriver", "MyApplication");
            label2.Text = "Ver " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

        }


        private void NganhTieuDiem()
        {

            //vì lười không xác định số ngành tiêu điểm nên mặc định cho 100 ngành
            for (int j = 1; j < 2; j++)
            {
                // khi không tìm thấy ngành tiêu điểm nó sẽ tự thoát vòng for
                if (driver1.FindElements(By.XPath("//*[@id='niengiampages_content']/div[1]/div[1]/div[2]/div[" + j + "]/div[2]/a")).Count ==0)
                {
                    break; // thoat loop khi khong ton tai nganh tieu diem nao nua
                }
                driver1.FindElement(By.XPath("//*[@id='niengiampages_content']/div[1]/div[1]/div[2]/div[" + j + "]/div[2]/a")).Click();
                //gọi đến hàm list mã số thuế để lấy thông tin mst cuart từng công ty
                listmst();

                //hàm back này có nhiệm vụ quay lại trang truowcs đó 
                driver1.Navigate().Back();
            }
        }
        private void listmst()
        {
            while(true)
            {
                //đoạn này là chỗ kiểm tra trang cuối cùng lúc đầu đã nói để dừng
                if (driver1.FindElement(By.XPath("//*[@id='paging']")).FindElements(By.ClassName("page_active")).Count == 2)
                {
                    break; //cong ty cuoi cung
                }

                //danh sách công ty được hiển thị trong 1 page: listcty
                listcty = driver1.FindElements(By.ClassName("company_name"));
                for (int k = 2; k <= listcty.Count + 1; k++)
                {//vào từng công ty để lấy thông tin
                    driver1.FindElement(By.CssSelector("#listingsearch > div:nth-child("+ k +") > div.listings_top > div.noidungchinh > h2 > a")).Click();
                    //đoạn này xác định vị trí cell để lưu thông tin
                    long lastRow = ws.Cells[fullRow, 2].End(XlDirection.xlUp).Row;
                    cell1 = ws.Cells[lastRow + 1, 1]; //ngành nghề
                    cell2 = ws.Cells[lastRow + 1, 2]; //mã số thuế
                    cell3 = ws.Cells[lastRow + 1, 3]; //tên công ty
                    try
                    {
                        //3 truong hop vi tri cua MST
                        int cnt = driver1.FindElements(By.ClassName("thongtinchitiet")).Count - 1;

                        for (int m = 0; m < driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElements(By.ClassName("hosocongty_tite_text")).Count; m++)
                        {
                            switch (driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElements(By.ClassName("hosocongty_tite_text"))[m].Text)
                            {
                                case "Mã số thuế:":
                                    if (driver1.FindElements(By.CssSelector("#listing_detail_left > div:nth-child(3) > div:nth-child(6) > div.hosocongty_text")).Count == 1)
                                        cell2.Value = driver1.FindElement(By.CssSelector("#listing_detail_left > div:nth-child(3) > div:nth-child(6) > div.hosocongty_text")).Text;
                                    else if (driver1.FindElements(By.CssSelector("#listing_detail_left > div:nth-child(4) > div:nth-child(6) > div.hosocongty_text")).Count == 1)
                                        cell2.Value = driver1.FindElement(By.CssSelector("#listing_detail_left > div:nth-child(4) > div:nth-child(6) > div.hosocongty_text")).Text;
                                    else
                                        cell2.Value = driver1.FindElement(By.CssSelector("#listing_detail_left > div:nth-child(5) > div:nth-child(6) > div.hosocongty_text")).Text;


                                    cell1.Value = nganhnghe;
                                    cell3.Value = driver1.FindElement(By.XPath("//*[@id='listing_basic_info']/div[1]/h1")).Text;

                                    break;
                                default:
                                    break;
                            }
                        }

                            //int cnt = driver1.FindElements(By.ClassName("thongtinchitiet")).Count - 1;

                            //for (int m = 0; m < driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElements(By.ClassName("hosocongty_tite_text")).Count; m++)
                            //{
                            //    switch (driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElements(By.ClassName("hosocongty_tite_text"))[m].Text)
                            //    {
                            //        case "Mã số thuế:":

                            //            if (driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElements(By.CssSelector("#listing_detail_left > div:nth-child("+(m+1)+") > div:nth-child(6) > div.hosocongty_text")).Count == 1)
                            //                cell2.Value = driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElement(By.CssSelector("#listing_detail_left > div:nth-child(" + (m + 1) + ") > div:nth-child(6) > div.hosocongty_text")).Text;
                            //            else
                            //                cell2.Value = driver1.FindElements(By.ClassName("thongtinchitiet"))[cnt].FindElement(By.CssSelector("#listing_detail_left > div:nth-child("+m+") > div:nth-child(6) > div.hosocongty_text")).Text;
                            //            cell1.Value = nganhnghe;
                            //            cell3.Value = driver1.FindElement(By.XPath("//*[@id='listing_basic_info']/div[1]/h1")).Text;
                            //            break;
                            //        default:
                            //            break;
                            //    }
                            //}

                        }


                        int cnt = findEleentsBy("ClassName", "ThongTinChiTiet")
                    catch (Exception)
                    {
                        //neu khong tim duoc MST no se chay vao day de chon cty khac
                        driver1.Navigate().Back();
                        continue;
                    }
                    
                    driver1.Navigate().Back();
                    
                }
                //xong 1 page sẽ click vào button Tiếp để sang trang kế bên
                driver1.FindElement(By.Id("paging")).FindElement(By.LinkText("Tiếp")).Click();
               
            }
        }
    }
}

