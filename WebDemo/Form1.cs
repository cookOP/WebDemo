using Demo;
using HtmlAgilityPack;
using HttpCode.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net;
using System.Threading;
using System.Web;
using System.Windows.Forms;

namespace WebDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {          
            string txtURL = textBox1.Text.Trim();
            if (string.IsNullOrEmpty(txtURL))
                MessageBox.Show("请填写地址");
            else
            {
                Enabled = false;
                GetHtmlContent(txtURL);                
                button1.BackgroundImage = null;
            }
            Enabled = true;
        }
        /// <summary>
        /// 获取HTML内容
        /// </summary>
        /// <param name="txtURL">url地址</param>
        public void GetHtmlContent(string txtURL)
        {
            HttpHelpers httpHelpers = new HttpHelpers();
            HttpItems items = new HttpItems();          
            var u = "";
            var page = 0;
            int totalPage = page == 0 ? 1 : page;//总页数
            string urlParam = txtURL; //地址栏参数             
            string urlParam2 = "";
            string excelFileNmae = "",school="";
            
            DataTable tblDatas = new DataTable("Datas");
            tblDatas.Columns.Add("Name", Type.GetType("System.String"));
            tblDatas.Columns.Add("ShoolName", Type.GetType("System.String"));
            tblDatas.Columns.Add("ProjectNname", Type.GetType("System.String"));
            tblDatas.Columns.Add("GPA", Type.GetType("System.String"));
            tblDatas.Rows.Add(new object[] { null, null, null, null });
            List<string> listUrl = new List<string>();
            try
            {
                for (int i = 1; i <= totalPage; i++)
                {
                    if (!string.IsNullOrEmpty(urlParam2) && totalPage > 1 && i > 1)
                    {
                        urlParam = urlParam2 + "&page=" + i;
                    }
                    items.Url = urlParam;//请求地址
                    items.Method = "GET";
                    items.ContentType = "text/html; charset=utf-8";
                    HttpResults hr = httpHelpers.GetHtml(items);
                  
                    //解析数据
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    //加载html
                    doc.LoadHtml(hr.Html);
                    var eFileName = doc.DocumentNode.SelectNodes("/html/body/div[2]/div[5]/ul/li[@class='newli']");
                    if (eFileName.Count > 0)
                    {
                        string efn = "";
                        foreach (var item in eFileName)
                        {
                            var name = item.SelectSingleNode("span").InnerText;
                            efn += "-" + name;
                        }
                        excelFileNmae = DateTime.Now.ToString("yyyyMMdd") + efn;

                    }
                    //获取 class=post_item_body 的div列表
                    HtmlNodeCollection itemNodes = doc.DocumentNode.SelectNodes("//div[5]/div/div[1]/div[1]/ul/li");
                    foreach (var item in itemNodes)
                    {
                        var name = HtmlHelper.NoHTML(item.SelectSingleNode("div[@class='case-student-name']").InnerText);
                        var studentName = HtmlHelper.NoHTML(item.SelectSingleNode("div[@class='student-name']/div").InnerText);
                        var projectNname = HtmlHelper.NoHTML(item.SelectSingleNode("div[@class='student-name']/p[@class='project-name']/span").InnerText);
                        var GPA = item.SelectSingleNode("div[@class='student-name']/p[2]/span").InnerText;
                        var url2 = item.SelectSingleNode("a").GetAttributeValue("href", "");
                        listUrl.Add(url2);
                        tblDatas.Rows.Add(new object[] { name, studentName, projectNname, GPA });
                    }
                    if (totalPage > 1)
                        Thread.Sleep(3000); //每抓取一页数据 暂停三秒
                }
                Dictionary<string, DataTable> dic = new Dictionary<string, DataTable>();
                dic.Add("TableWC", tblDatas);
                ExcelHelper.Export(dic, GetTuple(excelFileNmae).Item1, out Tuple<string, string> t, excelFileNmae);                
                DataTable tblDatas2 = new DataTable("Datas2");
                tblDatas2.Columns.Add("标题", Type.GetType("System.String"));
                tblDatas2.Columns.Add("录取学校数量", Type.GetType("System.String"));
                tblDatas2.Columns.Add("录取学校基本信息", Type.GetType("System.String"));
                tblDatas2.Columns.Add("描述", Type.GetType("System.String"));
                tblDatas2.Columns.Add("给类似情况学生的建议", Type.GetType("System.String"));
                tblDatas2.Rows.Add(new object[] { null, null, null, null, null });
                foreach (var item in listUrl)
                {
                    items.Url = "http://www.liuxue.com" + item;//请求地址
                    items.Method = "GET";
                    items.ContentType = "text/html; charset=utf-8";
                    HttpResults hr = httpHelpers.GetHtml(items);
                    //解析数据
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    //加载html
                    doc.LoadHtml(hr.Html);
                    //获取 class=post_item_body 的div列表
                    HtmlNodeCollection itemNodes = doc.DocumentNode.SelectNodes("//div[3]/div[@class='main-body']");
                    if (itemNodes.Count > 0)
                    {
                        foreach (var itemDetail in itemNodes)
                        {
                            var caseTile = itemDetail.SelectSingleNode("div[@class='case-title']").InnerText;
                            var caseStudent = itemDetail.SelectSingleNode("div[@class='case-student']/p").InnerText;
                            //case-student-offers
                            var schoolName = "";
                            var logoList = itemDetail.SelectNodes("div[@class='case-student']/div[@class='case-student-offers']/ul/li");
                            foreach (var item3 in logoList)
                            {
                                var imgLogo = item3.SelectSingleNode("div[@class='img']/img").GetAttributeValue("src", "");
                                var enoroll = item3.SelectSingleNode("div[@class='content-list']").InnerText;
                                var imgLogoName = item3.SelectSingleNode("div[@class='content-list']/p[1]").InnerText;

                                var imgEnorllo = item3.SelectSingleNode("div[@class='img-container offer-img-wrap']/img").GetAttributeValue("src", "");
                                var fileName = imgLogoName + "logo.jpg";
                                var fileName2 = imgLogoName + ".png";
                                schoolName += enoroll;
                                SaveImage(imgLogo, fileName);
                                SaveImage(imgEnorllo, fileName2);
                                Thread.Sleep(1500);//暂停1.5秒获取图片
                            }
                            var caseApplyDesc = itemDetail.SelectSingleNode("div[@class='case-apply-analyse']").InnerText;
                            var caseSuggest = "";
                            try
                            {
                                caseSuggest = itemDetail.SelectSingleNode("div[@class='case-suggest']").InnerText;
                            }
                            catch (Exception ex)
                            {

                               
                            }
                            tblDatas2.Rows.Add(new object[] { caseTile, caseStudent, schoolName, caseApplyDesc, caseSuggest });
                        }
                    }
                    if (listUrl.Count > 1)
                       Thread.Sleep(1500); //暂停1.5秒获详情界面数据
                }
                Dictionary<string, DataTable> dic2 = new Dictionary<string, DataTable>();
                dic2.Add("TableWC2", tblDatas2);
                ExcelHelper.Export(dic2, GetTuple(excelFileNmae + "Detail").Item1, out Tuple<string, string> t2, excelFileNmae + "Details");


            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取数据失败,错误信息：{ex.Message}");
                return;
            }
            MessageBox.Show("获取数据成功!");
        }
        /// <summary>
        /// 公用路径
        /// </summary>
        /// <returns></returns>
        private Tuple<string, string> GetTuple(string excelFileNmae)
        {
            string sWebRootFolder =AppDomain.CurrentDomain.BaseDirectory + "\\Upload";
            string sFileName = $"{excelFileNmae}.xlsx";
            if (!File.Exists(sWebRootFolder))
                Directory.CreateDirectory(sWebRootFolder);
            if (File.Exists(sWebRootFolder + "\\" + sFileName))
            {
                string delUrl = sWebRootFolder + "\\" + sFileName;
                File.Delete(delUrl);
            }
            return Tuple.Create(sWebRootFolder, sFileName);
        }

        public bool SaveImage(string url, string fileName)
        {
            try
            {
                string wwwroot = Directory.GetCurrentDirectory();
                if (!File.Exists(wwwroot + "\\Upload\\Images"))
                {
                    Directory.CreateDirectory(wwwroot + "\\Upload\\Images");
                }
                if (File.Exists(wwwroot + "\\Upload\\Images\\" + fileName))
                {
                    return true;
                }
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(new Uri(url), wwwroot + "\\Upload\\Images\\" + fileName);

                    //OR 
                    client.Dispose();
                    // client.DownloadFileAsync(new Uri(url), @"c:\temp\image35.png");
                }
            }
            catch (Exception ex)
            {
                var msg = ex.Message;
                return false;

            }
            return true;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //复制粘贴url地址的时候解码
            var txt = HttpUtility.UrlDecode(textBox1.Text.ToString(), System.Text.Encoding.UTF8);
            if (!string.IsNullOrEmpty(txt))
                textBox1.Text = txt;
            
            
        }
    }
}
