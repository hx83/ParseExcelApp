using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

using System.Text;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;


namespace ParseExcelApp
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string cachePath;

        private string filePath = "";
        private string savePath = "";
        //
        private string saveFileName;
        private string excelFilePath;
        //

        private string className;
        //

        private List<XmlNodeInfo> xmlInfoList;
        private List<int> columnList;
        public MainWindow()
        {
            InitializeComponent();

            columnList = new List<int>();
            xmlInfoList = new List<XmlNodeInfo>();


            cachePath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\data.txt";

            CheckCache();
        }


        private void SelectExcelFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            //dialog.Filter = "所有文件(*.*)|*.*";
            dialog.DefaultExt = ".xml";
            dialog.Filter = "xml file|*.xml";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                //Console.Write(file);
                importFileTxt.Text = file;
                filePath = file;
            }
        }

        private void SelectSavePath(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.Description = "请选择一个路径";
            if (FBD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                exportPathTxt.Text = FBD.SelectedPath;
                savePath = FBD.SelectedPath + "\\";
            }
        }

        /// <summary>
        /// 解析
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Parse(object sender, RoutedEventArgs e)
        {
            if (filePath == "" || savePath == "")
            {
                clsTxt.Text = "请选择文件和输出路径！";
                return;
            }

            xmlInfoList.Clear();
            columnList.Clear();

            FileStream fs = File.OpenRead(filePath);
            
            XmlDocument xml = new XmlDocument();
            xml.Load(fs);
            
            //Console.Write(xml.InnerXml.ToString());
            XmlNode xn = xml.SelectSingleNode("config/table");
            XmlAttributeCollection xc = xn.Attributes;
            saveFileName = xc.GetNamedItem("name").Value + ".bytes";

            className = xc.GetNamedItem("name").Value + "Def";

            excelFilePath = filePath.Substring(0, filePath.LastIndexOf('\\')+1) + xc.GetNamedItem("ExcelFile").Value;

            xn = xml.SelectSingleNode("config/table/fields");
            for (int i = 0; i < xn.ChildNodes.Count; i++)
            {
                XmlNode node = xn.ChildNodes.Item(i);

                XmlNodeInfo info = new XmlNodeInfo();
                info.codeName = node.Attributes.GetNamedItem("codename").Value;
                info.excelName = node.Attributes.GetNamedItem("name").Value;
                info.type = node.Attributes.GetNamedItem("type").Value;

                XmlNode attNode = node.Attributes.GetNamedItem("size");
                if (attNode != null)
                    info.size = Convert.ToInt32(attNode.Value.ToString());

                xmlInfoList.Add(info);
            }

            ReadExcel();
        }


        private void ReadExcel()
        {   
            Excel.Application excel = new Excel.Application();
            if (File.Exists(excelFilePath) == false)
            {
                clsTxt.Text = "excel文件不存在！";
                return;
            }

            Excel.Workbook wb = excel.Workbooks.Open(excelFilePath);
            //取得第一个工作薄  
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            //取得总记录行数    (包括标题列)  
            int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
            int colint = ws.UsedRange.Cells.Columns.Count; //得到列数

            for (int i = 0; i < xmlInfoList.Count; i++)
            {
                XmlNodeInfo info = xmlInfoList[i];
                for (int j = 1; j <= colint; j++)
                {
                    string str = (string)ws.Cells[1, j].Text;

                    if (str == info.excelName)
                    {
                        columnList.Add(j);
                        break;
                    }
                }
            }
            //
            ByteArray byteArr = new ByteArray();
            //
            int dataLen = rowsint - 1;

            byteArr.WriteSignedInt(dataLen);

            for (int i = 2; i <= rowsint; i++)
            {
                for (int j = 0; j < xmlInfoList.Count; j++)
                {
                    XmlNodeInfo info = xmlInfoList[j];
                    if(info.type == "int")
                    {
                        string str = (string)ws.Cells[i, columnList[j]].Text;
                        if (str == "")
                            str = "0";
                        int intvalue = Convert.ToInt32(str);
                        Console.WriteLine(intvalue);
                        byteArr.WriteSignedInt(intvalue);
                    }
                    else if(info.type == "float")
                    {
                        string str = (string)ws.Cells[i, columnList[j]].Text;
                        if (str == "")
                            str = "0";
                        float floatvalue = float.Parse(str);
                        Console.WriteLine(floatvalue);
                        byteArr.WriteBytes(BitConverter.GetBytes(floatvalue), 4);
                    }
                    else if(info.type == "string")
                    {
                        string strvalue = (string)ws.Cells[i, columnList[j]].Text;
                        Console.WriteLine(strvalue);
                        byteArr.WriteUTFBytes(strvalue, info.size);
                    }
                }
            }

            byteArr.Compress();

            FileStream fs = new FileStream(savePath + saveFileName, FileMode.Create);
            fs.Write(byteArr.Bytes,0,byteArr.Length);
            fs.Close();

            //
            excel.Quit();
            excel = null;
            //
            //
            CreateCode();

            //
            SaveSelectPath();
        }

        private string classStr = "using System;\r\nusing System.Collections.Generic;\r\n\r\n\r\npublic class $classname\r\n{\r\n$prop\r\n\r\n\tpublic void Read(ByteArray buf)\r\n\t{\r\n$content\r\n\t}\r\n    \r\n}\r\n";
        private string classTemp;
        private void CreateCode()
        {
            classTemp = "";
            //StreamReader sr = File.OpenText("Template\\DefTemplate.txt");
            //classTemp = sr.ReadToEnd();
            //sr.Close();

            classTemp = classStr.Replace("$classname", className);
            //
            string propStr = "";
            string contentStr = "";

            for (int i = 0; i < xmlInfoList.Count; i++)
            {
                XmlNodeInfo info = xmlInfoList[i];
                propStr += "\t//" + info.excelName+"\r";
                propStr += "\tpublic " + info.type + " " + info.codeName + ";\r";

                
                string readStr = "";
                if (info.type == "string")
                    readStr = "buf.ReadUTFByte(" + info.size + ");";
                else if (info.type == "int")
                    readStr = "buf.ReadSignedInt();";
                else if (info.type == "float")
                    readStr = "buf.ReadFloat();" ;

                contentStr += "\t\t" + info.codeName + " = " + readStr + "\r";
            }
            classTemp = classTemp.Replace("$prop", propStr);
            //
            classTemp = classTemp.Replace("$content", contentStr);

            Console.WriteLine(classTemp);

            clsTxt.Text = classTemp;

            //ByteArray by = new ByteArray();
            //by.WriteUTFBytes(classTemp, classTemp.Length);
            //byte[] strbytes = classTemp.to;
            //BitConverter.GetBytes();

            byte[] strbytes = Encoding.Default.GetBytes(classTemp);

            FileStream fs = new FileStream(savePath + className + ".cs", FileMode.Create);
            fs.Write(strbytes, 0, strbytes.Length);
            fs.Close();
        }

        private void SaveSelectPath()
        {
            FileStream fs = File.Create(cachePath);
            string str = filePath + "|" + savePath;

            byte[] strbytes = Encoding.Default.GetBytes(str);
            fs.Write(strbytes, 0, strbytes.Length);
            fs.Close();
        }

        private void CheckCache()
        {
            if (File.Exists(cachePath) == false)
                return;

            StreamReader sr = File.OpenText(cachePath);
            string str = sr.ReadToEnd();

            string[] arr = str.Split(new char[] { '|' });
            if(arr.Length == 2)
            {
                filePath = arr[0];
                savePath = arr[1];

                importFileTxt.Text = filePath;
                exportPathTxt.Text = savePath;
            }
            sr.Close();
            
        }
    }
}


public class XmlNodeInfo
{
    public string codeName;
    public string excelName;
    private string _type;
    public int size;


    public string type
    {
        set
        {
            _type = value;

            if(_type == "string")
            {
                size = 16;
            }
        }
        get
        {
            return _type;
        }
    }
}