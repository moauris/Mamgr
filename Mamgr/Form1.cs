using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;

namespace Mamgr
{
    public partial class Form1 : Form
    { 

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileInfo inputf = new FileInfo(textBox1.Text);

            XElement xmlRoot = VonExcel(inputf, textBox2, progressBar1);
            GC.Collect(); //With this, Excel will not lock up!
            
            xmlRoot.Save(@".\Result.xml");
            //Process.Start(@".\Result.xml");
            //EXCEL => xml conversion completed
            //Next step is go get xml to a structured database
            // ...
            XElement madb = new XElement("M-Agent_Database");

            XElement madb_server = new XElement("SERVER");

            //Need a method String Get_Value_By_Loc(XElement xmlRoot, String Loc){}
            //Total fields to fetch data:
            //madb_hostname_c1v
            //madb_hostname_c1p
            //madb_hostname_c1s
            //madb_hostname_c2v
            //madb_hostname_c2p
            //madb_hostname_c2s
            //madb_hostname_c3v
            //madb_hostname_c3p
            //madb_hostname_c3s
            int virtual_split = 1; //When Single, 1; when Cluster, 2;
            int virtual_index = 1; //When Single, 1; when cluster, v=0, p=1, s=2;
            XElement madb_hostname_c1v = 
                new XElement(Get_Value_By_Loc(xmlRoot, "H49"));

            Debug.Print("Does VIP Exist?");
            Debug.Print((madb_hostname_c1v == null).ToString());
            XElement madb_hostname_c1p =
                new XElement(Get_Value_By_Loc(xmlRoot, "H51"));
            madb_hostname_c1p.Add(new XElement("IP_Addr", Get_Value_By_Loc(xmlRoot, "H52")));
            madb_hostname_c1p.Add(new XElement("Maker", Get_Value_By_Loc(xmlRoot, "H53")));
            madb_hostname_c1p.Add(new XElement("Model", Get_Value_By_Loc(xmlRoot, "H54")));
            madb_hostname_c1p.Add(new XElement("CPU_Num", Get_Value_By_Loc(xmlRoot, "H55")));
            madb_hostname_c1p.Add(new XElement("CPU_Micro", Get_Value_By_Loc(xmlRoot, "H56")));
            //Need a method string Get_Value_By_Cbx(XElement xmlRoot, string() Loc){}
            // Loc is the string of *CheckBox* Locations
            string[] addr_OS = { "$H$57", "$H$58", "$H$59", "$J$57", "$J$58" };
            madb_hostname_c1p.Add(new XElement("OS", Get_Value_By_Cbx(xmlRoot, addr_OS)));
            madb_hostname_c1p.Add(new XElement("Version", Get_Value_By_Loc(xmlRoot, "H60")));
            madb_hostname_c1p.Add(new XElement("Bit", Get_Value_By_Loc(xmlRoot, "H61")));
            madb_hostname_c1p.Add(new XElement("Virtual_Split", Get_Value_By_Loc(xmlRoot, "H62")));
            madb_hostname_c1p.Add(new XElement("Virtual_Index", Get_Value_By_Loc(xmlRoot, "H63")));
            madb_hostname_c1p.Add(new XElement("Cluster_VIP", Get_Value_By_Loc(xmlRoot, "H49")));
            madb_hostname_c1p.Add(new XElement("Cluster_PRI", Get_Value_By_Loc(xmlRoot, "H51")));
            madb_hostname_c1p.Add(new XElement("Cluster_SEC", Get_Value_By_Loc(xmlRoot, "H64")));
            textBox2.AppendText(madb_hostname_c1p.ToString() + "\r\n");

        }

        string Get_Value_By_Cbx(XElement xmlRoot, string[] Loc)
        {// The address array is always 3 digit, such as H49
            string Result = "未入力";
            string col = "";
            string row = "";
            string value_addr = "";
            //Judge CheckBox Value, there should only be one check box checked

            foreach (string s in Loc)
            {//Parse String array, 0 is the character, 1 and 2 makes the row digit
                IEnumerable<XElement> EnumX =
                    from elx in xmlRoot.Elements("EXCEL_ELEMENT")
                    where elx.Element("Name").Value.Contains("Check Box") &&
                          elx.Element("Value").Value == "1" &&
                          elx.Element("Address").Value == s
                    select elx;
                Debug.Print(EnumX.Count().ToString() + " Objects Found");
                foreach (XElement xe in EnumX) Debug.Print(xe.Name.ToString());
                if (EnumX.Count() == 1) //Only continues when 1 element is selected
                {
                    col = ((char)((int)s[1] + 1)).ToString();
                    Debug.Print("Col is : " + col);
                    row = s[3].ToString() + s[4].ToString();
                    Debug.Print("Row is : " + row);
                    value_addr = col + row;
                    Debug.Print("Find Value by Loc at : " + value_addr);
                    Result = Get_Value_By_Loc(xmlRoot, value_addr);
                    Debug.Print(Result);
                }
                else
                {
                    if(EnumX.Elements().Count() != 0)
                        throw new Exception("More than One CheckBox is Checked");
                }


            }


            return Result;
        }

        private string Get_Value_By_Loc(XElement xmlRoot, string Loc)
        {
            string Result = "未入力";
            IEnumerable<XElement> EnumX =
                from elx in xmlRoot.Elements("EXCEL_ELEMENT")
                where elx.Element("Name").Value == Loc
                select elx;
            foreach (XElement xel in EnumX)
            {
                Result = xel.Element("Value").Value;
            }

            return Result;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Debug Disabled");
            string sdate = "　　　　　　2017年　　　　07月　　　28日";
            DateTime dt1 = Convert.ToDateTime(sdate);
            MessageBox.Show(sdate + "\r\n" + dt1.ToLongDateString()); //So powerful...
        }

        public static XElement VonExcel
            (FileInfo Input_File, TextBox WriteTo, ProgressBar update_pbar)
        {
            DateTime stamp_start = DateTime.Now;
            XElement xmlRoot = new XElement("EXCEL_DATA");
            XElement xmlGrid = new XElement("MA_GRID"); //Combine into on xml
            update_pbar.Maximum = 1000;
            update_pbar.Value = 10;
            FileInfo inputf = new FileInfo(Input_File.FullName);
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks xlWorkBooks = xlApp.Workbooks;
            Excel.Workbook xlWbk = xlWorkBooks.Open(inputf.FullName);
            Excel.Sheets xlWorksheets = xlWbk.Worksheets;
            Excel.Worksheet xlWks = xlWbk.ActiveSheet;
            Excel.Range xlRange = xlWks.UsedRange;
            update_pbar.Value += 10;

            ProgressContext<Excel.Range> IEmRange =
                new ProgressContext<Excel.Range>(
                    from Excel.Range rng in xlRange
                    where rng.Value != null
                    select rng);
            IEmRange.UpdateProgress += (sender0, e0) =>
            {
                update_pbar.Value += 1;
            };

            for (int r = 1; r < 100; r++)
            {
                XElement xmlRow = new XElement("Row");
                xmlRow.Add(new XElement("Name", r));
                xmlRow.Add(new XElement("Top", xlWks.Cells[r, 1].Top));

                xmlRow.Add(new XElement("Top_plus",
                    xlWks.Cells[r, 1].Top + xlWks.Cells[r, 1].Height));
                xmlRow.Add(new XElement("Value", 0));
                xmlGrid.Add(xmlRow);
            }
            for (int c = 1; c < 24; c++)
            {
                XElement xmlCol = new XElement("Col");
                xmlCol.Add(new XElement("Name", ((char)(64 + c)).ToString()));
                xmlCol.Add(new XElement("Left", xlWks.Cells[1, c].Left));
                xmlCol.Add(new XElement("Left_plus",
                    xlWks.Cells[1, c].Left + xlWks.Cells[1, c].Width));
                xmlCol.Add(new XElement("Value", 0));
                xmlGrid.Add(xmlCol);
            }
            
            foreach (Excel.Range rng in IEmRange)
            {
                XElement xmlE = new XElement("EXCEL_ELEMENT");
                //WriteTo.AppendText(rng.Address + 
                //" : " + rng.Address.GetType() +
                //" : " + rng.Value.ToString());
                xmlE.Add(new XElement("Name", 
                    rng.Address.Replace("$", "")));
                xmlE.Add(new XElement("Address", rng.Address));
                xmlE.Add(new XElement("Left", rng.Left));
                xmlE.Add(new XElement("Top", rng.Top));
                xmlE.Add(new XElement("Value", rng.Value.ToString()));
                xmlE.Add(new XElement("Option", 0));
                
                xmlRoot.Add(xmlE);
            }

            ProgressContext<Excel.Shape> IEmShapes =
                new ProgressContext<Excel.Shape>(
                    from Excel.Shape s in xlWks.Shapes
                    where s.Name.Contains("Check Box")
                    && (s.OLEFormat.Object.Value == 1)
                    select s);

            IEmShapes.UpdateProgress += (sender0, e0) =>
            {
                update_pbar.Value += 1;
            };

            foreach (Excel.Shape s in IEmShapes)
            {
                XElement xmlE = new XElement("EXCEL_ELEMENT");
                //WriteTo.AppendText(
                //s.OLEFormat.Object.Value.ToString());
                xmlE.Add(new XElement("Name",s.Name));
                xmlE.Add(new XElement("Address",
                    Extract_Address(s.Left, s.Top, s.Name, xmlGrid)));
                xmlE.Add(new XElement("Left", s.Left));
                xmlE.Add(new XElement("Top", s.Top));
                xmlE.Add(new XElement("Value", 
                    s.OLEFormat.Object.Value.ToString()));
                xmlRoot.Add(xmlE);
            }
            // This code fetches all top for rows, and left for columns

            xlWbk.Close();
            xlWorkBooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlWbk);
            Marshal.ReleaseComObject(xlWorksheets);
            Marshal.ReleaseComObject(xlWks);
            Marshal.ReleaseComObject(xlRange);

            DateTime stamp_End = DateTime.Now;
            TimeSpan run_duration = stamp_End - stamp_start;
            WriteTo.AppendText("Process Completed: ");
            WriteTo.AppendText(run_duration.ToString());
            WriteTo.AppendText("\r\n");

            return xmlRoot;
        }

        private static string Extract_Address(
            Double left, Double top, string name, XElement xmlGrid)
        {
            string res = "未入力";
            IEnumerable<XElement> IEmLocateR =
                from XElement xe in xmlGrid.Elements("Row")
                where (top > Convert.ToDouble(xe.Element("Top").Value)) &&
                      (top < Convert.ToDouble(xe.Element("Top_plus").Value))
                select xe;
            IEnumerable<XElement> IEmLocateC =
                from XElement xa in xmlGrid.Elements("Col")
                where (left > Convert.ToDouble(xa.Element("Left").Value)) &&
                      (left < Convert.ToDouble(xa.Element("Left_plus").Value))
                select xa;
            foreach (XElement xe in IEmLocateR)
            {
                foreach (XElement xa in IEmLocateC)
                {  
                    res = "$" +
                    xa.Element("Name").Value + 
                    "$" +
                    xe.Element("Name").Value;
                }
            }
            return res;
        }
        
        private static void Check_Status(Task Target_Task)
        {
            while (!Target_Task.IsCompleted)
            {
                Thread.Sleep(TimeSpan.FromSeconds(3));
            }
        }

        private static void releasedObject(ref Excel.Application obj)
        {
            if(obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
            obj = null;
        }
    }
    

    public class ProgressArgs : EventArgs
    {
        public ProgressArgs(int count)
        {
            this.Count = count;
        }
        public int Count { get; private set; }
    }
    public class ProgressContext<T> : IEnumerable<T>
    {
    private IEnumerable<T> source;

    public ProgressContext(IEnumerable<T> source)
    {
        this.source = source;
    }

    public event EventHandler<ProgressArgs> UpdateProgress;

    protected virtual void OnUpdateProgress(int count)
    {
            this.UpdateProgress?.Invoke(this, new ProgressArgs(count));
        }

    public IEnumerator<T> GetEnumerator()
    {
        int count = 0;
        foreach (var item in source)
        {
            // The yield holds execution until the next iteration,
            // so trigger the update event first.
            OnUpdateProgress(++count);
            yield return item;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }    }
}
