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
        delegate void Add_Ma_Element(XElement XE_Add, string El_Name, string El_Value);
        delegate void Add_Ma_Element_cbx(XElement XE_Add, string El_Name, string[] El_Value);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Utl u = new Utl();
            u.dbg("Button1 Clicked, Procedure Started");
            FileInfo inputf = new FileInfo(textBox1.Text);
            u.dbg("Using File info: " + inputf.FullName);
            u.dbg("Executing procedure VonExcel");
            XElement xmlRoot = VonExcel(inputf, textBox2, progressBar1, u);
            GC.Collect(); //With this, Excel will not lock up!
            u.dbg("VonExcel Garbage Collection");
            xmlRoot.Save(@".\Result.xml");
            u.dbg("xml Saved.");

            //Process.Start(@".\Result.xml");
            //EXCEL => xml conversion completed
            //Next step is go get xml to a structured database
            // ...
            u.dbg("Creating madb Elements");
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
            Add_Ma_Element maer = //Delegate to add 2 strings to an XElement
                (XElement xe, string s1, string s2)
                => xe.Add(new XElement(s1, Get_Value_By_Loc(xmlRoot, s2)));
            Add_Ma_Element_cbx maex = //Variation: Find CheckBox
                (XElement xe, string s1, string[] s2)
                => xe.Add(new XElement(s1, Get_Value_By_Cbx(xmlRoot, s2)));

            XElement madb_hostname_c1v = 
                new XElement(Get_Value_By_Loc(xmlRoot, "H49"));
            maer(madb_hostname_c1v, "IP_Addr", "H50");
            maer(madb_hostname_c1v, "Cluster_VIP", "H49");
            maer(madb_hostname_c1v, "Cluster_PRI", "H51");
            maer(madb_hostname_c1v, "Cluster_SEC", "H64");
            XElement madb_hostname_c1p =
                new XElement(Get_Value_By_Loc(xmlRoot, "H51"));
            maer(madb_hostname_c1p, "IP_Addr", "H52");
            maer(madb_hostname_c1p, "Maker", "H53");
            maer(madb_hostname_c1p, "Model", "H54");
            maer(madb_hostname_c1p, "CPU_Num", "H55");
            maer(madb_hostname_c1p, "CPU_Micro", "H56");
            //Need a method string Get_Value_By_Cbx(XElement xmlRoot, string() Loc){}
            // Loc is the string of *CheckBox* Locations
            string[] addr_OS = { "$H$57", "$H$58", "$H$59", "$J$57", "$J$58" };
            maex(madb_hostname_c1p, "OS", addr_OS);
            maer(madb_hostname_c1p, "Version", "H60");
            maer(madb_hostname_c1p, "Bit", "H61");
            maer(madb_hostname_c1p, "Virtual_Split", "H62");
            maer(madb_hostname_c1p, "Virtual_Index", "H63");
            maer(madb_hostname_c1p, "Cluster_VIP", "H49");
            maer(madb_hostname_c1p, "Cluster_PRI", "H51");
            maer(madb_hostname_c1p, "Cluster_SEC", "H64");
            XElement madb_hostname_c1s =
                new XElement(Get_Value_By_Loc(xmlRoot, "H64"));
            maer(madb_hostname_c1s, "IP_Addr", "H65");
            maer(madb_hostname_c1s, "Maker", "H66");
            maer(madb_hostname_c1s, "Model", "H67");
            maer(madb_hostname_c1s, "CPU_Num", "H68");
            maer(madb_hostname_c1s, "CPU_Micro", "H69");
            string[] addr_OS2 = { "$H$70", "$H$71", "$H$72", "$J$70", "$J$71" };
            maex(madb_hostname_c1s, "OS", addr_OS2);
            maer(madb_hostname_c1s, "Version", "H73");
            maer(madb_hostname_c1s, "Bit", "H74");
            maer(madb_hostname_c1s, "Virtual_Split", "H75");
            maer(madb_hostname_c1s, "Virtual_Index", "H76");
            maer(madb_hostname_c1s, "Cluster_VIP", "H49");
            maer(madb_hostname_c1s, "Cluster_PRI", "H51");
            maer(madb_hostname_c1s, "Cluster_SEC", "H64");

            textBox2.AppendText(madb_hostname_c1v.ToString() + "\r\n");
            textBox2.AppendText(madb_hostname_c1p.ToString() + "\r\n");
            textBox2.AppendText(madb_hostname_c1s.ToString() + "\r\n");
            u.dbg("MA information get.");
            u.dbg_done();

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
                if (EnumX.Count() == 1) //Only continues when 1 element is selected
                {
                    col = ((char)((int)s[1] + 1)).ToString();
                    row = s[3].ToString() + s[4].ToString();
                    value_addr = col + row;
                    Result = Get_Value_By_Loc(xmlRoot, value_addr);
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
            (FileInfo Input_File, TextBox WriteTo, ProgressBar update_pbar, Utl u)
        {
            u.dbg("VonExcel: Started");
            XElement xmlRoot = new XElement("EXCEL_DATA");
            XElement xmlGrid = new XElement("MA_GRID"); //Combine into on xml
            u.dbg("VonExcel: Root XElements created.");
            update_pbar.Maximum = 1000;
            update_pbar.Value = 100;
            FileInfo inputf = new FileInfo(Input_File.FullName);
            u.dbg("VonExcel: Opening up Excel Application.");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks xlWorkBooks = xlApp.Workbooks;
            Excel.Workbook xlWbk = xlWorkBooks.Open(inputf.FullName);
            Excel.Sheets xlWorksheets = xlWbk.Worksheets;
            Excel.Worksheet xlWks = xlWbk.ActiveSheet;
            Excel.Range xlRange = xlWks.UsedRange;
            update_pbar.Value += 10;
            u.dbg("VonExcel: Excel Application Loaded Successfully.");


            ProgressContext<Excel.Range> IEmRange =
                new ProgressContext<Excel.Range>(
                    from Excel.Range rng in xlRange
                    where rng.Value != null
                    select rng);
            u.dbg("VonExcel: Making LINQ Query for Non-Empty Used Ranges.");

            IEmRange.UpdateProgress += (sender0, e0) =>
            {
                update_pbar.Value += 1;
                //u.dbg(sender0.ToString());
                //u.dbg(e0.ToString());
            };
            u.dbg("VonExcel: UpDate Progress Behavior Delegated.");
            u.dbg("VonExcel: Creating Excel Grid XML.");

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

            u.dbg("VonExcel: Grid XML Created.");
            u.dbg("VonExcel: Populating Range XML.");
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
            u.dbg("VonExcel: Range XML Populated.");
            ProgressContext<Excel.Shape> IEmShapes =
                new ProgressContext<Excel.Shape>(
                    from Excel.Shape s in xlWks.Shapes
                    where s.Name.Contains("Check Box")
                    && (s.OLEFormat.Object.Value == 1)
                    select s);

            u.dbg("VonExcel: Making a LINQ Query for Checkbox.");
            IEmShapes.UpdateProgress += (sender0, e0) =>
            {
                update_pbar.Value += 1;
            };

            u.dbg("VonExcel: Populating Checkbox XML.");
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

            u.dbg("VonExcel: Checkbox XML Populated.");
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
            u.dbg("VonExcel: Closing Workbook Application, Recycling. Returning Value.");
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

    public class Utl
    {
        public void dbg(string Message) // Prints a debug information
        {
            //Format as below
            // [HH:MM:SS] (HH:MM:SS) [i]: Print Debug Message Here
            this.This_time = DateTime.Now;
            TimeSpan exect = this.This_time - this.Last_time;
            StringBuilder Message_strb = new StringBuilder();

            Message_strb.Append("[Debug: " + Run_Cycle + "] ");
            Message_strb.Append("[" + This_time.ToString("hh:MM:ss") + "] ");
            Message_strb.Append("<" + exect.ToString() + ">: ");
            Message_strb.Append(Message);
            Debug.Print(Message_strb.ToString());
            this.Last_time = this.This_time;
            this.Total_exect += exect;
            this.Run_Cycle++;
        }

        public void dbg_done()
        {
            //Complete Procedure report
            this.dbg("Procedure Completed, Total TimeSpan: " + this.Total_exect.ToString());
        }
        private DateTime This_time { get; set; } = DateTime.Now;
        private DateTime Last_time { get; set; } = DateTime.Now;
        private TimeSpan Total_exect { get; set; } = TimeSpan.FromSeconds(0);
        private int Run_Cycle { get; set; } = 1;
    }
}
