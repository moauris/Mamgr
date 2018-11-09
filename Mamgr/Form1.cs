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
        delegate string Get_Range_Value(string Addr);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Branch branch_vonExcel_Direct: Get xml document formated to be Agent structured.
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
            u.dbg("Temp Disabled.");
            u.dbg_done();
            return;
            /*
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
            u.dbg_done(); */

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

            #region Excel Application Open, Sync Content
            Get_Range_Value grv = (addr) =>
            {
                if (xlWks.Range[addr].Value == null)
                    return "未入力";
                else
                    return xlWks.Range[addr].Value.ToString();
            };
            //Debug 1: try to fetch application_file info
            XElement Xapp_file = new XElement(xlWbk.Name);
            //Need a string to represent date
            Xapp_file.Add(new XElement("Apply_Date",
                (Convert.ToDateTime(grv("H7")))
                .ToShortDateString()));
            Xapp_file.Add(new XElement("Applicant", grv("H8")));
            Xapp_file.Add(new XElement("Email", grv("H9")));
            Xapp_file.Add(new XElement("Phone", grv("H10")));
            Xapp_file.Add(new XElement("Approver", grv("H11")));
            Xapp_file.Add(new XElement("INC_Bango")); 
            //Need a procedure to fetch INC Bango
            WriteTo.AppendText(Xapp_file.ToString() + "\r\n");
            u.dbg("VonExcel: Fetched app_file Info -- 5 cells");
            //Next, fetch h Column M/Agent
            //Need a macro for getting check box value from a group of cells.
            // private string get_cbx_grp_value (Excel.Range rng, IEnumerable<Excel.Shape> IE_cbx) {}
            IEnumerable<Excel.Shape> CheckBoxes = //Represents the collection of checked checkboxes
                from Excel.Shape s in xlWks.Shapes
                where s.Name.Contains("Check Box")
                && s.OLEFormat.Object.Value == 1
                select s;
            Excel.Shape checked_Box = 
                cbx_grp(xlWks.Range["H32,H33,J32"], CheckBoxes, u);
            string cbx_test = get_cbx_grp_value
                (xlWks.Range["H32,H33,J32"], checked_Box, u); //this test returns the result of option
            CheckBoxes = CheckBoxes.Where(s0 => s0.Name != checked_Box.Name).ToList();
            WriteTo.AppendText(cbx_test + "\r\n" + "\r\n");

            XElement Xserver = new XElement(grv("H49"));
            Xserver.Add(new XElement("IP_Addr", grv("H50")));
            Xserver.Add(new XElement("VIP", grv("H49")));
            Xserver.Add(new XElement("PRI", grv("H51")));
            Xserver.Add(new XElement("SEC", grv("H64")));
            WriteTo.AppendText(Xserver.ToString() + "\r\n" + "\r\n");

            Xserver = new XElement(grv("H51"));
            Xserver.Add(new XElement("IP_Addr", grv("H52")));
            Xserver.Add(new XElement("Maker", grv("H53")));
            Xserver.Add(new XElement("Model", grv("H54")));
            Xserver.Add(new XElement("CPU_Num", grv("H55")));
            Xserver.Add(new XElement("CPU_Micro", grv("H56")));
            Excel.Range cbx_range = 
                xlWks.Range["H57,H58,H59,J57,J58"];
            checked_Box = cbx_grp(
                cbx_range,
                CheckBoxes, u);
            Xserver.Add(new XElement("OS", get_cbx_grp_value(
                cbx_range, checked_Box, u)));
            CheckBoxes = CheckBoxes.Where(s0 => s0.Name != checked_Box.Name).ToList();
            Xserver.Add(new XElement("Version", grv("H60")));
            Xserver.Add(new XElement("Bit", grv("H61")));
            Xserver.Add(new XElement("Virtual_Split", grv("H62")));
            Xserver.Add(new XElement("Virtual_Index", grv("H63")));

            Xserver.Add(new XElement("VIP", grv("H49")));
            Xserver.Add(new XElement("PRI", grv("H51")));
            Xserver.Add(new XElement("SEC", grv("H64")));
            WriteTo.AppendText(Xserver.ToString() + "\r\n" + "\r\n");

            Xserver = new XElement(grv("H64"));
            Xserver.Add(new XElement("IP_Addr", grv("H65")));
            Xserver.Add(new XElement("Maker", grv("H66")));
            Xserver.Add(new XElement("Model", grv("H67")));
            Xserver.Add(new XElement("CPU_Num", grv("H68")));
            Xserver.Add(new XElement("CPU_Micro", grv("H69")));
            cbx_range =
                xlWks.Range["H70,H71,H72,J70,J71"];
            checked_Box = cbx_grp(
                cbx_range,
                CheckBoxes, u);
            Xserver.Add(new XElement("OS", get_cbx_grp_value(
                cbx_range, checked_Box, u)));
            CheckBoxes = CheckBoxes.Where(s0 => s0.Name != checked_Box.Name).ToList();
            Xserver.Add(new XElement("Version", grv("H73")));
            Xserver.Add(new XElement("Bit", grv("H74")));
            Xserver.Add(new XElement("Virtual_Split", grv("H75")));
            Xserver.Add(new XElement("Virtual_Index", grv("H76")));
            Xserver.Add(new XElement("VIP", grv("H49")));
            Xserver.Add(new XElement("PRI", grv("H51")));
            Xserver.Add(new XElement("SEC", grv("H64")));
            WriteTo.AppendText(Xserver.ToString() + "\r\n" + "\r\n");

            #endregion

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
        private static Excel.Shape cbx_grp
            (Excel.Range rng, IEnumerable<Excel.Shape> IE_cbx, Utl u)
        {// returns checked box within a range
            u.dbg("cbx_grp: Start.");
            u.dbg("cbx_grp: There are " + rng.Count + " Cells");
            u.dbg("cbx_grp: There are " + IE_cbx.Count() + " CheckBoxes");
            IEnumerable<Excel.Shape> checked_box =
                from Excel.Range r in rng
                from Excel.Shape s in IE_cbx
                where s.Top > r.Top && s.Top < r.Top + r.Height
                   && s.Left > r.Left && s.Left < r.Left + r.Width
                select s;
            u.dbg("cbx_grp: End, return result");
            if (checked_box.Count() > 1)
            {
                throw new Exception("Error: More than one Checkbox Checked!");
            }
            else
            {
                return checked_box.First();
            }
        }
        private static string get_cbx_grp_value
            (Excel.Range rng, Excel.Shape cbx, Utl u)
        {
            string result = "";
            u.dbg("get_cbx_grp_value: Start.");
            u.dbg("get_cbx_grp_value: There are " + rng.Count + " Cells");
            IEnumerable<string> checked_cell_Val =
                from Excel.Range r in rng
                where cbx.Top > r.Top && cbx.Top < r.Top + r.Height
                   && cbx.Left > r.Left && cbx.Left < r.Left + r.Width
                select (string)r.Offset[0, 1].Value;
            //u.dbg("There are " + checked_cell.Count() + " Checked Cells Found");
            foreach (string r in checked_cell_Val)
                result = r;

            u.dbg("get_cbx_grp_value: End, returning result.");

            return result;
        }
    }
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
