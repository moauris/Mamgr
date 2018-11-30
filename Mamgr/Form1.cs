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
            //This is part of the cbx solution: mark all cbx covered range black.
            IEnumerable<Excel.Shape> CheckBoxes = //Represents the collection of checked checkboxes
                from Excel.Shape s in xlWks.Shapes
                where s.Name.Contains("Check Box")
                && s.OLEFormat.Object.Value == 1
                select s;
            #region Sync All H Column Content
            string cbx_test = get_cbx_grp_value
                (xlWks.Range["H32,H33,J32"], CheckBoxes, u); //this test returns the result of option
            //Removed Checked Boxes from the collection of Checkboxes.
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
            Xserver.Add(new XElement("OS", get_cbx_grp_value(
                cbx_range, CheckBoxes, u)));
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
            //checked_Box = cbx_grp(
                //cbx_range,
                //CheckBoxes, u);
            Xserver.Add(new XElement("OS", get_cbx_grp_value(
                cbx_range, CheckBoxes, u)));
            Xserver.Add(new XElement("Version", grv("H73")));
            Xserver.Add(new XElement("Bit", grv("H74")));
            Xserver.Add(new XElement("Virtual_Split", grv("H75")));
            Xserver.Add(new XElement("Virtual_Index", grv("H76")));
            Xserver.Add(new XElement("VIP", grv("H49")));
            Xserver.Add(new XElement("PRI", grv("H51")));
            Xserver.Add(new XElement("SEC", grv("H64")));
            WriteTo.AppendText(Xserver.ToString() + "\r\n" + "\r\n");
            #endregion
            #endregion
            xlApp.Visible = true;

            
            xlWbk.SaveAs(@".\Changed_Application_Form.xlsx");
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
            if (checked_box.Count() > 1 || checked_box.Count() == 0)
            {
                throw new Exception("Error: None or More than one Checkbox Checked!");
            }
            else
            {
                return checked_box.First();
            }
        }
        private static string get_cbx_grp_value
            (Excel.Range rng, IEnumerable<Excel.Shape> IE_cbx, Utl u)
        {
            string result = "";
            u.dbg("get_cbx_grp_value: Start.");
            u.dbg("get_cbx_grp_value: There are " + rng.Count + " Cells");
            IEnumerable<string> checked_cell_Val =
                from Excel.Range r in rng
                from Excel.Shape s in IE_cbx
                where s.Top > r.Top && s.Top < r.Top + r.Height
                   && s.Left > r.Left && s.Left < r.Left + r.Width
                select (string)r.Offset[0, 1].Value;
            u.dbg("There are " + checked_cell_Val.Count() + " Checked Value Found");
            try
            {
                result = checked_cell_Val.First();
            }
            catch (Exception ex)
            {
                u.dbg("Getting value from checked_cell_val unsuccessful");
                u.dbg(ex.Message);
                result = "未入力";
            }

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
            Message_strb.Append("[" + This_time.ToString("hh:mm:ss") + "] ");
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
    public class Server_Info
    {
        private XElement GetAsXml { get; set; }
        private void GetFromRange
            ()
        {

        }
    }
}
