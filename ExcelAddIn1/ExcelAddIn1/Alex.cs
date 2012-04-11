using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Net;
using Microsoft.Office.Core;
using System.IO;



namespace ExcelAddIn1
{
    public partial class Alex
    {
        static Worksheet _sheet; //current worksheet
        static Workbook _book;
        static Range _activeCell = null;
        static int _lineStyle;
        static int _borderWeight;

        static Network net = null;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            string workbookPath = @"D:\svn\sissi\trunk\examples\SACHS_for_Winograd_neu.xlsx";
            //register a change in the selection on a sheet
            //!!! remember: always keep a reference to the variable (make _sheet a class variable) otherwise it is
            // considered not used and disposed => only one trigger of the event
            _book = Globals.ThisAddIn.Application.Workbooks.Open(workbookPath,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, true);
            _sheet = Globals.ThisAddIn.Application.ActiveSheet;
            _sheet.SelectionChange += new DocEvents_SelectionChangeEventHandler(sheet_SelectionChange);
            Globals.ThisAddIn.Application.ActiveWorkbook.BeforeClose += new WorkbookEvents_BeforeCloseEventHandler(beforeClose);
        }

        private void beforeClose(ref bool Cancel)
        {
            if (net != null)
            {
                net.stop();
            }
        }

        private static void removeBorder()
        {
            if (_activeCell != null)
            {
                // TODO: this is a hack, but it works for our test case
                // the borders should be saved and then restored
                foreach (Border b in _activeCell.Borders)
                {
                    b.LineStyle = XlLineStyle.xlLineStyleNone;
                }
            }
        }

        private static void addBorder()
        {
            _activeCell = Globals.ThisAddIn.Application.ActiveCell;
            _lineStyle = _activeCell.Borders.LineStyle;
            _borderWeight = _activeCell.Borders.Weight;
            _activeCell.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
        }

        private void sheet_SelectionChange(Range rng)
        {
           
           
            if (net != null)
            {
                removeBorder();
                addBorder();                
                // get the A1Format address using Globals.ThisAddIn.Application.ActiveCell.Address
                // get the coordinates using CellTopLeftPixels(Globals.ThisAddIn.Application.ActiveCell)               
                Display.CellTopLeftPixels(rng);

                Dictionary<string, string> dict = new Dictionary<string, string>();
                dict.Add("select", Globals.ThisAddIn.Application.ActiveCell.Address.ToString());
                dict.Add("pos", Display.left + "," + Display.top);
                Message m = new Message("alex.click", dict);
                net.send(Json.serialize(m));
            }
            else
            {
                //MessageBox.Show(Globals.ThisAddIn.Application.ActiveCell.Address);
                //do nothing?
            }
        }
        


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //event triggger
            //this.label3.Label = "";
            //CellTopLeftPixels(Globals.ThisAddIn.Application.ActiveCell);

            if (net == null)
            {
                net = new Network(this);
            }
            else
            {
                MessageBox.Show("Sally already running. Disconnecting and reconnecting.");
                net.stop();
                net = new Network(this);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(Globals.ThisAddIn.Application.ActiveCell.Count.ToString() + " " + Globals.ThisAddIn.Application.ActiveCell.MergeCells);            
        }

        public void parseMessage(String str)
        {
            Message m = Json.deserialize(str);
            switch (m.action)
            {
                case "init":
                    //send the whoami message
                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    dict.Add("type", "alex");
                    dict.Add("doctype", "spreadsheet");
                    dict.Add("setup", "desktop");
                    Message whoami = new Message("whoami", dict);                    
                    net.send(Json.serialize(whoami));

                    dict.Clear();
                    //send the alex.imap message
                    dict.Add("imap", "test");
                    Message imap = new Message("alex.imap", dict);                    
                    net.send(Json.serialize(imap));

                    break;
                case "alex.select":                    
                default:
                    MessageBox.Show(str);
                    break;
            }
        }
    }
}
