using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using System.Xml.Schema;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography.X509Certificates;

namespace Extraccion_datos_SCD
{
    public partial class Form1 : Form
    {
        string ruta_planilla = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Planilla template GOOSE\\Planilla Consolidada de GOOSE.xlsx";
        //string ruta_planilla = "C:\\Users\\JZ4874\\Desktop\\Planilla GOOSE\\Planilla Consolidada de GOOSE.xlsx";

        public Form1()
        {
            InitializeComponent();
            procesar.Enabled = false;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openFileDialog1.InitialDirectory = "C:\\Users\\JZ4874\\Desktop\\";
            openFileDialog1.Filter = "SCD Files(.scd)|*.scd";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                procesar.Enabled = true;
                string selectedFileName = openFileDialog1.FileName;
                routeSCD.Text = selectedFileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "C:\\Users\\JZ4874\\Desktop\\";
            saveFileDialog1.Filter = "Excel Files(.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            else
            {
                string selectedFileName = saveFileDialog1.FileName;
                savePath.Text = selectedFileName;
            }
        }

        
        public static bool printStatement(string value)
        {
            Console.WriteLine(value);
            return true;
        }
        
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

        private void procesar_Click(object sender, EventArgs e)
        {
            if(!(System.IO.File.Exists(ruta_planilla)))
            {
                MessageBox.Show("No se pudo encontrar el template de la planilla consolidada de GOOSE.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Ruta planilla: " + ruta_planilla);
                
            } 
            else if (Path.GetExtension(savePath.Text) != ".xlsx")
            {
                MessageBox.Show("No es válida la ruta de guardado.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                System.IO.File.Copy(ruta_planilla, savePath.Text, true);

                Excel.Application app = new Excel.Application();
                app.Visible = true;
                app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                Excel.Workbook book = app.Workbooks.Open(savePath.Text);

                Dictionary<string, int> references = new Dictionary<string, int>();

                var xml = XDocument.Load(routeSCD.Text);

                XNamespace ns = xml.Root.Attribute("xmlns").Value;

                var ieds = xml.Root.Descendants(ns + "IED");

                int row = 7;
                int col = 6;

                for (int i = 0; i < ieds.Count(); i++)
                {
                    var ied = ieds.ElementAt(i);

                    var ext_references = xml.Root.Descendants(ns + "ExtRef")
                            .Where(x => ((x.Parent.Parent.Parent.Parent.Parent.Parent.Attribute("name") == ied.Attribute("name"))
                                         && (x.Attribute("daName").Value != "q")
                                         && (x.Attribute("daName").Value != "t")));

                    for (int j = 0; j < ext_references.Count(); j++)
                    {
                        row++;
                        col++;

                        int ref_col = col;

                        var block_nbr = "";
                        var input = "";
                        var ext_ref = ext_references.ElementAt(j);

                        if (references.ContainsKey(ext_ref.Attribute("iedName").Value
                                    + ext_ref.Attribute("ldInst").Value + ext_ref.Attribute("prefix").Value
                                    + ext_ref.Attribute("lnClass").Value + ext_ref.Attribute("lnInst").Value
                                    + ext_ref.Attribute("doName").Value + ext_ref.Attribute("daName").Value))
                        {
                            col = references[ext_ref.Attribute("iedName").Value
                                    + ext_ref.Attribute("ldInst").Value + ext_ref.Attribute("prefix").Value
                                    + ext_ref.Attribute("lnClass").Value + ext_ref.Attribute("lnInst").Value
                                    + ext_ref.Attribute("doName").Value + ext_ref.Attribute("daName").Value];
                        } 
                        else {
                            references.Add(ext_ref.Attribute("iedName").Value
                                    + ext_ref.Attribute("ldInst").Value + ext_ref.Attribute("prefix").Value
                                    + ext_ref.Attribute("lnClass").Value + ext_ref.Attribute("lnInst").Value
                                    + ext_ref.Attribute("doName").Value + ext_ref.Attribute("daName").Value, col);
                        }

                        if(ext_ref.Attribute("intAddr") != null)
                        {
                            var int_addr = (ext_ref.Attribute("intAddr").Value).Split(new string[] { "GOOSE" }, StringSplitOptions.None);

                            if (int_addr.Count() > 1)
                            {
                                int_addr = int_addr[int_addr.Count() - 1].Split(new string[] { "IN" }, StringSplitOptions.None);

                                if (int_addr.Count() > 1)
                                {
                                    block_nbr = int_addr[0].Replace(".", "");

                                    var input_arr = int_addr[1].Split(new string[] { "Value" }, StringSplitOptions.None);

                                    if (input_arr.Count() > 1)
                                    {
                                        input = input_arr[0].Replace(".", "");
                                    }
                                }
                            }
                        }

                        if(ext_ref.Attribute("intAddr") != null)
                        {
                            var int_addr = (ext_ref.Attribute("intAddr").Value).Split(new string[] { "GOOSE" }, StringSplitOptions.None);

                            if(int_addr.Count() > 1)
                            {
                                int_addr = int_addr[int_addr.Count() - 1].Split(new string[] { "Value" }, StringSplitOptions.None);

                                app.ActiveWorkbook.Sheets[2].Cells[row, 5] = "GOOSE" + int_addr[0];
                            }
                            else
                            {
                                app.ActiveWorkbook.Sheets[2].Cells[row, 5] = (ext_ref.Attribute("intAddr").Value).Split('/').Last();
                            }
                        }

                        app.ActiveWorkbook.Sheets[2].Cells[row, 2] = ied.Attribute("name").Value;

                        app.ActiveWorkbook.Sheets[2].Cells[2, col] = ext_ref.Attribute("iedName").Value;

                        var function_desc = xml.Root.Descendants(ns + "DAI")
                                        .Where(x => ((x.Parent.Parent.Parent.Parent.Parent.Parent.Attribute("name")) != null)
                                            && (x.Parent.Parent.Parent.Parent.Parent.Parent.Attribute("name").Value == ext_ref.Attribute("iedName").Value)
                                            && (x.Parent.Parent.Parent.Attribute("inst") != null)
                                            && (x.Parent.Parent.Parent.Attribute("inst").Value == ext_ref.Attribute("ldInst").Value)
                                            && (x.Parent.Parent.Attribute("prefix") != null)
                                            && (x.Parent.Parent.Attribute("prefix").Value == ext_ref.Attribute("prefix").Value)
                                            && (x.Parent.Parent.Attribute("lnClass") != null)
                                            && (x.Parent.Parent.Attribute("lnClass").Value == ext_ref.Attribute("lnClass").Value)
                                            && (x.Parent.Parent.Attribute("inst") != null)
                                            && (x.Parent.Parent.Attribute("inst").Value == ext_ref.Attribute("lnInst").Value)
                                            && (x.Parent.Attribute("name") != null)
                                            && (x.Parent.Attribute("name").Value == ext_ref.Attribute("doName").Value)
                                            && (x.Attribute("name") != null)
                                            && (x.Attribute("name").Value == ext_ref.Attribute("daName").Value));

                        if(function_desc.Count() > 0)
                        {
                            if(function_desc.FirstOrDefault().Attribute("desc") != null)
                            {
                                app.ActiveWorkbook.Sheets[2].Cells[4, col] = function_desc.FirstOrDefault().Attribute("desc").Value;
                            }
                        }

                        app.ActiveWorkbook.Sheets[2].Cells[row, 4] = app.ActiveWorkbook.Sheets[2].Cells[4, col];

                        app.ActiveWorkbook.Sheets[2].Cells[5, col] = ext_ref.Attribute("ldInst").Value + ext_ref.Attribute("prefix").Value
                                                            + ext_ref.Attribute("lnClass").Value + ext_ref.Attribute("lnInst").Value
                                                            + ext_ref.Attribute("doName").Value + ext_ref.Attribute("daName").Value;

                        app.ActiveWorkbook.Sheets[2].Cells[row, col] = "X";

                        if(col != ref_col) col = ref_col - 1;
                    }
                }

                book.Save();
                System.Threading.Thread.Sleep(2000);
                app.Quit();

                MessageBox.Show("Se ha terminado la generación de la planilla de datos consolidados GOOSE",
                    "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);


                System.Diagnostics.Process.Start("explorer.exe", Path.GetDirectoryName(savePath.Text));


                //var query = xml.Root.Descendants()
                //                    .Where(x => (string)x.Attribute("lnClass") == "LLN0");
                //.Where(x => (string)x.Attribute("lnClass") == "LLN0").FirstOrDefault();

                //MessageBox.Show(query.ToString());


                //MessageBox.Show("Number of elements: " + query.Count());
                //MessageBox.Show(string.Join(",", (IEnumerable<XElement>)query.ToArray()));


                //XDocument doc = new XDocument();
                //doc = XDocument.Parse("<parent><foo id='bar' option='12345'/><foo id='bar' option='abcde'/></parent>");
                ////doc.Load(routeSCD.Text);

                //string idToFind = "bar";

                //var selectedElement = doc.Descendants()
                //        .Where(x => (string)x.Attribute("id") == idToFind);
                ////.Where(x => (string)x.Attribute("id") == idToFind).FirstOrDefault();

                //MessageBox.Show(string.Join(",", (IEnumerable<XElement>) selectedElement.ToArray()));
                ////MessageBox.Show(selectedElement.ToString());
                ////Console.WriteLine(selectedElement);
            }
        }
    }
}
