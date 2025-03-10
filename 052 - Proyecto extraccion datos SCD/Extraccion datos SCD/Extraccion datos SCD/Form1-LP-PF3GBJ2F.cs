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

namespace Extraccion_datos_SCD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            procesar.Enabled = false;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "C:\\Users\\JZ4874\\Desktop\\";
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

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

        private void procesar_Click(object sender, EventArgs e)
        {
            var xml = XDocument.Load(routeSCD.Text);
            XNamespace ns = "http://www.iec.ch/61850/2003/SCL";

            var ieds = xml.Root.Descendants(ns + "IED");

            Console.WriteLine("Number of ieds: " + ieds.Count());

            for (int i = 0; i < ieds.Count(); i++)
            {
                var ied = ieds.ElementAt(i);

                Console.WriteLine("IED name: " + ied.Attribute("name"));


                var datasets = xml.Descendants(ns + "DataSet")
                                  .Where(x => (((string)x.Parent.Parent.Parent.Parent.Parent.Attribute("name") == (string)ied.Attribute("name"))));

                Console.WriteLine("Number of datasets: " + datasets.Count());

                for (int j = 0; j < datasets.Count(); j++)
                {
                    var dataset = datasets.ElementAt(j);

                    var values = xml.Descendants(ns + "FCDA")
                          .Where(x => (((string)x.Parent.Parent.Parent.Parent.Parent.Parent.Attribute("name") == (string)ied.Attribute("name")) &&
                                       ((string)x.Parent.Attribute("name") == (string)dataset.Attribute("name"))));

                    Console.WriteLine("Number of values in dataset: " + values.Count());
                }


                //.Descendants("DataSet");

                ////var datasets = ied.Descendants(ns + "IED\\AccessPoint\\Server\\LDevice\\LN0");
                ////var datasets = ied.Descendants("LN0");
                ////var datasets = ied.Descendants("AccessPoint");

                ////var datasets = xml.Element("IED").
                ////.Where(n => (string)n.Attribute("name") = ied.Attribute("name").Value);

                ////var datasets = ((XDocument) ied).Descendants();
                ////var datasets = ied.SelectMany(a => (a.NextNode as XElement).Descendants("LN0"));

                //Console.WriteLine("Number of datasets: " + datasets.Count());

                //for (int j = 0; j < datasets.Count(); i++)
                //{
                //    var dataset = datasets.ElementAt(j);

                //    var variables = dataset.Descendants("Dataset");

                //    //Console.WriteLine("Number of datasets: " + variables.Count());

                //    for(int k = 0; k < variables.Count(); k++)
                //    {
                //        //Console.WriteLine(variables.ElementAt(k).ToString());
                //    }
                //}
            }


            MessageBox.Show("Se ha finalizado el proceso");


                                //.Where(x => (string)x.Attribute("lnClass") == "LLN0");
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
