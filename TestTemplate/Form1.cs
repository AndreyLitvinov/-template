using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TemplateEngine.Docx;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TestTemplate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*FileStream fileStream = null;
            fileStream = System.IO.File.Open("C:\\001.docx", FileMode.Open);
            byte[] byteFile = new byte[fileStream.Length];
            fileStream.Read(byteFile, 0, Convert.ToInt32(fileStream.Length));
            fileStream.Close();

            byte[] byteArray = File.ReadAllBytes("c:\\001.docx");
            MemoryStream stream = new MemoryStream();
            stream.Write(byteArray, 0, (int)byteArray.Length);

            byte[] newArray = null;

            var valuesToFill = new Content(new FieldContent("test", "ololo"));

            using (var outputDocument = new TemplateProcessor("C:\\003.docx").SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();

                newArray= stream.GetBuffer();
                //File.WriteAllBytes("C:\\002.docx", newArray);
                stream.Close();
            }*/
            var valuesToFill = new Content(
				new FieldContent("El2", "Данные для вставки в контрол"));

            byte[] byteArray = File.ReadAllBytes("D:\\test2.docx");
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (var outputDocument = new TemplateProcessor("D:\\test2.docx").SetRemoveContentControls(true))
                {
					
					outputDocument.FillContent(valuesToFill);
					outputDocument.SaveChanges(s => s.StartsWith("Title") || s.StartsWith("block"));
                    richTextBox1.Text = outputDocument.Document.ToString();
                }
                // Save the file with the new name
                File.WriteAllBytes("D:\\newFileName.docx", stream.ToArray());
            }

        }
       
        /*public static byte[] InsertCustomXml(byte[] document, string xmlString)
        {
            //System.IO.File.WriteAllText(@"d:\temp\template\fullXML.xml", xmlString); 

            MemoryStream documentStream = new MemoryStream();
            documentStream.Write(document, 0, document.Length); // must be done this way so that the memorystream is expandable 

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentStream, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                    CustomXmlPart customXmlPart1 = (CustomXmlPart)mainPart.GetPartById("rId1");

                    System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.Unicode);
                    writer.WriteRaw(xmlString);
                    writer.Flush();
                    writer.Close();

                    wordDoc.MainDocumentPart.PutXDocument();
                }
            }
            catch (Exception ex)
            {
                Helpers.DiagnosticHelper.RecordProblem("issue in InsertCustomXml", ex);
                return null;
            }
            
            //MemoryStream ms = new MemoryStream();
            //documentStream.WriteTo(ms); // It is important to write to this second stream - otherwise data is messed up 
            //return ms.GetBuffer();
        }*/

    }
}
