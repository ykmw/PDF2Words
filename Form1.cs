using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

// Support PDF to words
using Spire.Pdf;
// Support C# file, IO
using System.IO;
// Open executable file at runtime
using System.Diagnostics;

namespace PDF2Word
{
    public partial class PDF : Form
    {
        public PDF()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        // define the global string variable
        string path = "";


        // Browse Button
        private void Openfile(object sender, EventArgs e)
        {
            // Create open file dialog object
            OpenFileDialog fd = new OpenFileDialog();
            // set filter options
            fd.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|(*.*)";

            // show the dialog using its instance method and verify it using
            // its built-in constant properties
            // it returns the full path of the file, if user selects any pdf file else returns null
            if (fd.ShowDialog() == DialogResult.OK)
            {
                // get the selected file path and save it to global variable path
                path = fd.FileName;
                // assign this selected file path to textBox1
                textBox1.Text = fd.FileName;
            }
        }
        // Convert Button
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // create an object for PDFDocument class
                PdfDocument obj = new PdfDocument();
                // load the selected pdf file (via open file dialog) to the instance method of the above class
                obj.LoadFromFile(path);

                // start the convert process and save the resultant word document using
                // its instance method called SaveToFile()
                obj.SaveToFile("ConvertW.docx", FileFormat.DOCX);

                // here i ve given the current project path of resultant word file
                // you can give either current path or aboslute path
                // now check whether the new docx will be created or not using Exists() method of File class
                // if the condition below returns true means, the output word document will be ready
                if (File.Exists("ConvertW.docx") == true)
                {
                    // now we need to open the created word document instantly
                    // this is done by using the static method of Process class called Start(path)
                    Process.Start("ConvertW.docx");
                    // that's all
                }
            }
            catch (Exception ext)
            {
                System.Console.WriteLine(ext.Message);
            }

        }
    }
}