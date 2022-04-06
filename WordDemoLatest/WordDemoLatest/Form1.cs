using System;
using System.Windows.Forms;
using Xceed.Words.NET;
using System.Diagnostics;
namespace WordDemoLatest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        // First Button Now We can start coding part
        // set the file path to create which can be local path or global path
        // here i ve set global path instead of local path, because this app will be run on all machine (target all machine instead of one)
        // this is optional one.
        // this global path is done by using the Special Class Environment
        string gpath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"/Akshay.docx";
        private void b1_Click(object sender, EventArgs e)
        {
    // check the validation
    // if text box is not empty then write code for docx generation else provide validation via Message Box alert message
            if(mb.Text!="")
            {
                // get the user message dynmaically via Text Property of Text Box
                string userMessage = mb.Text;
                // create a new object for DocX class for the preparation of word document
                DocX obj=DocX.Create(gpath);
                // add paragraphs to new document via DocX obj
                obj.InsertParagraph(userMessage);
                // finally save the document using Save() or SaveAs() method
                obj.Save();
                // that's all. next step is to start / launch the newly created word document automatically via C#
                // for that just use Static method of Process class Start()
                // where first argument is the external application name to launch. here our application is word.
                // so that, winword.exe is used
                // second argument is the exact location of newly created word document which is saved in the variable gpath
                Process.Start("winword.exe", gpath);
                // that's all. Now we can run our application
            }
            else
            {
        // create an alert box using MessageBox class
        // first argument is message and second argument is heading of the message box
                MessageBox.Show("Text Box is Empty\nPlease type your Message...", "Empty");
            }
        }
// Second Button
        private void clear(object sender, EventArgs e)
        {
            // clear the text box
            mb.Text = "";
        }
    }
}
