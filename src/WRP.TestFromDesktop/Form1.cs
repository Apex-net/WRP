using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WRP.TestFromDesktop
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // init
            textBox1.Text = WRP.TestFromDesktop.Properties.Settings.Default.URI;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // load url
            var url = textBox1.Text;
            try
            {
                webBrowser1.Navigate(url);
            }
            catch (Exception ex)
            {
                 MessageBox.Show(String.Format("WRP.TestFromDesktop - Navigate : {0}", ex.Message));
            }
 
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var message = String.Format("WRP.TestFromDesktop - DocumentCompleted url : {0}", e.Url.ToString());
            Console.WriteLine(message);

            
        }
    }
}
