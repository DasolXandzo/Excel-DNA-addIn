using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;


namespace Excel_DNA
{
    public partial class MyForm : Form
    {
        public MyForm()
        {
            InitializeComponent();

        }

        private void webView21_Click(object sender, EventArgs e)
        {
            webView21.Cursor = Cursors.WaitCursor;
            webView21.CoreWebView2.Navigate("https://www.bing.com");
            var test2 = 5;
        }
        private async Task initizated()
        {
            await webView21.EnsureCoreWebView2Async(null);
        }
        public async void InitBrowser()
        {
            await initizated();
            webView21.CoreWebView2.Navigate("https://www.youtube.com/");
        }

        private void MyForm_Load_1(object sender, EventArgs e)
        {
            InitBrowser();
        }
    }
}
