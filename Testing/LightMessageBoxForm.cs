using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class LightMessageBoxForm : Form
    {
        public static bool OKCancel;

        public LightMessageBoxForm(bool ShowCancelButton, string Text, string Header)
        {
            InitializeComponent();

            label1.Text = Text;

            CancelMessageButton.Visible = ShowCancelButton;

            if (ShowCancelButton == false)
                OKMessageButton.Left = (Width - OKMessageButton.Width) / 2;

            label2.Text = Header;
        }
        
        public LightMessageBoxForm(bool ShowCancelButton, string Text, string Header, string OKMessageButtonText, string CancelMessageButtonText)
        {
            InitializeComponent();
            OKMessageButton.Text = OKMessageButtonText;
            CancelMessageButton.Text = CancelMessageButtonText;
            label1.Text = Text;

            CancelMessageButton.Visible = ShowCancelButton;

            if (ShowCancelButton == false)
                OKMessageButton.Left = (Width - OKMessageButton.Width) / 2;

            label2.Text = Header;
        }

        private void OKMessageButton_Click(object sender, EventArgs e)
        {
            OKCancel = true;

            Close();
        }

        private void CancelMessageButton_Click(object sender, EventArgs e)
        {
            OKCancel = false;

            Close();
        }

        private void LightMessageBoxForm_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //    OKMessageButton_Click(null, null);
            //if (e.KeyCode == Keys.Cancel)
            //    CancelMessageButton_Click(null, null);
        }
    }
}
