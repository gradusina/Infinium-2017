﻿
using System;
using System.Windows.Forms;

namespace Infinium
{
    public partial class NewCabFurRackForm : Form
    {
        public bool bOk = false;
        public string newName = string.Empty;

        Form TopForm = null;

        public NewCabFurRackForm(string oldName)
        {
            InitializeComponent();

            tbName.Text = oldName;
        }

        public NewCabFurRackForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbName.Text))
            {
                Infinium.LightMessageBox.Show(ref TopForm, false,
                                    "Некорректное название",
                                    "Ошибка");
                return;
            }

            newName = tbName.Text;

            bOk = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == 6 && m.WParam.ToInt32() == 1)
            {
                if (TopForm != null)
                    TopForm.Activate();
            }
        }
    }
}
