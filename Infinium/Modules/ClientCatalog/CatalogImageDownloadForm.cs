using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class CatalogImageDownloadForm : Form
    {
        public static int Result;//0 cancel, 1 open, 2 save

        private FileManager fm;
        private DecorCatalog decorCatalog;
        private int configId;
        private string fileName;

        private bool bStopTransfer;

        public CatalogImageDownloadForm(int configId, string fileName, FileManager fm, DecorCatalog decorCatalog)
        {
            InitializeComponent();
            this.fm = fm;
            this.decorCatalog = decorCatalog;
            this.configId = configId;
            this.fileName = fileName;

        }

        private Thread T;
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (fm.TotalFileSize == 0)
                return;

            if (fm.Position == fm.TotalFileSize || fm.Position > fm.TotalFileSize)
            {
                ProgressBar.Value = 100;
                timer1.Enabled = false;
                return;
            }

            DownloadedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(fm.Position / 1024)) + " / " +
                                   FileManager.GetIntegerWithThousands(Convert.ToInt32((fm.TotalFileSize / 1024))) + " КБайт";

            SpeedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(fm.CurrentSpeed)) + " КБайт/c";

            ProgressBar.Value = Convert.ToInt32(100 * fm.Position / fm.TotalFileSize);
            PercentsLabel.Text = ProgressBar.Value + " %";
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void CatalogImageDownloadForm_Load(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "(*" + Path.GetExtension(fileName) + ")|*" + Path.GetExtension(fileName);
            saveFileDialog1.FileName = fileName;

            DialogResult dialogResult = saveFileDialog1.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                timer1.Enabled = true;

                ProgressBar.Visible = true;

                T = new Thread(delegate () { decorCatalog.SaveDecorTechStoreFile(configId, saveFileDialog1.FileName); });
                T.Start();

                while (T.IsAlive)
                {
                    T.Join(50);
                    Application.DoEvents();

                    if (bStopTransfer)
                    {
                        fm.bStopTransfer = true;
                        bStopTransfer = false;
                        timer1.Enabled = false;
                        return;
                    }
                }

                timer1.Enabled = false;
                Close();
                return;
            }

            bStopTransfer = true;

            if (T != null)
                T.Abort();


            timer1.Enabled = false;
            Close();
        }
    }
}
