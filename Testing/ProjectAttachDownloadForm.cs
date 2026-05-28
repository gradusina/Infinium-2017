using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Infinium
{
    public partial class ProjectAttachDownloadForm : Form
    {
        public static int Result;//0 cancel, 1 open, 2 save

        private FileManager FM;
        private InfiniumProjects InfiniumProjects;
        private int NewsAttachID;

        private bool bStopTransfer;

        public ProjectAttachDownloadForm(int iNewsAttachID, ref FileManager tFM, ref InfiniumProjects tInfiniumProjects)
        {
            InitializeComponent();

            NewsAttachID = iNewsAttachID;
            FM = tFM;
            InfiniumProjects = tInfiniumProjects;

        }

        private Thread T;

        private void OpenButton_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;

            label1.Visible = false;
            ProgressBar.Visible = true;

            string temppath = "";

            {
                T = new Thread(delegate ()
                { temppath = InfiniumProjects.SaveFile(NewsAttachID); });
                T.Start();

                OpenButton.Enabled = false;
                SaveButton.Enabled = false;

                while (T.IsAlive)
                {
                    T.Join(50);
                    Application.DoEvents();

                    if (bStopTransfer)
                    {
                        FM.bStopTransfer = true;
                        bStopTransfer = false;
                        timer1.Enabled = false;
                        return;
                    }
                }

                if (!bStopTransfer && temppath != null)
                    Process.Start(temppath);
            }

            timer1.Enabled = false;
            Close();
        }

        private void CancelMessageButton_Click(object sender, EventArgs e)
        {
            bStopTransfer = true;

            if (T != null)
                T.Abort();

            //while (T.IsAlive)
            //    System.Threading.Thread.Sleep(50);

            Close();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            {
                string FileName = InfiniumProjects.GetAttachmentName(NewsAttachID);
                saveFileDialog1.Filter = "(*" + Path.GetExtension(FileName) + ")|*" + Path.GetExtension(FileName);
                saveFileDialog1.FileName = FileName;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    timer1.Enabled = true;

                    label1.Visible = false;
                    ProgressBar.Visible = true;

                    T = new Thread(delegate ()
                    { InfiniumProjects.SaveFile(NewsAttachID, saveFileDialog1.FileName); });
                    T.Start();

                    OpenButton.Enabled = false;
                    SaveButton.Enabled = false;

                    while (T.IsAlive)
                    {
                        T.Join(50);
                        Application.DoEvents();

                        if (bStopTransfer)
                        {
                            FM.bStopTransfer = true;
                            bStopTransfer = false;
                            timer1.Enabled = false;
                            return;
                        }
                    }

                    timer1.Enabled = false;
                    Close();
                    return;
                }
            }

            timer1.Enabled = false;
            Close();
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            if (FM.TotalFileSize == 0)
                return;

            if (FM.Position == FM.TotalFileSize || FM.Position > FM.TotalFileSize)
            {
                ProgressBar.Value = 100;
                timer1.Enabled = false;
                return;
            }

            DownloadedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(FM.Position / 1024)) + " / " +
                                   FileManager.GetIntegerWithThousands(Convert.ToInt32((FM.TotalFileSize / 1024))) + " КБайт";

            SpeedLabel.Text = FileManager.GetIntegerWithThousands(Convert.ToInt32(FM.CurrentSpeed)) + " КБайт/c";

            ProgressBar.Value = Convert.ToInt32(100 * FM.Position / FM.TotalFileSize);
            PercentsLabel.Text = ProgressBar.Value + " %";
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //LightNews.SaveFile(NewsAttachID, System.IO.Path.GetFileName(saveFileDialog1.FileName), saveFileDialog1.FileName);
        }
    }
}
