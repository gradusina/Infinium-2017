﻿using System.ComponentModel;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;

namespace Infinium
{
    partial class CatalogImageDownloadForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.kryptonBorderEdge20 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge21 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge22 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge23 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.StandardButtonsPalette = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonBorderEdge19 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.kryptonBorderEdge2 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SpeedLabel = new System.Windows.Forms.Label();
            this.DownloadedLabel = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.PercentsLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonBorderEdge20
            // 
            this.kryptonBorderEdge20.Dock = System.Windows.Forms.DockStyle.Right;
            this.kryptonBorderEdge20.Location = new System.Drawing.Point(372, 0);
            this.kryptonBorderEdge20.Name = "kryptonBorderEdge20";
            this.kryptonBorderEdge20.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge20.Size = new System.Drawing.Size(1, 144);
            this.kryptonBorderEdge20.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge20.Text = "kryptonBorderEdge20";
            // 
            // kryptonBorderEdge21
            // 
            this.kryptonBorderEdge21.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonBorderEdge21.Location = new System.Drawing.Point(0, 0);
            this.kryptonBorderEdge21.Name = "kryptonBorderEdge21";
            this.kryptonBorderEdge21.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge21.Size = new System.Drawing.Size(1, 144);
            this.kryptonBorderEdge21.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge21.Text = "kryptonBorderEdge21";
            // 
            // kryptonBorderEdge22
            // 
            this.kryptonBorderEdge22.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonBorderEdge22.Location = new System.Drawing.Point(1, 143);
            this.kryptonBorderEdge22.Name = "kryptonBorderEdge22";
            this.kryptonBorderEdge22.Size = new System.Drawing.Size(371, 1);
            this.kryptonBorderEdge22.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge22.Text = "kryptonBorderEdge22";
            // 
            // kryptonBorderEdge23
            // 
            this.kryptonBorderEdge23.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonBorderEdge23.Location = new System.Drawing.Point(1, 0);
            this.kryptonBorderEdge23.Name = "kryptonBorderEdge23";
            this.kryptonBorderEdge23.Size = new System.Drawing.Size(371, 1);
            this.kryptonBorderEdge23.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge23.Text = "kryptonBorderEdge23";
            // 
            // StandardButtonsPalette
            // 
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(177)))), ((int)(((byte)(229)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(167)))), ((int)(((byte)(214)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color1 = System.Drawing.Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Color2 = System.Drawing.Color.Gray;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.OverrideDefault.Border.Rounding = 0;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(177)))), ((int)(((byte)(229)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.Color2 = System.Drawing.Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color1 = System.Drawing.Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Color2 = System.Drawing.Color.Black;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Border.Rounding = 0;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Color2 = System.Drawing.Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.MultiLine = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.MultiLineH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateCommon.Content.ShortText.TextV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.Color2 = System.Drawing.Color.White;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color1 = System.Drawing.Color.Transparent;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Color2 = System.Drawing.Color.Black;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.StandardButtonsPalette.ButtonStyles.ButtonCommon.StateTracking.Border.Rounding = 0;
            this.StandardButtonsPalette.CustomisedKryptonPaletteFilePath = null;
            // 
            // kryptonBorderEdge19
            // 
            this.kryptonBorderEdge19.AutoSize = false;
            this.kryptonBorderEdge19.Location = new System.Drawing.Point(27, 424);
            this.kryptonBorderEdge19.Name = "kryptonBorderEdge19";
            this.kryptonBorderEdge19.Size = new System.Drawing.Size(810, 1);
            this.kryptonBorderEdge19.StateCommon.Color1 = System.Drawing.Color.DarkGray;
            this.kryptonBorderEdge19.Text = "kryptonBorderEdge19";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(373, 32);
            this.panel1.TabIndex = 239;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Light", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 1);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(176, 28);
            this.label2.TabIndex = 240;
            this.label2.Text = "Сохранение файла";
            // 
            // kryptonBorderEdge2
            // 
            this.kryptonBorderEdge2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonBorderEdge2.AutoSize = false;
            this.kryptonBorderEdge2.Location = new System.Drawing.Point(-480, 77);
            this.kryptonBorderEdge2.Name = "kryptonBorderEdge2";
            this.kryptonBorderEdge2.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge2.Size = new System.Drawing.Size(4, 0);
            this.kryptonBorderEdge2.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(150)))), ((int)(((byte)(201)))));
            this.kryptonBorderEdge2.StateCommon.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.kryptonBorderEdge2.Text = "kryptonBorderEdge2";
            // 
            // ProgressBar
            // 
            this.ProgressBar.Location = new System.Drawing.Point(31, 79);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(310, 10);
            this.ProgressBar.TabIndex = 254;
            this.ProgressBar.Visible = false;
            // 
            // timer1
            // 
            this.timer1.Interval = 400;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // SpeedLabel
            // 
            this.SpeedLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.SpeedLabel.Location = new System.Drawing.Point(79, 98);
            this.SpeedLabel.Name = "SpeedLabel";
            this.SpeedLabel.Size = new System.Drawing.Size(262, 23);
            this.SpeedLabel.TabIndex = 255;
            this.SpeedLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // DownloadedLabel
            // 
            this.DownloadedLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.DownloadedLabel.Location = new System.Drawing.Point(61, 49);
            this.DownloadedLabel.Name = "DownloadedLabel";
            this.DownloadedLabel.Size = new System.Drawing.Size(280, 23);
            this.DownloadedLabel.TabIndex = 256;
            this.DownloadedLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // PercentsLabel
            // 
            this.PercentsLabel.AutoSize = true;
            this.PercentsLabel.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.PercentsLabel.ForeColor = System.Drawing.Color.Gray;
            this.PercentsLabel.Location = new System.Drawing.Point(27, 98);
            this.PercentsLabel.Name = "PercentsLabel";
            this.PercentsLabel.Size = new System.Drawing.Size(0, 20);
            this.PercentsLabel.TabIndex = 257;
            this.PercentsLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // CatalogImageDownloadForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(373, 144);
            this.Controls.Add(this.PercentsLabel);
            this.Controls.Add(this.DownloadedLabel);
            this.Controls.Add(this.SpeedLabel);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.kryptonBorderEdge2);
            this.Controls.Add(this.kryptonBorderEdge19);
            this.Controls.Add(this.kryptonBorderEdge23);
            this.Controls.Add(this.kryptonBorderEdge22);
            this.Controls.Add(this.kryptonBorderEdge21);
            this.Controls.Add(this.kryptonBorderEdge20);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "CatalogImageDownloadForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Infinium";
            this.Load += new System.EventHandler(this.CatalogImageDownloadForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private KryptonBorderEdge kryptonBorderEdge20;
        private KryptonBorderEdge kryptonBorderEdge21;
        private KryptonBorderEdge kryptonBorderEdge22;
        private KryptonBorderEdge kryptonBorderEdge23;
        private KryptonBorderEdge kryptonBorderEdge19;
        private Panel panel1;
        private Label label2;
        private KryptonBorderEdge kryptonBorderEdge2;
        private KryptonPalette StandardButtonsPalette;
        private ProgressBar ProgressBar;
        private Timer timer1;
        private Label SpeedLabel;
        private Label DownloadedLabel;
        private SaveFileDialog saveFileDialog1;
        private Label PercentsLabel;
    }
}