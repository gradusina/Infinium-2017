﻿using System.ComponentModel;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;

namespace Infinium
{
    partial class AddProjectNewsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddProjectNewsForm));
            this.kryptonBorderEdge20 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge21 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge22 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.kryptonBorderEdge23 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.OKNewsButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.StandardButtonsPalette = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.CancelNewsButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonBorderEdge19 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.AttachButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.DetachButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.AnimateTimer = new System.Windows.Forms.Timer(this.components);
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            this.LoadTimer = new System.Windows.Forms.Timer(this.components);
            this.PercentsLabel = new System.Windows.Forms.Label();
            this.SpeedLabel = new System.Windows.Forms.Label();
            this.DownloadLabel = new System.Windows.Forms.Label();
            this.LoadLabel = new System.Windows.Forms.Label();
            this.CancelLoadingFilesButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.BodyTextEdit = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
            this.kryptonTextBox1 = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.AllUsersNotifyCheckBox = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.NoNotifyCheckBox = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.AttachmentsGrid = new Infinium.PercentageDataGrid();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AttachmentsGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // kryptonBorderEdge20
            // 
            this.kryptonBorderEdge20.Dock = System.Windows.Forms.DockStyle.Right;
            this.kryptonBorderEdge20.Location = new System.Drawing.Point(1013, 0);
            this.kryptonBorderEdge20.Name = "kryptonBorderEdge20";
            this.kryptonBorderEdge20.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge20.Size = new System.Drawing.Size(1, 444);
            this.kryptonBorderEdge20.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge20.Text = "kryptonBorderEdge20";
            // 
            // kryptonBorderEdge21
            // 
            this.kryptonBorderEdge21.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonBorderEdge21.Location = new System.Drawing.Point(0, 0);
            this.kryptonBorderEdge21.Name = "kryptonBorderEdge21";
            this.kryptonBorderEdge21.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge21.Size = new System.Drawing.Size(1, 444);
            this.kryptonBorderEdge21.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge21.Text = "kryptonBorderEdge21";
            // 
            // kryptonBorderEdge22
            // 
            this.kryptonBorderEdge22.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonBorderEdge22.Location = new System.Drawing.Point(1, 443);
            this.kryptonBorderEdge22.Name = "kryptonBorderEdge22";
            this.kryptonBorderEdge22.Size = new System.Drawing.Size(1012, 1);
            this.kryptonBorderEdge22.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge22.Text = "kryptonBorderEdge22";
            // 
            // kryptonBorderEdge23
            // 
            this.kryptonBorderEdge23.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonBorderEdge23.Location = new System.Drawing.Point(1, 0);
            this.kryptonBorderEdge23.Name = "kryptonBorderEdge23";
            this.kryptonBorderEdge23.Size = new System.Drawing.Size(1012, 1);
            this.kryptonBorderEdge23.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge23.Text = "kryptonBorderEdge23";
            // 
            // OKNewsButton
            // 
            this.OKNewsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.OKNewsButton.Location = new System.Drawing.Point(641, 387);
            this.OKNewsButton.Name = "OKNewsButton";
            this.OKNewsButton.Palette = this.StandardButtonsPalette;
            this.OKNewsButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.OKNewsButton.Size = new System.Drawing.Size(145, 38);
            this.OKNewsButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.OKNewsButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.OKNewsButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.OKNewsButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.OKNewsButton.StateCommon.Content.ShortText.TextV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.OKNewsButton.StateDisabled.Back.Color1 = System.Drawing.Color.LightGray;
            this.OKNewsButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.OKNewsButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.OKNewsButton.TabIndex = 174;
            this.OKNewsButton.TabStop = false;
            this.OKNewsButton.Values.Text = "OK";
            this.OKNewsButton.Click += new System.EventHandler(this.OKDateButton_Click);
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
            // 
            // CancelNewsButton
            // 
            this.CancelNewsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CancelNewsButton.Location = new System.Drawing.Point(804, 387);
            this.CancelNewsButton.Name = "CancelNewsButton";
            this.CancelNewsButton.Palette = this.StandardButtonsPalette;
            this.CancelNewsButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.CancelNewsButton.Size = new System.Drawing.Size(145, 38);
            this.CancelNewsButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelNewsButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelNewsButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.CancelNewsButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.CancelNewsButton.StateCommon.Content.ShortText.TextV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelNewsButton.StateDisabled.Back.Color1 = System.Drawing.Color.LightGray;
            this.CancelNewsButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.CancelNewsButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.CancelNewsButton.TabIndex = 175;
            this.CancelNewsButton.TabStop = false;
            this.CancelNewsButton.Values.Text = "ОТМЕНА";
            this.CancelNewsButton.Click += new System.EventHandler(this.CancelDateButton_Click);
            // 
            // kryptonBorderEdge19
            // 
            this.kryptonBorderEdge19.AutoSize = false;
            this.kryptonBorderEdge19.Location = new System.Drawing.Point(27, 368);
            this.kryptonBorderEdge19.Name = "kryptonBorderEdge19";
            this.kryptonBorderEdge19.Size = new System.Drawing.Size(968, 1);
            this.kryptonBorderEdge19.StateCommon.Color1 = System.Drawing.Color.DarkGray;
            this.kryptonBorderEdge19.Text = "kryptonBorderEdge19";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1012, 38);
            this.panel1.TabIndex = 239;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Light", 23F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 1);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(267, 32);
            this.label2.TabIndex = 240;
            this.label2.Text = "Добавление сообщения";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(23, 42);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(146, 23);
            this.label4.TabIndex = 241;
            this.label4.Text = "Текст сообщения";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(693, 42);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(195, 23);
            this.label5.TabIndex = 268;
            this.label5.Text = "Прикрепленные файлы";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.AttachmentsGrid);
            this.panel2.Location = new System.Drawing.Point(698, 68);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(251, 227);
            this.panel2.TabIndex = 270;
            // 
            // AttachButton
            // 
            this.AttachButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.AttachButton.Location = new System.Drawing.Point(954, 69);
            this.AttachButton.Name = "AttachButton";
            this.AttachButton.OverrideDefault.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(158)))), ((int)(((byte)(0)))));
            this.AttachButton.Palette = this.StandardButtonsPalette;
            this.AttachButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.AttachButton.Size = new System.Drawing.Size(42, 39);
            this.AttachButton.StateCommon.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(158)))), ((int)(((byte)(0)))));
            this.AttachButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.AttachButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.AttachButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.AttachButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.AttachButton.StatePressed.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(158)))), ((int)(((byte)(0)))));
            this.AttachButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.AttachButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachButton.TabIndex = 272;
            this.AttachButton.TabStop = false;
            this.AttachButton.Values.Image = ((System.Drawing.Image)(resources.GetObject("AttachButton.Values.Image")));
            this.AttachButton.Values.Text = "";
            this.AttachButton.Click += new System.EventHandler(this.AttachButton_Click);
            // 
            // DetachButton
            // 
            this.DetachButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.DetachButton.Location = new System.Drawing.Point(954, 114);
            this.DetachButton.Name = "DetachButton";
            this.DetachButton.Palette = this.StandardButtonsPalette;
            this.DetachButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.DetachButton.Size = new System.Drawing.Size(42, 39);
            this.DetachButton.StateCommon.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(150)))), ((int)(((byte)(201)))));
            this.DetachButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.DetachButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.DetachButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.DetachButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DetachButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.DetachButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.DetachButton.TabIndex = 273;
            this.DetachButton.TabStop = false;
            this.DetachButton.Values.Image = ((System.Drawing.Image)(resources.GetObject("DetachButton.Values.Image")));
            this.DetachButton.Values.Text = "";
            this.DetachButton.Click += new System.EventHandler(this.DetachButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // AnimateTimer
            // 
            this.AnimateTimer.Interval = 1;
            this.AnimateTimer.Tick += new System.EventHandler(this.AnimateTimer_Tick);
            // 
            // ProgressBar
            // 
            this.ProgressBar.Location = new System.Drawing.Point(27, 399);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(439, 8);
            this.ProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.ProgressBar.TabIndex = 282;
            this.ProgressBar.Visible = false;
            // 
            // LoadTimer
            // 
            this.LoadTimer.Interval = 400;
            this.LoadTimer.Tick += new System.EventHandler(this.LoadTimer_Tick);
            // 
            // PercentsLabel
            // 
            this.PercentsLabel.AutoSize = true;
            this.PercentsLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.PercentsLabel.ForeColor = System.Drawing.Color.Gray;
            this.PercentsLabel.Location = new System.Drawing.Point(23, 411);
            this.PercentsLabel.Name = "PercentsLabel";
            this.PercentsLabel.Size = new System.Drawing.Size(0, 20);
            this.PercentsLabel.TabIndex = 291;
            // 
            // SpeedLabel
            // 
            this.SpeedLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.SpeedLabel.ForeColor = System.Drawing.Color.Black;
            this.SpeedLabel.Location = new System.Drawing.Point(301, 410);
            this.SpeedLabel.Name = "SpeedLabel";
            this.SpeedLabel.Size = new System.Drawing.Size(165, 20);
            this.SpeedLabel.TabIndex = 292;
            this.SpeedLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // DownloadLabel
            // 
            this.DownloadLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.DownloadLabel.ForeColor = System.Drawing.Color.Black;
            this.DownloadLabel.Location = new System.Drawing.Point(301, 376);
            this.DownloadLabel.Name = "DownloadLabel";
            this.DownloadLabel.Size = new System.Drawing.Size(165, 20);
            this.DownloadLabel.TabIndex = 293;
            this.DownloadLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // LoadLabel
            // 
            this.LoadLabel.AutoSize = true;
            this.LoadLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.LoadLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(91)))), ((int)(((byte)(167)))), ((int)(((byte)(231)))));
            this.LoadLabel.Location = new System.Drawing.Point(24, 375);
            this.LoadLabel.Name = "LoadLabel";
            this.LoadLabel.Size = new System.Drawing.Size(0, 20);
            this.LoadLabel.TabIndex = 294;
            // 
            // CancelLoadingFilesButton
            // 
            this.CancelLoadingFilesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CancelLoadingFilesButton.Location = new System.Drawing.Point(484, 384);
            this.CancelLoadingFilesButton.Name = "CancelLoadingFilesButton";
            this.CancelLoadingFilesButton.OverrideDefault.Back.Color1 = System.Drawing.Color.Transparent;
            this.CancelLoadingFilesButton.Palette = this.StandardButtonsPalette;
            this.CancelLoadingFilesButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.CancelLoadingFilesButton.Size = new System.Drawing.Size(34, 34);
            this.CancelLoadingFilesButton.StateCommon.Back.Color1 = System.Drawing.Color.Transparent;
            this.CancelLoadingFilesButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelLoadingFilesButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelLoadingFilesButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelLoadingFilesButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CancelLoadingFilesButton.StatePressed.Back.Color1 = System.Drawing.Color.Transparent;
            this.CancelLoadingFilesButton.StateTracking.Back.Color1 = System.Drawing.Color.Transparent;
            this.CancelLoadingFilesButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.CancelLoadingFilesButton.StateTracking.Border.Color1 = System.Drawing.Color.Silver;
            this.CancelLoadingFilesButton.StateTracking.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelLoadingFilesButton.TabIndex = 295;
            this.CancelLoadingFilesButton.TabStop = false;
            this.CancelLoadingFilesButton.Values.Image = ((System.Drawing.Image)(resources.GetObject("CancelLoadingFilesButton.Values.Image")));
            this.CancelLoadingFilesButton.Values.Text = "";
            this.CancelLoadingFilesButton.Visible = false;
            this.CancelLoadingFilesButton.Click += new System.EventHandler(this.CancelLoadingFilesButton_Click);
            // 
            // BodyTextEdit
            // 
            this.BodyTextEdit.Location = new System.Drawing.Point(28, 68);
            this.BodyTextEdit.Name = "BodyTextEdit";
            this.BodyTextEdit.Size = new System.Drawing.Size(645, 226);
            this.BodyTextEdit.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.BodyTextEdit.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.BodyTextEdit.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.BodyTextEdit.TabIndex = 308;
            this.BodyTextEdit.Text = "";
            // 
            // kryptonTextBox1
            // 
            this.kryptonTextBox1.Location = new System.Drawing.Point(27, 83);
            this.kryptonTextBox1.Name = "kryptonTextBox1";
            this.kryptonTextBox1.Size = new System.Drawing.Size(922, 31);
            this.kryptonTextBox1.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.kryptonTextBox1.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonTextBox1.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI Semilight", 12F);
            this.kryptonTextBox1.TabIndex = 307;
            // 
            // AllUsersNotifyCheckBox
            // 
            this.AllUsersNotifyCheckBox.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl;
            this.AllUsersNotifyCheckBox.Location = new System.Drawing.Point(28, 302);
            this.AllUsersNotifyCheckBox.Name = "AllUsersNotifyCheckBox";
            this.AllUsersNotifyCheckBox.Size = new System.Drawing.Size(136, 24);
            this.AllUsersNotifyCheckBox.StateCommon.ShortText.Color1 = System.Drawing.Color.Black;
            this.AllUsersNotifyCheckBox.TabIndex = 347;
            this.AllUsersNotifyCheckBox.Text = "Уведомить всех";
            this.AllUsersNotifyCheckBox.Values.Text = "Уведомить всех";
            this.AllUsersNotifyCheckBox.CheckedChanged += new System.EventHandler(this.AllUsersNotifyCheckBox_CheckedChanged);
            // 
            // NoNotifyCheckBox
            // 
            this.NoNotifyCheckBox.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl;
            this.NoNotifyCheckBox.Location = new System.Drawing.Point(175, 302);
            this.NoNotifyCheckBox.Name = "NoNotifyCheckBox";
            this.NoNotifyCheckBox.Size = new System.Drawing.Size(148, 24);
            this.NoNotifyCheckBox.StateCommon.ShortText.Color1 = System.Drawing.Color.Black;
            this.NoNotifyCheckBox.TabIndex = 348;
            this.NoNotifyCheckBox.Text = "Без уведомлений";
            this.NoNotifyCheckBox.Values.Text = "Без уведомлений";
            this.NoNotifyCheckBox.CheckedChanged += new System.EventHandler(this.NoNotifyCheckBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label1.ForeColor = System.Drawing.Color.Gray;
            this.label1.Location = new System.Drawing.Point(29, 327);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(556, 17);
            this.label1.TabIndex = 349;
            this.label1.Text = "* Отметьте \"Уведомить всех\", если нужно, чтобы все пользователи получили уведомле" +
    "ние.";
            // 
            // AttachmentsGrid
            // 
            this.AttachmentsGrid.AllowUserToAddRows = false;
            this.AttachmentsGrid.AllowUserToDeleteRows = false;
            this.AttachmentsGrid.AllowUserToResizeColumns = false;
            this.AttachmentsGrid.AllowUserToResizeRows = false;
            this.AttachmentsGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.AttachmentsGrid.BackText = "Нет данных";
            this.AttachmentsGrid.ColumnHeadersHeight = 40;
            this.AttachmentsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.AttachmentsGrid.ColumnHeadersVisible = false;
            this.AttachmentsGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.AttachmentsGrid.Location = new System.Drawing.Point(0, 0);
            this.AttachmentsGrid.Name = "AttachmentsGrid";
            this.AttachmentsGrid.PercentLineWidth = 0;
            this.AttachmentsGrid.RowHeadersVisible = false;
            this.AttachmentsGrid.RowTemplate.Height = 30;
            this.AttachmentsGrid.SelectedColorStyle = Infinium.PercentageDataGrid.ColorStyle.Orange;
            this.AttachmentsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.AttachmentsGrid.Size = new System.Drawing.Size(249, 225);
            this.AttachmentsGrid.StandardStyle = false;
            this.AttachmentsGrid.StateCommon.Background.Color1 = System.Drawing.Color.White;
            this.AttachmentsGrid.StateCommon.Background.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachmentsGrid.StateCommon.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridBackgroundList;
            this.AttachmentsGrid.StateCommon.DataCell.Back.Color1 = System.Drawing.Color.White;
            this.AttachmentsGrid.StateCommon.DataCell.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(210)))), ((int)(((byte)(210)))), ((int)(((byte)(210)))));
            this.AttachmentsGrid.StateCommon.DataCell.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachmentsGrid.StateCommon.DataCell.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.AttachmentsGrid.StateCommon.DataCell.Content.Color1 = System.Drawing.Color.Black;
            this.AttachmentsGrid.StateCommon.DataCell.Content.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.AttachmentsGrid.StateCommon.HeaderColumn.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(150)))), ((int)(((byte)(150)))), ((int)(((byte)(150)))));
            this.AttachmentsGrid.StateCommon.HeaderColumn.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachmentsGrid.StateCommon.HeaderColumn.Border.Color1 = System.Drawing.Color.Black;
            this.AttachmentsGrid.StateCommon.HeaderColumn.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachmentsGrid.StateCommon.HeaderColumn.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.AttachmentsGrid.StateCommon.HeaderColumn.Content.Color1 = System.Drawing.Color.White;
            this.AttachmentsGrid.StateCommon.HeaderColumn.Content.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.AttachmentsGrid.StateSelected.DataCell.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(158)))), ((int)(((byte)(0)))));
            this.AttachmentsGrid.StateSelected.DataCell.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.AttachmentsGrid.StateSelected.DataCell.Content.Color1 = System.Drawing.Color.White;
            this.AttachmentsGrid.TabIndex = 309;
            this.AttachmentsGrid.TabStop = false;
            this.AttachmentsGrid.UseCustomBackColor = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label3.ForeColor = System.Drawing.Color.Gray;
            this.label3.Location = new System.Drawing.Point(38, 345);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(734, 17);
            this.label3.TabIndex = 355;
            this.label3.Text = "Иначе уведомление получат только участники проекта. Отметьте \"Без уведомлений\", е" +
    "сли не хотите никого уведомлять.";
            // 
            // AddProjectNewsForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1014, 444);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.NoNotifyCheckBox);
            this.Controls.Add(this.AllUsersNotifyCheckBox);
            this.Controls.Add(this.BodyTextEdit);
            this.Controls.Add(this.CancelLoadingFilesButton);
            this.Controls.Add(this.LoadLabel);
            this.Controls.Add(this.DownloadLabel);
            this.Controls.Add(this.SpeedLabel);
            this.Controls.Add(this.PercentsLabel);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.DetachButton);
            this.Controls.Add(this.AttachButton);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.kryptonBorderEdge19);
            this.Controls.Add(this.CancelNewsButton);
            this.Controls.Add(this.OKNewsButton);
            this.Controls.Add(this.kryptonBorderEdge23);
            this.Controls.Add(this.kryptonBorderEdge22);
            this.Controls.Add(this.kryptonBorderEdge21);
            this.Controls.Add(this.kryptonBorderEdge20);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AddProjectNewsForm";
            this.Opacity = 0D;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Infinium";
            this.Shown += new System.EventHandler(this.AddNewsForm_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.AttachmentsGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private KryptonBorderEdge kryptonBorderEdge20;
        private KryptonBorderEdge kryptonBorderEdge21;
        private KryptonBorderEdge kryptonBorderEdge22;
        private KryptonBorderEdge kryptonBorderEdge23;
        private KryptonButton OKNewsButton;
        private KryptonButton CancelNewsButton;
        private KryptonBorderEdge kryptonBorderEdge19;
        private Panel panel1;
        private Label label2;
        private Label label4;
        private KryptonPalette StandardButtonsPalette;
        private Label label5;
        private Panel panel2;
        private KryptonButton AttachButton;
        private KryptonButton DetachButton;
        private OpenFileDialog openFileDialog1;
        private Timer AnimateTimer;
        private ProgressBar ProgressBar;
        private Timer LoadTimer;
        private Label PercentsLabel;
        private Label SpeedLabel;
        private Label DownloadLabel;
        private Label LoadLabel;
        private KryptonButton CancelLoadingFilesButton;
        private KryptonRichTextBox BodyTextEdit;
        private KryptonTextBox kryptonTextBox1;
        private PercentageDataGrid AttachmentsGrid;
        private KryptonCheckBox AllUsersNotifyCheckBox;
        private KryptonCheckBox NoNotifyCheckBox;
        private Label label1;
        private Label label3;
    }
}