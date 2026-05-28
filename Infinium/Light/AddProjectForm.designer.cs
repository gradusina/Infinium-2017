using System.ComponentModel;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;

namespace Infinium
{
    partial class AddProjectForm
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
            this.OKNewsButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.StandardButtonsPalette = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.CancelNewsButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonBorderEdge19 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.AnimateTimer = new System.Windows.Forms.Timer(this.components);
            this.LoadTimer = new System.Windows.Forms.Timer(this.components);
            this.PercentsLabel = new System.Windows.Forms.Label();
            this.LoadLabel = new System.Windows.Forms.Label();
            this.ProjectDescriptionRichTextBox = new ComponentFactory.Krypton.Toolkit.KryptonRichTextBox();
            this.ProjectNameTextEdit = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.kryptonTextBox1 = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.MembersTree = new System.Windows.Forms.TreeView();
            this.CreateLabel = new System.Windows.Forms.Label();
            this.SymbolCountText = new System.Windows.Forms.Label();
            this.AllUsersNotifyCheckBox = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.ProposCheckBox = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.DescCountLabel = new System.Windows.Forms.Label();
            this.tbArticle = new ComponentFactory.Krypton.Toolkit.KryptonTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonBorderEdge20
            // 
            this.kryptonBorderEdge20.Dock = System.Windows.Forms.DockStyle.Right;
            this.kryptonBorderEdge20.Location = new System.Drawing.Point(1013, 0);
            this.kryptonBorderEdge20.Name = "kryptonBorderEdge20";
            this.kryptonBorderEdge20.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge20.Size = new System.Drawing.Size(1, 477);
            this.kryptonBorderEdge20.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge20.Text = "kryptonBorderEdge20";
            // 
            // kryptonBorderEdge21
            // 
            this.kryptonBorderEdge21.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonBorderEdge21.Location = new System.Drawing.Point(0, 0);
            this.kryptonBorderEdge21.Name = "kryptonBorderEdge21";
            this.kryptonBorderEdge21.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge21.Size = new System.Drawing.Size(1, 477);
            this.kryptonBorderEdge21.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge21.Text = "kryptonBorderEdge21";
            // 
            // kryptonBorderEdge22
            // 
            this.kryptonBorderEdge22.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonBorderEdge22.Location = new System.Drawing.Point(1, 476);
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
            this.OKNewsButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.OKNewsButton.Location = new System.Drawing.Point(701, 433);
            this.OKNewsButton.Name = "OKNewsButton";
            this.OKNewsButton.Palette = this.StandardButtonsPalette;
            this.OKNewsButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.OKNewsButton.Size = new System.Drawing.Size(136, 32);
            this.OKNewsButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.OKNewsButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.OKNewsButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.OKNewsButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.OKNewsButton.StateCommon.Content.ShortText.TextV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.OKNewsButton.StateDisabled.Back.Color1 = System.Drawing.Color.LightGray;
            this.OKNewsButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.OKNewsButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.OKNewsButton.TabIndex = 174;
            this.OKNewsButton.TabStop = false;
            this.OKNewsButton.Values.Text = "OK";
            this.OKNewsButton.Click += new System.EventHandler(this.OKNewsButton_Click);
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
            // CancelNewsButton
            // 
            this.CancelNewsButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.CancelNewsButton.Location = new System.Drawing.Point(851, 433);
            this.CancelNewsButton.Name = "CancelNewsButton";
            this.CancelNewsButton.Palette = this.StandardButtonsPalette;
            this.CancelNewsButton.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.CancelNewsButton.Size = new System.Drawing.Size(136, 32);
            this.CancelNewsButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelNewsButton.StateCommon.Content.Image.ImageH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelNewsButton.StateCommon.Content.Image.ImageV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.CancelNewsButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.CancelNewsButton.StateCommon.Content.ShortText.TextV = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.CancelNewsButton.StateDisabled.Back.Color1 = System.Drawing.Color.LightGray;
            this.CancelNewsButton.StateTracking.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(164)))), ((int)(((byte)(61)))));
            this.CancelNewsButton.StateTracking.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.CancelNewsButton.TabIndex = 175;
            this.CancelNewsButton.TabStop = false;
            this.CancelNewsButton.Values.Text = "ОТМЕНА";
            this.CancelNewsButton.Click += new System.EventHandler(this.CancelNewsButton_Click);
            // 
            // kryptonBorderEdge19
            // 
            this.kryptonBorderEdge19.AutoSize = false;
            this.kryptonBorderEdge19.Location = new System.Drawing.Point(27, 421);
            this.kryptonBorderEdge19.Name = "kryptonBorderEdge19";
            this.kryptonBorderEdge19.Size = new System.Drawing.Size(960, 1);
            this.kryptonBorderEdge19.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
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
            this.label2.Location = new System.Drawing.Point(3, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 31);
            this.label2.TabIndex = 240;
            this.label2.Text = "Новый проект";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(24, 51);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(158, 23);
            this.label3.TabIndex = 240;
            this.label3.Text = "Название проекта:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(24, 127);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 23);
            this.label4.TabIndex = 241;
            this.label4.Text = "Описание:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // AnimateTimer
            // 
            this.AnimateTimer.Interval = 1;
            this.AnimateTimer.Tick += new System.EventHandler(this.AnimateTimer_Tick);
            // 
            // LoadTimer
            // 
            this.LoadTimer.Interval = 400;
            // 
            // PercentsLabel
            // 
            this.PercentsLabel.AutoSize = true;
            this.PercentsLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.PercentsLabel.ForeColor = System.Drawing.Color.Gray;
            this.PercentsLabel.Location = new System.Drawing.Point(23, 433);
            this.PercentsLabel.Name = "PercentsLabel";
            this.PercentsLabel.Size = new System.Drawing.Size(0, 20);
            this.PercentsLabel.TabIndex = 291;
            // 
            // LoadLabel
            // 
            this.LoadLabel.AutoSize = true;
            this.LoadLabel.Font = new System.Drawing.Font("Segoe UI Semilight", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.LoadLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(91)))), ((int)(((byte)(167)))), ((int)(((byte)(231)))));
            this.LoadLabel.Location = new System.Drawing.Point(24, 399);
            this.LoadLabel.Name = "LoadLabel";
            this.LoadLabel.Size = new System.Drawing.Size(0, 20);
            this.LoadLabel.TabIndex = 294;
            // 
            // ProjectDescriptionRichTextBox
            // 
            this.ProjectDescriptionRichTextBox.Location = new System.Drawing.Point(28, 152);
            this.ProjectDescriptionRichTextBox.MaxLength = 500;
            this.ProjectDescriptionRichTextBox.Name = "ProjectDescriptionRichTextBox";
            this.ProjectDescriptionRichTextBox.Size = new System.Drawing.Size(598, 236);
            this.ProjectDescriptionRichTextBox.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
            this.ProjectDescriptionRichTextBox.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.ProjectDescriptionRichTextBox.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI Semilight", 11.25F);
            this.ProjectDescriptionRichTextBox.StateCommon.Content.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Inherit;
            this.ProjectDescriptionRichTextBox.TabIndex = 308;
            this.ProjectDescriptionRichTextBox.Text = "";
            this.ProjectDescriptionRichTextBox.TextChanged += new System.EventHandler(this.ProjectDescriptionRichTextBox_TextChanged);
            // 
            // ProjectNameTextEdit
            // 
            this.ProjectNameTextEdit.Location = new System.Drawing.Point(27, 76);
            this.ProjectNameTextEdit.MaxLength = 60;
            this.ProjectNameTextEdit.Name = "ProjectNameTextEdit";
            this.ProjectNameTextEdit.Size = new System.Drawing.Size(601, 29);
            this.ProjectNameTextEdit.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
            this.ProjectNameTextEdit.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.ProjectNameTextEdit.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI Semilight", 12F);
            this.ProjectNameTextEdit.StateCommon.Content.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Inherit;
            this.ProjectNameTextEdit.TabIndex = 307;
            this.ProjectNameTextEdit.TextChanged += new System.EventHandler(this.ProjectNameTextEdit_TextChanged);
            // 
            // kryptonTextBox1
            // 
            this.kryptonTextBox1.Location = new System.Drawing.Point(27, 83);
            this.kryptonTextBox1.Name = "kryptonTextBox1";
            this.kryptonTextBox1.Size = new System.Drawing.Size(922, 29);
            this.kryptonTextBox1.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.kryptonTextBox1.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonTextBox1.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI Semilight", 12F);
            this.kryptonTextBox1.StateCommon.Content.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Inherit;
            this.kryptonTextBox1.TabIndex = 307;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(642, 127);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(164, 23);
            this.label5.TabIndex = 316;
            this.label5.Text = "Участники проекта:";
            // 
            // MembersTree
            // 
            this.MembersTree.CheckBoxes = true;
            this.MembersTree.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.MembersTree.Location = new System.Drawing.Point(646, 152);
            this.MembersTree.Name = "MembersTree";
            this.MembersTree.ShowLines = false;
            this.MembersTree.Size = new System.Drawing.Size(341, 201);
            this.MembersTree.TabIndex = 328;
            this.MembersTree.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.MembersTree_AfterCheck);
            // 
            // CreateLabel
            // 
            this.CreateLabel.AutoSize = true;
            this.CreateLabel.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.CreateLabel.ForeColor = System.Drawing.Color.DimGray;
            this.CreateLabel.Location = new System.Drawing.Point(24, 428);
            this.CreateLabel.Name = "CreateLabel";
            this.CreateLabel.Size = new System.Drawing.Size(170, 23);
            this.CreateLabel.TabIndex = 334;
            this.CreateLabel.Text = "Создание проекта....";
            this.CreateLabel.Visible = false;
            // 
            // SymbolCountText
            // 
            this.SymbolCountText.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.SymbolCountText.ForeColor = System.Drawing.Color.Gray;
            this.SymbolCountText.Location = new System.Drawing.Point(462, 108);
            this.SymbolCountText.Name = "SymbolCountText";
            this.SymbolCountText.Size = new System.Drawing.Size(166, 23);
            this.SymbolCountText.TabIndex = 340;
            this.SymbolCountText.Text = "60 символов осталось";
            this.SymbolCountText.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // AllUsersNotifyCheckBox
            // 
            this.AllUsersNotifyCheckBox.Location = new System.Drawing.Point(28, 393);
            this.AllUsersNotifyCheckBox.Name = "AllUsersNotifyCheckBox";
            this.AllUsersNotifyCheckBox.Size = new System.Drawing.Size(188, 20);
            this.AllUsersNotifyCheckBox.StateCommon.ShortText.Color1 = System.Drawing.Color.Black;
            this.AllUsersNotifyCheckBox.TabIndex = 346;
            this.toolTip1.SetToolTip(this.AllUsersNotifyCheckBox, "Отметьте, если хотите, чтобы все сотрудники получили уведомление \r\nо создании это" +
        "го проекта. Иначе, уведомление получат только участники проекта");
            this.AllUsersNotifyCheckBox.Values.Text = "Уведомить всех сотрудников";
            // 
            // ProposCheckBox
            // 
            this.ProposCheckBox.Location = new System.Drawing.Point(273, 393);
            this.ProposCheckBox.Name = "ProposCheckBox";
            this.ProposCheckBox.Size = new System.Drawing.Size(103, 20);
            this.ProposCheckBox.StateCommon.ShortText.Color1 = System.Drawing.Color.Black;
            this.ProposCheckBox.TabIndex = 358;
            this.toolTip1.SetToolTip(this.ProposCheckBox, "Отметьте, если хотите, чтобы все сотрудники получили уведомление \r\nо создании это" +
        "го проекта. Иначе, уведомление получат только участники проекта");
            this.ProposCheckBox.Values.Text = "Предложение";
            // 
            // DescCountLabel
            // 
            this.DescCountLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.DescCountLabel.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.DescCountLabel.ForeColor = System.Drawing.Color.Gray;
            this.DescCountLabel.Location = new System.Drawing.Point(462, 389);
            this.DescCountLabel.Name = "DescCountLabel";
            this.DescCountLabel.Size = new System.Drawing.Size(166, 23);
            this.DescCountLabel.TabIndex = 352;
            this.DescCountLabel.Text = "500 символов осталось";
            this.DescCountLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // tbArticle
            // 
            this.tbArticle.Location = new System.Drawing.Point(646, 76);
            this.tbArticle.MaxLength = 40;
            this.tbArticle.Name = "tbArticle";
            this.tbArticle.Size = new System.Drawing.Size(341, 29);
            this.tbArticle.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
            this.tbArticle.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.tbArticle.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI Semilight", 12F);
            this.tbArticle.StateCommon.Content.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Inherit;
            this.tbArticle.TabIndex = 365;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(643, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(145, 23);
            this.label1.TabIndex = 364;
            this.label1.Text = "Артикул проекта:";
            // 
            // panel5
            // 
            this.panel5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel5.Controls.Add(this.button1);
            this.panel5.Controls.Add(this.textBox1);
            this.panel5.Location = new System.Drawing.Point(646, 359);
            this.panel5.Name = "panel5";
            this.panel5.Padding = new System.Windows.Forms.Padding(3);
            this.panel5.Size = new System.Drawing.Size(341, 29);
            this.panel5.TabIndex = 371;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(266, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 329;
            this.button1.Text = "Найти";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox1.Location = new System.Drawing.Point(1, 7);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(259, 22);
            this.textBox1.TabIndex = 328;
            // 
            // AddProjectForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1014, 477);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.tbArticle);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProposCheckBox);
            this.Controls.Add(this.DescCountLabel);
            this.Controls.Add(this.AllUsersNotifyCheckBox);
            this.Controls.Add(this.SymbolCountText);
            this.Controls.Add(this.CreateLabel);
            this.Controls.Add(this.MembersTree);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.ProjectNameTextEdit);
            this.Controls.Add(this.ProjectDescriptionRichTextBox);
            this.Controls.Add(this.LoadLabel);
            this.Controls.Add(this.PercentsLabel);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.kryptonBorderEdge19);
            this.Controls.Add(this.CancelNewsButton);
            this.Controls.Add(this.OKNewsButton);
            this.Controls.Add(this.kryptonBorderEdge23);
            this.Controls.Add(this.kryptonBorderEdge22);
            this.Controls.Add(this.kryptonBorderEdge21);
            this.Controls.Add(this.kryptonBorderEdge20);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AddProjectForm";
            this.Opacity = 0D;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Infinium";
            this.Shown += new System.EventHandler(this.AddNewsForm_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
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
        private Label label3;
        private Label label4;
        private KryptonPalette StandardButtonsPalette;
        private OpenFileDialog openFileDialog1;
        private Timer AnimateTimer;
        private Timer LoadTimer;
        private Label PercentsLabel;
        private Label LoadLabel;
        private KryptonRichTextBox ProjectDescriptionRichTextBox;
        private KryptonTextBox ProjectNameTextEdit;
        private KryptonTextBox kryptonTextBox1;
        private Label label5;
        private TreeView MembersTree;
        private Label CreateLabel;
        private Label SymbolCountText;
        private KryptonCheckBox AllUsersNotifyCheckBox;
        private ToolTip toolTip1;
        private Label DescCountLabel;
        private KryptonCheckBox ProposCheckBox;
        private KryptonTextBox tbArticle;
        private Label label1;
        private Panel panel5;
        private Button button1;
        private TextBox textBox1;
    }
}