﻿using System.ComponentModel;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;
using ToolTipController = DevExpress.Utils.ToolTipController;

namespace Infinium
{
    partial class LightMessageBoxForm
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
            this.toolTipController1 = new DevExpress.Utils.ToolTipController(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.kryptonBorderEdge1 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.label2 = new System.Windows.Forms.Label();
            this.kryptonBorderEdge2 = new ComponentFactory.Krypton.Toolkit.KryptonBorderEdge();
            this.CancelMessageButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.OKMessageButton = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonBorderEdge20
            // 
            this.kryptonBorderEdge20.Dock = System.Windows.Forms.DockStyle.Right;
            this.kryptonBorderEdge20.Location = new System.Drawing.Point(444, 0);
            this.kryptonBorderEdge20.Name = "kryptonBorderEdge20";
            this.kryptonBorderEdge20.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge20.Size = new System.Drawing.Size(1, 209);
            this.kryptonBorderEdge20.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge20.Text = "kryptonBorderEdge20";
            // 
            // kryptonBorderEdge21
            // 
            this.kryptonBorderEdge21.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonBorderEdge21.Location = new System.Drawing.Point(0, 0);
            this.kryptonBorderEdge21.Name = "kryptonBorderEdge21";
            this.kryptonBorderEdge21.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge21.Size = new System.Drawing.Size(1, 209);
            this.kryptonBorderEdge21.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge21.Text = "kryptonBorderEdge21";
            // 
            // kryptonBorderEdge22
            // 
            this.kryptonBorderEdge22.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonBorderEdge22.Location = new System.Drawing.Point(1, 208);
            this.kryptonBorderEdge22.Name = "kryptonBorderEdge22";
            this.kryptonBorderEdge22.Size = new System.Drawing.Size(443, 1);
            this.kryptonBorderEdge22.StateCommon.Color1 = System.Drawing.Color.Black;
            this.kryptonBorderEdge22.Text = "kryptonBorderEdge22";
            // 
            // kryptonBorderEdge23
            // 
            this.kryptonBorderEdge23.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonBorderEdge23.Location = new System.Drawing.Point(1, 0);
            this.kryptonBorderEdge23.Name = "kryptonBorderEdge23";
            this.kryptonBorderEdge23.Size = new System.Drawing.Size(443, 1);
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
            // toolTipController1
            // 
            this.toolTipController1.Appearance.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolTipController1.Appearance.Options.UseFont = true;
            this.toolTipController1.Appearance.Options.UseTextOptions = true;
            this.toolTipController1.AutoPopDelay = 50000;
            this.toolTipController1.ToolTipStyle = DevExpress.Utils.ToolTipStyle.Windows7;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Font = new System.Drawing.Font("Segoe UI", 17.81F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(15, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(415, 107);
            this.label1.TabIndex = 225;
            this.label1.Text = "Подтвердите удаление сообщения. Сообщения отправленные из группы ";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.panel1.Controls.Add(this.kryptonBorderEdge1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(864, 43);
            this.panel1.TabIndex = 239;
            // 
            // kryptonBorderEdge1
            // 
            this.kryptonBorderEdge1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonBorderEdge1.Location = new System.Drawing.Point(0, 42);
            this.kryptonBorderEdge1.Name = "kryptonBorderEdge1";
            this.kryptonBorderEdge1.Size = new System.Drawing.Size(864, 1);
            this.kryptonBorderEdge1.StateCommon.Color1 = System.Drawing.Color.Gray;
            this.kryptonBorderEdge1.Text = "kryptonBorderEdge1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 23F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(3, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(232, 31);
            this.label2.TabIndex = 240;
            this.label2.Text = "Добавление новости";
            // 
            // kryptonBorderEdge2
            // 
            this.kryptonBorderEdge2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonBorderEdge2.AutoSize = false;
            this.kryptonBorderEdge2.Location = new System.Drawing.Point(-408, 77);
            this.kryptonBorderEdge2.Name = "kryptonBorderEdge2";
            this.kryptonBorderEdge2.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.kryptonBorderEdge2.Size = new System.Drawing.Size(4, 0);
            this.kryptonBorderEdge2.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(150)))), ((int)(((byte)(201)))));
            this.kryptonBorderEdge2.StateCommon.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.kryptonBorderEdge2.Text = "kryptonBorderEdge2";
            // 
            // CancelMessageButton
            // 
            this.CancelMessageButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.CancelMessageButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelMessageButton.Location = new System.Drawing.Point(238, 159);
            this.CancelMessageButton.Name = "CancelMessageButton";
            this.CancelMessageButton.OverrideDefault.Back.Color1 = System.Drawing.Color.Silver;
            this.CancelMessageButton.OverrideDefault.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.False;
            this.CancelMessageButton.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelMessageButton.OverrideDefault.Border.Rounding = 0;
            this.CancelMessageButton.Size = new System.Drawing.Size(108, 34);
            this.CancelMessageButton.StateCommon.Back.Color1 = System.Drawing.Color.Silver;
            this.CancelMessageButton.StateCommon.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.CancelMessageButton.StateCommon.Border.Color1 = System.Drawing.Color.White;
            this.CancelMessageButton.StateCommon.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.False;
            this.CancelMessageButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelMessageButton.StateCommon.Border.Rounding = 0;
            this.CancelMessageButton.StateCommon.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.CancelMessageButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.CancelMessageButton.StateTracking.Border.Color1 = System.Drawing.Color.White;
            this.CancelMessageButton.StateTracking.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.CancelMessageButton.StateTracking.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.CancelMessageButton.StateTracking.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.CancelMessageButton.StateTracking.Border.Rounding = 0;
            this.CancelMessageButton.TabIndex = 472;
            this.CancelMessageButton.Values.Text = "Отмена";
            this.CancelMessageButton.Click += new System.EventHandler(this.CancelMessageButton_Click);
            // 
            // OKMessageButton
            // 
            this.OKMessageButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.OKMessageButton.Location = new System.Drawing.Point(99, 159);
            this.OKMessageButton.Name = "OKMessageButton";
            this.OKMessageButton.OverrideDefault.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(184)))), ((int)(((byte)(238)))));
            this.OKMessageButton.OverrideDefault.Border.Color1 = System.Drawing.Color.White;
            this.OKMessageButton.OverrideDefault.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.OKMessageButton.OverrideDefault.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.False;
            this.OKMessageButton.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.OKMessageButton.OverrideDefault.Border.Rounding = 0;
            this.OKMessageButton.Size = new System.Drawing.Size(108, 34);
            this.OKMessageButton.StateCommon.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(184)))), ((int)(((byte)(238)))));
            this.OKMessageButton.StateCommon.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.OKMessageButton.StateCommon.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.False;
            this.OKMessageButton.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.OKMessageButton.StateCommon.Border.Rounding = 0;
            this.OKMessageButton.StateCommon.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.OKMessageButton.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.OKMessageButton.StateTracking.Border.Color1 = System.Drawing.Color.White;
            this.OKMessageButton.StateTracking.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.OKMessageButton.StateTracking.Border.Draw = ComponentFactory.Krypton.Toolkit.InheritBool.True;
            this.OKMessageButton.StateTracking.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.OKMessageButton.StateTracking.Border.Rounding = 0;
            this.OKMessageButton.TabIndex = 471;
            this.OKMessageButton.Values.Text = "ОК";
            this.OKMessageButton.Click += new System.EventHandler(this.OKMessageButton_Click);
            // 
            // LightMessageBoxForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.CancelButton = this.CancelMessageButton;
            this.ClientSize = new System.Drawing.Size(445, 209);
            this.Controls.Add(this.CancelMessageButton);
            this.Controls.Add(this.OKMessageButton);
            this.Controls.Add(this.kryptonBorderEdge2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.kryptonBorderEdge19);
            this.Controls.Add(this.kryptonBorderEdge23);
            this.Controls.Add(this.kryptonBorderEdge22);
            this.Controls.Add(this.kryptonBorderEdge21);
            this.Controls.Add(this.kryptonBorderEdge20);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "LightMessageBoxForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Infinium";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.LightMessageBoxForm_KeyDown);
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
        private ToolTipController toolTipController1;
        private Label label1;
        private Panel panel1;
        private Label label2;
        private KryptonBorderEdge kryptonBorderEdge2;
        private KryptonPalette StandardButtonsPalette;
        private KryptonButton CancelMessageButton;
        private KryptonButton OKMessageButton;
        private KryptonBorderEdge kryptonBorderEdge1;
    }
}