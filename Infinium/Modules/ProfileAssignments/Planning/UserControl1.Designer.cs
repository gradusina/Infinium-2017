
namespace Infinium.Modules.ProfileAssignments.Planning
{
    partial class UserControl1
    {
        /// <summary> 
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary> 
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lbTotalProduceCount = new System.Windows.Forms.Label();
            this.panel7 = new System.Windows.Forms.Panel();
            this.dgvOrders = new System.Windows.Forms.DataGridView();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lbSheetThickness = new System.Windows.Forms.Label();
            this.btnCollapsePanel = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.panel7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOrders)).BeginInit();
            this.pnlTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbTotalProduceCount
            // 
            this.lbTotalProduceCount.AutoSize = true;
            this.lbTotalProduceCount.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbTotalProduceCount.Location = new System.Drawing.Point(3, 289);
            this.lbTotalProduceCount.Name = "lbTotalProduceCount";
            this.lbTotalProduceCount.Size = new System.Drawing.Size(245, 19);
            this.lbTotalProduceCount.TabIndex = 452;
            this.lbTotalProduceCount.Text = "Общее кол-во штук к производству: ";
            // 
            // panel7
            // 
            this.panel7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel7.Controls.Add(this.dgvOrders);
            this.panel7.Location = new System.Drawing.Point(7, 27);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(571, 259);
            this.panel7.TabIndex = 453;
            // 
            // dgvOrders
            // 
            this.dgvOrders.AllowDrop = true;
            this.dgvOrders.AllowUserToAddRows = false;
            this.dgvOrders.AllowUserToDeleteRows = false;
            this.dgvOrders.AllowUserToOrderColumns = true;
            this.dgvOrders.AllowUserToResizeRows = false;
            this.dgvOrders.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvOrders.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvOrders.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvOrders.ColumnHeadersHeight = 30;
            this.dgvOrders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvOrders.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvOrders.Location = new System.Drawing.Point(0, 0);
            this.dgvOrders.MultiSelect = false;
            this.dgvOrders.Name = "dgvOrders";
            this.dgvOrders.RowHeadersVisible = false;
            this.dgvOrders.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dgvOrders.RowTemplate.Height = 25;
            this.dgvOrders.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvOrders.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvOrders.Size = new System.Drawing.Size(569, 257);
            this.dgvOrders.TabIndex = 55;
            this.dgvOrders.VirtualMode = true;
            // 
            // pnlTop
            // 
            this.pnlTop.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlTop.Controls.Add(this.lbSheetThickness);
            this.pnlTop.Controls.Add(this.btnCollapsePanel);
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(581, 25);
            this.pnlTop.TabIndex = 489;
            // 
            // lbSheetThickness
            // 
            this.lbSheetThickness.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbSheetThickness.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbSheetThickness.ForeColor = System.Drawing.Color.Black;
            this.lbSheetThickness.Location = new System.Drawing.Point(30, 0);
            this.lbSheetThickness.Name = "lbSheetThickness";
            this.lbSheetThickness.Size = new System.Drawing.Size(539, 25);
            this.lbSheetThickness.TabIndex = 485;
            this.lbSheetThickness.Text = "Толщина листа:";
            this.lbSheetThickness.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnCollapsePanel
            // 
            this.btnCollapsePanel.Location = new System.Drawing.Point(0, 0);
            this.btnCollapsePanel.Name = "btnCollapsePanel";
            this.btnCollapsePanel.Orientation = ComponentFactory.Krypton.Toolkit.VisualOrientation.Right;
            this.btnCollapsePanel.OverrideDefault.Back.Color1 = System.Drawing.Color.White;
            this.btnCollapsePanel.OverrideDefault.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.btnCollapsePanel.OverrideDefault.Border.Color1 = System.Drawing.Color.White;
            this.btnCollapsePanel.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnCollapsePanel.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            this.btnCollapsePanel.Size = new System.Drawing.Size(24, 22);
            this.btnCollapsePanel.StateCommon.Back.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.btnCollapsePanel.StateCommon.Border.Color1 = System.Drawing.Color.White;
            this.btnCollapsePanel.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnCollapsePanel.StatePressed.Border.Color1 = System.Drawing.Color.Black;
            this.btnCollapsePanel.StatePressed.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.btnCollapsePanel.StatePressed.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnCollapsePanel.StateTracking.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnCollapsePanel.StateTracking.Border.ColorStyle = ComponentFactory.Krypton.Toolkit.PaletteColorStyle.Solid;
            this.btnCollapsePanel.StateTracking.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnCollapsePanel.TabIndex = 475;
            this.btnCollapsePanel.TabStop = false;
            this.btnCollapsePanel.Tag = "0";
            this.btnCollapsePanel.Values.Image = global::Infinium.Properties.Resources.Collapsed;
            this.btnCollapsePanel.Values.Text = "";
            this.btnCollapsePanel.Click += new System.EventHandler(this.btnCollapsePanel_Click);
            // 
            // UserControl1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.pnlTop);
            this.Controls.Add(this.panel7);
            this.Controls.Add(this.lbTotalProduceCount);
            this.DoubleBuffered = true;
            this.Name = "UserControl1";
            this.Size = new System.Drawing.Size(581, 312);
            this.panel7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvOrders)).EndInit();
            this.pnlTop.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lbTotalProduceCount;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.DataGridView dgvOrders;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lbSheetThickness;
        private ComponentFactory.Krypton.Toolkit.KryptonButton btnCollapsePanel;
    }
}
