namespace Infinium.Catalog.UserControls
{
    partial class MachineControl
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbName = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgvForAgreed = new System.Windows.Forms.DataGridView();
            this.panel6 = new System.Windows.Forms.Panel();
            this.dgvOnProduction = new System.Windows.Forms.DataGridView();
            this.panel5 = new System.Windows.Forms.Panel();
            this.dgvAgreed = new System.Windows.Forms.DataGridView();
            this.panel4 = new System.Windows.Forms.Panel();
            this.dgvCalculations = new System.Windows.Forms.DataGridView();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lbSumRank = new System.Windows.Forms.Label();
            this.dgvNotConfirmed = new System.Windows.Forms.DataGridView();
            this.panel7 = new System.Windows.Forms.Panel();
            this.panel8 = new System.Windows.Forms.Panel();
            this.dgvInProduction = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvForAgreed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOnProduction)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAgreed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalculations)).BeginInit();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNotConfirmed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvInProduction)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lbName);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(307, 40);
            this.panel1.TabIndex = 0;
            // 
            // lbName
            // 
            this.lbName.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lbName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbName.Font = new System.Drawing.Font("Microsoft Sans Serif", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbName.Location = new System.Drawing.Point(0, 0);
            this.lbName.Name = "lbName";
            this.lbName.Size = new System.Drawing.Size(307, 40);
            this.lbName.TabIndex = 0;
            this.lbName.Text = "Станок 1";
            this.lbName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel2
            // 
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add(this.dgvInProduction);
            this.panel2.Controls.Add(this.panel8);
            this.panel2.Controls.Add(this.dgvOnProduction);
            this.panel2.Controls.Add(this.panel7);
            this.panel2.Controls.Add(this.dgvAgreed);
            this.panel2.Controls.Add(this.panel6);
            this.panel2.Controls.Add(this.dgvForAgreed);
            this.panel2.Controls.Add(this.panel5);
            this.panel2.Controls.Add(this.dgvNotConfirmed);
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Controls.Add(this.dgvCalculations);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 77);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(307, 692);
            this.panel2.TabIndex = 1;
            this.panel2.Scroll += new System.Windows.Forms.ScrollEventHandler(this.panel2_Scroll);
            // 
            // dgvForAgreed
            // 
            this.dgvForAgreed.AllowUserToAddRows = false;
            this.dgvForAgreed.AllowUserToDeleteRows = false;
            this.dgvForAgreed.AllowUserToResizeRows = false;
            this.dgvForAgreed.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvForAgreed.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvForAgreed.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvForAgreed.ColumnHeadersVisible = false;
            this.dgvForAgreed.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvForAgreed.Location = new System.Drawing.Point(0, 293);
            this.dgvForAgreed.MultiSelect = false;
            this.dgvForAgreed.Name = "dgvForAgreed";
            this.dgvForAgreed.ReadOnly = true;
            this.dgvForAgreed.RowHeadersVisible = false;
            this.dgvForAgreed.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvForAgreed.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvForAgreed.RowTemplate.DefaultCellStyle.Format = "N4";
            this.dgvForAgreed.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvForAgreed.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvForAgreed.Size = new System.Drawing.Size(290, 170);
            this.dgvForAgreed.TabIndex = 4;
            this.dgvForAgreed.Visible = false;
            this.dgvForAgreed.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvForAgreed_Scroll);
            this.dgvForAgreed.SelectionChanged += new System.EventHandler(this.dgvForAgreed_SelectionChanged);
            // 
            // panel6
            // 
            this.panel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel6.Location = new System.Drawing.Point(0, 463);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(290, 50);
            this.panel6.TabIndex = 5;
            // 
            // dgvOnProduction
            // 
            this.dgvOnProduction.AllowUserToAddRows = false;
            this.dgvOnProduction.AllowUserToDeleteRows = false;
            this.dgvOnProduction.AllowUserToResizeRows = false;
            this.dgvOnProduction.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvOnProduction.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvOnProduction.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvOnProduction.ColumnHeadersVisible = false;
            this.dgvOnProduction.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvOnProduction.Location = new System.Drawing.Point(0, 733);
            this.dgvOnProduction.MultiSelect = false;
            this.dgvOnProduction.Name = "dgvOnProduction";
            this.dgvOnProduction.ReadOnly = true;
            this.dgvOnProduction.RowHeadersVisible = false;
            this.dgvOnProduction.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvOnProduction.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvOnProduction.RowTemplate.DefaultCellStyle.Format = "N4";
            this.dgvOnProduction.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvOnProduction.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvOnProduction.Size = new System.Drawing.Size(290, 170);
            this.dgvOnProduction.TabIndex = 6;
            this.dgvOnProduction.Visible = false;
            this.dgvOnProduction.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvOnProduction_Scroll);
            this.dgvOnProduction.SelectionChanged += new System.EventHandler(this.dgvOnProduction_SelectionChanged);
            // 
            // panel5
            // 
            this.panel5.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel5.Location = new System.Drawing.Point(0, 243);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(290, 50);
            this.panel5.TabIndex = 3;
            // 
            // dgvAgreed
            // 
            this.dgvAgreed.AllowUserToAddRows = false;
            this.dgvAgreed.AllowUserToDeleteRows = false;
            this.dgvAgreed.AllowUserToResizeRows = false;
            this.dgvAgreed.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvAgreed.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvAgreed.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAgreed.ColumnHeadersVisible = false;
            this.dgvAgreed.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvAgreed.Location = new System.Drawing.Point(0, 513);
            this.dgvAgreed.MultiSelect = false;
            this.dgvAgreed.Name = "dgvAgreed";
            this.dgvAgreed.ReadOnly = true;
            this.dgvAgreed.RowHeadersVisible = false;
            this.dgvAgreed.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvAgreed.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvAgreed.RowTemplate.DefaultCellStyle.Format = "N4";
            this.dgvAgreed.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvAgreed.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAgreed.Size = new System.Drawing.Size(290, 170);
            this.dgvAgreed.TabIndex = 1;
            this.dgvAgreed.Visible = false;
            this.dgvAgreed.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvAgreed_Scroll);
            this.dgvAgreed.SelectionChanged += new System.EventHandler(this.dgvAgreed_SelectionChanged);
            // 
            // panel4
            // 
            this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel4.Location = new System.Drawing.Point(0, 23);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(290, 50);
            this.panel4.TabIndex = 2;
            // 
            // dgvCalculations
            // 
            this.dgvCalculations.AllowUserToAddRows = false;
            this.dgvCalculations.AllowUserToDeleteRows = false;
            this.dgvCalculations.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvCalculations.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvCalculations.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgvCalculations.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCalculations.ColumnHeadersVisible = false;
            this.dgvCalculations.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvCalculations.Location = new System.Drawing.Point(0, 0);
            this.dgvCalculations.Name = "dgvCalculations";
            this.dgvCalculations.ReadOnly = true;
            this.dgvCalculations.RowHeadersVisible = false;
            this.dgvCalculations.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvCalculations.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvCalculations.RowTemplate.DefaultCellStyle.Format = "N3";
            this.dgvCalculations.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvCalculations.Size = new System.Drawing.Size(290, 23);
            this.dgvCalculations.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.lbSumRank);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 40);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(307, 37);
            this.panel3.TabIndex = 2;
            // 
            // lbSumRank
            // 
            this.lbSumRank.BackColor = System.Drawing.SystemColors.ControlDark;
            this.lbSumRank.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lbSumRank.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbSumRank.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            this.lbSumRank.Location = new System.Drawing.Point(0, 0);
            this.lbSumRank.Name = "lbSumRank";
            this.lbSumRank.Size = new System.Drawing.Size(307, 37);
            this.lbSumRank.TabIndex = 0;
            this.lbSumRank.Text = "Сумма по станку";
            this.lbSumRank.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgvNotConfirmed
            // 
            this.dgvNotConfirmed.AllowUserToAddRows = false;
            this.dgvNotConfirmed.AllowUserToDeleteRows = false;
            this.dgvNotConfirmed.AllowUserToResizeRows = false;
            this.dgvNotConfirmed.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvNotConfirmed.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvNotConfirmed.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvNotConfirmed.ColumnHeadersVisible = false;
            this.dgvNotConfirmed.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvNotConfirmed.Location = new System.Drawing.Point(0, 73);
            this.dgvNotConfirmed.MultiSelect = false;
            this.dgvNotConfirmed.Name = "dgvNotConfirmed";
            this.dgvNotConfirmed.ReadOnly = true;
            this.dgvNotConfirmed.RowHeadersVisible = false;
            this.dgvNotConfirmed.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvNotConfirmed.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvNotConfirmed.RowTemplate.DefaultCellStyle.Format = "N4";
            this.dgvNotConfirmed.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvNotConfirmed.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvNotConfirmed.Size = new System.Drawing.Size(290, 170);
            this.dgvNotConfirmed.TabIndex = 7;
            this.dgvNotConfirmed.Visible = false;
            this.dgvNotConfirmed.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvNotConfirmed_Scroll);
            this.dgvNotConfirmed.SelectionChanged += new System.EventHandler(this.dgvNotConfirmed_SelectionChanged);
            // 
            // panel7
            // 
            this.panel7.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel7.Location = new System.Drawing.Point(0, 683);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(290, 50);
            this.panel7.TabIndex = 8;
            // 
            // panel8
            // 
            this.panel8.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel8.Location = new System.Drawing.Point(0, 903);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(290, 50);
            this.panel8.TabIndex = 9;
            // 
            // dgvInProduction
            // 
            this.dgvInProduction.AllowUserToAddRows = false;
            this.dgvInProduction.AllowUserToDeleteRows = false;
            this.dgvInProduction.AllowUserToResizeRows = false;
            this.dgvInProduction.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvInProduction.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgvInProduction.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvInProduction.ColumnHeadersVisible = false;
            this.dgvInProduction.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvInProduction.Location = new System.Drawing.Point(0, 953);
            this.dgvInProduction.MultiSelect = false;
            this.dgvInProduction.Name = "dgvInProduction";
            this.dgvInProduction.ReadOnly = true;
            this.dgvInProduction.RowHeadersVisible = false;
            this.dgvInProduction.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dgvInProduction.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.dgvInProduction.RowTemplate.DefaultCellStyle.Format = "N4";
            this.dgvInProduction.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dgvInProduction.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvInProduction.Size = new System.Drawing.Size(290, 170);
            this.dgvInProduction.TabIndex = 10;
            this.dgvInProduction.Visible = false;
            this.dgvInProduction.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvInProduction_Scroll);
            this.dgvInProduction.SelectionChanged += new System.EventHandler(this.dgvInProduction_SelectionChanged);
            // 
            // MachineControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.Name = "MachineControl";
            this.Size = new System.Drawing.Size(307, 769);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvForAgreed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOnProduction)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAgreed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalculations)).EndInit();
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvNotConfirmed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvInProduction)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lbName;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dgvCalculations;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label lbSumRank;
        private System.Windows.Forms.DataGridView dgvAgreed;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.DataGridView dgvOnProduction;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.DataGridView dgvForAgreed;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.DataGridView dgvInProduction;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.DataGridView dgvNotConfirmed;
        private System.Windows.Forms.Panel panel7;
    }
}
