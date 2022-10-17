using System.Windows.Forms;

namespace Excel
{
    partial class MyExcel
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new Panel();
            this.dataGridView1 = new DataGridView();
            this.splitContainer1 = new SplitContainer();
            this.calculateButton = new Button();
            this.textBox1 = new TextBox();
            this.delRowButton = new Button();
            this.addRowButton = new Button();
            this.delColButton = new Button();
            this.rowLabel = new Label();
            this.addColButton = new Button();
            this.label1 = new Label();
            this.saveButton = new Button();
            this.openButton = new Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Location = new System.Drawing.Point(1, 47);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1136, 595);
            this.panel1.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 33;
            this.dataGridView1.Size = new System.Drawing.Size(1136, 595);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Cursor = Cursors.VSplit;
            this.splitContainer1.FixedPanel = FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(1, 2);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.calculateButton);
            this.splitContainer1.Panel1.Controls.Add(this.textBox1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.delRowButton);
            this.splitContainer1.Panel2.Controls.Add(this.addRowButton);
            this.splitContainer1.Panel2.Controls.Add(this.delColButton);
            this.splitContainer1.Panel2.Controls.Add(this.rowLabel);
            this.splitContainer1.Panel2.Controls.Add(this.addColButton);
            this.splitContainer1.Panel2.Controls.Add(this.label1);
            this.splitContainer1.Panel2.Controls.Add(this.saveButton);
            this.splitContainer1.Panel2.Controls.Add(this.openButton);
            this.splitContainer1.Size = new System.Drawing.Size(1136, 39);
            this.splitContainer1.SplitterDistance = 488;
            this.splitContainer1.TabIndex = 2;
            // 
            // calculateButton
            // 
            this.calculateButton.Location = new System.Drawing.Point(359, 4);
            this.calculateButton.Name = "calculateButton";
            this.calculateButton.Size = new System.Drawing.Size(126, 34);
            this.calculateButton.TabIndex = 1;
            this.calculateButton.Text = "Calculate";
            this.calculateButton.UseVisualStyleBackColor = true;
            this.calculateButton.Click += new System.EventHandler(this.calculateButton_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(3, 5);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(350, 31);
            this.textBox1.TabIndex = 0;
            // 
            // delRowButton
            // 
            this.delRowButton.Location = new System.Drawing.Point(309, 5);
            this.delRowButton.Name = "delRowButton";
            this.delRowButton.Size = new System.Drawing.Size(30, 32);
            this.delRowButton.TabIndex = 7;
            this.delRowButton.Text = "-";
            this.delRowButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.delRowButton.UseVisualStyleBackColor = true;
            this.delRowButton.Click += new System.EventHandler(this.delRowButton_Click);
            // 
            // addRowButton
            // 
            this.addRowButton.Location = new System.Drawing.Point(273, 5);
            this.addRowButton.Name = "addRowButton";
            this.addRowButton.Size = new System.Drawing.Size(30, 32);
            this.addRowButton.TabIndex = 6;
            this.addRowButton.Text = "+";
            this.addRowButton.UseVisualStyleBackColor = true;
            this.addRowButton.Click += new System.EventHandler(this.addRowButton_Click);
            // 
            // delColButton
            // 
            this.delColButton.Location = new System.Drawing.Point(140, 4);
            this.delColButton.Name = "delColButton";
            this.delColButton.Size = new System.Drawing.Size(30, 32);
            this.delColButton.TabIndex = 5;
            this.delColButton.Text = "-";
            this.delColButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.delColButton.UseVisualStyleBackColor = true;
            this.delColButton.Click += new System.EventHandler(this.delColButton_Click);
            // 
            // rowLabel
            // 
            this.rowLabel.AutoSize = true;
            this.rowLabel.Location = new System.Drawing.Point(213, 9);
            this.rowLabel.Name = "rowLabel";
            this.rowLabel.Size = new System.Drawing.Size(54, 25);
            this.rowLabel.TabIndex = 4;
            this.rowLabel.Text = "Rows";
            // 
            // addColButton
            // 
            this.addColButton.Location = new System.Drawing.Point(104, 4);
            this.addColButton.Name = "addColButton";
            this.addColButton.Size = new System.Drawing.Size(30, 32);
            this.addColButton.TabIndex = 3;
            this.addColButton.Text = "+";
            this.addColButton.UseVisualStyleBackColor = true;
            this.addColButton.Click += new System.EventHandler(this.addColButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 25);
            this.label1.TabIndex = 2;
            this.label1.Text = "Columns";
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(545, 3);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(96, 34);
            this.saveButton.TabIndex = 1;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // openButton
            // 
            this.openButton.Location = new System.Drawing.Point(431, 3);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(97, 34);
            this.openButton.TabIndex = 0;
            this.openButton.Text = "Open";
            this.openButton.UseVisualStyleBackColor = true;
            this.openButton.Click += new System.EventHandler(this.openButton_Click);
            // 
            // MyExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1136, 642);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.splitContainer1);
            this.Name = "MyExcel";
            this.Text = "MyExcel";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Panel panel1;
        private DataGridView dataGridView1;
        private SplitContainer splitContainer1;
        private Button calculateButton;
        private TextBox textBox1;
        private Label label1;
        private Button saveButton;
        private Button openButton;
        private Button addColButton;
        private Button delRowButton;
        private Button addRowButton;
        private Button delColButton;
        private Label rowLabel;
    }
}

