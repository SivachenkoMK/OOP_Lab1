﻿using System.Windows.Forms;
using System.Drawing;

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
            this.panel = new System.Windows.Forms.Panel();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.calculateButton = new System.Windows.Forms.Button();
            this.textBox = new System.Windows.Forms.TextBox();
            this.delRowButton = new System.Windows.Forms.Button();
            this.addRowButton = new System.Windows.Forms.Button();
            this.delColButton = new System.Windows.Forms.Button();
            this.rowLabel = new System.Windows.Forms.Label();
            this.addColButton = new System.Windows.Forms.Button();
            this.label = new System.Windows.Forms.Label();
            this.saveButton = new System.Windows.Forms.Button();
            this.openButton = new System.Windows.Forms.Button();
            this.panel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel
            // 
            this.panel.Controls.Add(this.dataGridView);
            this.panel.Location = new System.Drawing.Point(1, 44);
            this.panel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(1920, 964);
            this.panel.TabIndex = 1;
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(0, 0);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 60;
            this.dataGridView.RowTemplate.Height = 30;
            this.dataGridView.Size = new System.Drawing.Size(1920, 964);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // splitContainer
            // 
            this.splitContainer.Cursor = System.Windows.Forms.Cursors.VSplit;
            this.splitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer.Location = new System.Drawing.Point(1, 2);
            this.splitContainer.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.splitContainer.Name = "splitContainer";
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.calculateButton);
            this.splitContainer.Panel1.Controls.Add(this.textBox);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.delRowButton);
            this.splitContainer.Panel2.Controls.Add(this.addRowButton);
            this.splitContainer.Panel2.Controls.Add(this.delColButton);
            this.splitContainer.Panel2.Controls.Add(this.rowLabel);
            this.splitContainer.Panel2.Controls.Add(this.addColButton);
            this.splitContainer.Panel2.Controls.Add(this.label);
            this.splitContainer.Panel2.Controls.Add(this.saveButton);
            this.splitContainer.Panel2.Controls.Add(this.openButton);
            this.splitContainer.Size = new System.Drawing.Size(1920, 80);
            this.splitContainer.SplitterDistance = 488;
            this.splitContainer.SplitterWidth = 3;
            this.splitContainer.TabIndex = 2;
            // 
            // calculateButton
            // 
            this.calculateButton.Location = new System.Drawing.Point(287, 3);
            this.calculateButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.calculateButton.Name = "calculateButton";
            this.calculateButton.Size = new System.Drawing.Size(104, 32);
            this.calculateButton.TabIndex = 1;
            this.calculateButton.Text = "Evaluate";
            this.calculateButton.UseVisualStyleBackColor = true;
            this.calculateButton.Click += new System.EventHandler(this.calculateButton_Click);
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(2, 4);
            this.textBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(281, 27);
            this.textBox.TabIndex = 0;
            // 
            // delRowButton
            // 
            this.delRowButton.Location = new System.Drawing.Point(249, 4);
            this.delRowButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.delRowButton.Name = "delRowButton";
            this.delRowButton.Size = new System.Drawing.Size(32, 32);
            this.delRowButton.TabIndex = 7;
            this.delRowButton.Text = "-";
            this.delRowButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.delRowButton.UseVisualStyleBackColor = true;
            this.delRowButton.Click += new System.EventHandler(this.delRowButton_Click);
            // 
            // addRowButton
            // 
            this.addRowButton.Location = new System.Drawing.Point(218, 4);
            this.addRowButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.addRowButton.Name = "addRowButton";
            this.addRowButton.Size = new System.Drawing.Size(32, 32);
            this.addRowButton.TabIndex = 6;
            this.addRowButton.Text = "+";
            this.addRowButton.UseVisualStyleBackColor = true;
            this.addRowButton.Click += new System.EventHandler(this.addRowButton_Click);
            // 
            // delColButton
            // 
            this.delColButton.Location = new System.Drawing.Point(114, 3);
            this.delColButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.delColButton.Name = "delColButton";
            this.delColButton.Size = new System.Drawing.Size(32, 32);
            this.delColButton.TabIndex = 5;
            this.delColButton.Text = "-";
            this.delColButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.delColButton.UseVisualStyleBackColor = true;
            this.delColButton.Click += new System.EventHandler(this.delColButton_Click);
            // 
            // rowLabel
            // 
            this.rowLabel.AutoSize = true;
            this.rowLabel.Location = new System.Drawing.Point(170, 7);
            this.rowLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.rowLabel.Name = "rowLabel";
            this.rowLabel.Size = new System.Drawing.Size(47, 20);
            this.rowLabel.TabIndex = 4;
            this.rowLabel.Text = "Rows:";
            // 
            // addColButton
            // 
            this.addColButton.Location = new System.Drawing.Point(83, 3);
            this.addColButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.addColButton.Name = "addColButton";
            this.addColButton.Size = new System.Drawing.Size(32, 32);
            this.addColButton.TabIndex = 3;
            this.addColButton.Text = "+";
            this.addColButton.UseVisualStyleBackColor = true;
            this.addColButton.Click += new System.EventHandler(this.addColButton_Click);
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(13, 6);
            this.label.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(69, 20);
            this.label.TabIndex = 2;
            this.label.Text = "Columns:";
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(436, 2);
            this.saveButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(80, 32);
            this.saveButton.TabIndex = 1;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // openButton
            // 
            this.openButton.Location = new System.Drawing.Point(345, 2);
            this.openButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(80, 32);
            this.openButton.TabIndex = 0;
            this.openButton.Text = "Open";
            this.openButton.UseVisualStyleBackColor = true;
            this.openButton.Click += new System.EventHandler(this.openButton_Click);
            // 
            // MyExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1539, 514);
            this.Controls.Add(this.panel);
            this.Controls.Add(this.splitContainer);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "MyExcel";
            this.Text = "MyExcel";
            this.panel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel1.PerformLayout();
            this.splitContainer.Panel2.ResumeLayout(false);
            this.splitContainer.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Panel panel;
        private DataGridView dataGridView;
        private SplitContainer splitContainer;
        private Button calculateButton;
        private TextBox textBox;
        private Label label;
        private Button saveButton;
        private Button openButton;
        private Button addColButton;
        private Button delRowButton;
        private Button addRowButton;
        private Button delColButton;
        private Label rowLabel;
    }
}

