namespace QuizGenerator.UI
{
	partial class FormQuizGenerator
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
			labelInputFile = new Label();
			textBoxInputFile = new TextBox();
			buttonChooseInputFile = new Button();
			labelOutputFolder = new Label();
			textBoxOutputFolder = new TextBox();
			buttonChooseFolder = new Button();
			buttonGenerate = new Button();
			richTextBoxLogs = new RichTextBox();
			SuspendLayout();
			// 
			// labelInputFile
			// 
			labelInputFile.AutoSize = true;
			labelInputFile.Location = new Point(10, 10);
			labelInputFile.Margin = new Padding(4, 0, 4, 0);
			labelInputFile.Name = "labelInputFile";
			labelInputFile.Size = new Size(291, 25);
			labelInputFile.TabIndex = 0;
			labelInputFile.Text = "Quiz content (input) file (MS Word):";
			// 
			// textBoxInputFile
			// 
			textBoxInputFile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			textBoxInputFile.BorderStyle = BorderStyle.FixedSingle;
			textBoxInputFile.Location = new Point(15, 39);
			textBoxInputFile.Margin = new Padding(4);
			textBoxInputFile.Name = "textBoxInputFile";
			textBoxInputFile.Size = new Size(871, 31);
			textBoxInputFile.TabIndex = 1;
			// 
			// buttonChooseInputFile
			// 
			buttonChooseInputFile.Anchor = AnchorStyles.Top | AnchorStyles.Right;
			buttonChooseInputFile.Location = new Point(895, 37);
			buttonChooseInputFile.Margin = new Padding(4);
			buttonChooseInputFile.Name = "buttonChooseInputFile";
			buttonChooseInputFile.Size = new Size(149, 36);
			buttonChooseInputFile.TabIndex = 2;
			buttonChooseInputFile.Text = "Choose File";
			buttonChooseInputFile.UseVisualStyleBackColor = true;
			buttonChooseInputFile.Click += buttonChooseInputFile_Click;
			// 
			// labelOutputFolder
			// 
			labelOutputFolder.AutoSize = true;
			labelOutputFolder.Location = new Point(10, 74);
			labelOutputFolder.Name = "labelOutputFolder";
			labelOutputFolder.Size = new Size(163, 25);
			labelOutputFolder.TabIndex = 3;
			labelOutputFolder.Text = "Quiz output folder:";
			// 
			// textBoxOutputFolder
			// 
			textBoxOutputFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			textBoxOutputFolder.BorderStyle = BorderStyle.FixedSingle;
			textBoxOutputFolder.Location = new Point(15, 103);
			textBoxOutputFolder.Margin = new Padding(4);
			textBoxOutputFolder.Name = "textBoxOutputFolder";
			textBoxOutputFolder.Size = new Size(872, 31);
			textBoxOutputFolder.TabIndex = 4;
			// 
			// buttonChooseFolder
			// 
			buttonChooseFolder.Anchor = AnchorStyles.Top | AnchorStyles.Right;
			buttonChooseFolder.Location = new Point(895, 99);
			buttonChooseFolder.Margin = new Padding(4);
			buttonChooseFolder.Name = "buttonChooseFolder";
			buttonChooseFolder.Size = new Size(149, 36);
			buttonChooseFolder.TabIndex = 5;
			buttonChooseFolder.Text = "Choose Folder";
			buttonChooseFolder.UseVisualStyleBackColor = true;
			buttonChooseFolder.Click += buttonChooseFolder_Click;
			// 
			// buttonGenerate
			// 
			buttonGenerate.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
			buttonGenerate.Font = new Font("Segoe UI", 13.8F, FontStyle.Regular, GraphicsUnit.Point);
			buttonGenerate.Location = new Point(15, 152);
			buttonGenerate.Margin = new Padding(4);
			buttonGenerate.Name = "buttonGenerate";
			buttonGenerate.Size = new Size(1029, 47);
			buttonGenerate.TabIndex = 6;
			buttonGenerate.Text = "Generate Quiz";
			buttonGenerate.UseVisualStyleBackColor = true;
			buttonGenerate.Click += buttonGenerate_Click;
			// 
			// richTextBoxLogs
			// 
			richTextBoxLogs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
			richTextBoxLogs.Location = new Point(15, 216);
			richTextBoxLogs.Name = "richTextBoxLogs";
			richTextBoxLogs.ReadOnly = true;
			richTextBoxLogs.Size = new Size(1026, 413);
			richTextBoxLogs.TabIndex = 7;
			richTextBoxLogs.Text = "";
			// 
			// FormQuizGenerator
			// 
			AutoScaleDimensions = new SizeF(10F, 25F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(1053, 641);
			Controls.Add(richTextBoxLogs);
			Controls.Add(buttonGenerate);
			Controls.Add(buttonChooseFolder);
			Controls.Add(textBoxOutputFolder);
			Controls.Add(labelOutputFolder);
			Controls.Add(buttonChooseInputFile);
			Controls.Add(textBoxInputFile);
			Controls.Add(labelInputFile);
			Font = new Font("Segoe UI", 10.8F, FontStyle.Regular, GraphicsUnit.Point);
			Margin = new Padding(4);
			Name = "FormQuizGenerator";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "SoftUni Quiz Generator";
			Load += FormQuizGenerator_Load;
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private Label labelInputFile;
		private TextBox textBoxInputFile;
		private Button buttonChooseInputFile;
		private Label labelOutputFolder;
		private TextBox textBoxOutputFolder;
		private Button buttonChooseFolder;
		private Button buttonGenerate;
		private RichTextBox richTextBoxLogs;
	}
}