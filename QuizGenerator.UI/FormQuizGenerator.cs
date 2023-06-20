using QuizGenerator.Core;

namespace QuizGenerator.UI
{
	public partial class FormQuizGenerator : Form, ILogger
	{
		public FormQuizGenerator()
		{
			InitializeComponent();
		}

		public void Log(string msg, int indentTabs = 0)
		{
			if (this.Disposing || this.IsDisposed)
				return;

			msg = $"{new string(' ', indentTabs * 2)}{msg}";

			// Update the UI through the main UI thread (thread safe)
			this.Invoke((MethodInvoker)delegate
			{
				// Append indent tabs
				this.richTextBoxLogs.AppendText(new string(' ', indentTabs * 2));

				// Append the message to the logs (with formatting)
				richTextBoxLogs.SelectionStart = richTextBoxLogs.Text.Length;
				richTextBoxLogs.SelectionLength = 0;
				this.richTextBoxLogs.SelectionColor = ColorTranslator.FromHtml("#333");
				this.richTextBoxLogs.SelectionFont = new Font(richTextBoxLogs.Font, FontStyle.Regular);
				this.richTextBoxLogs.AppendText(msg);

				// Append a new line
				LogNewLine();
			});
		}

		public void LogError(string errMsg, string errTitle = "Error", int indentTabs = 0)
		{
			if (this.Disposing || this.IsDisposed)
				return;

			// Update the UI through the main UI thread (thread safe)
			this.Invoke((MethodInvoker)delegate
			{
				// Append indent tabs
				this.richTextBoxLogs.AppendText(new string(' ', indentTabs * 2));

				// Append the error title to the logs
				richTextBoxLogs.SelectionStart = richTextBoxLogs.Text.Length;
				richTextBoxLogs.SelectionLength = 0;
				this.richTextBoxLogs.SelectionColor = ColorTranslator.FromHtml("#922");
				this.richTextBoxLogs.SelectionFont = new Font(richTextBoxLogs.Font, FontStyle.Bold);
				this.richTextBoxLogs.AppendText(errTitle + ": ");

				// Append the error message to the logs
				richTextBoxLogs.SelectionStart = richTextBoxLogs.Text.Length;
				richTextBoxLogs.SelectionLength = 0;
				this.richTextBoxLogs.SelectionFont = new Font(richTextBoxLogs.Font, FontStyle.Regular);
				this.richTextBoxLogs.AppendText(errMsg);

				// Append a new line
				this.LogNewLine();
			});
		}

		public void LogNewLine()
		{
			this.richTextBoxLogs.AppendText("\n");
			richTextBoxLogs.SelectionStart = richTextBoxLogs.Text.Length;
			richTextBoxLogs.SelectionLength = 0;
			richTextBoxLogs.ScrollToCaret();
		}
		public void LogException(Exception ex)
		{
			this.LogError(ex.Message);
			this.LogError(ex.StackTrace, "Exception", 1);
		}

		private void FormQuizGenerator_Load(object sender, EventArgs e)
		{
			string startupFolder = Application.StartupPath;
			string inputFolder = Path.Combine(startupFolder, @"../../../../input");
			this.textBoxInputFile.Text = Path.GetFullPath(inputFolder + @"/questions.docx");

			string outputFolder = Path.Combine(startupFolder, @"../../../../output");
			this.textBoxOutputFolder.Text = Path.GetFullPath(outputFolder);

			this.ActiveControl = this.buttonGenerate;
		}

		private void buttonGenerate_Click(object sender, EventArgs e)
		{
			string inputFilePath = this.textBoxInputFile.Text;
			string outputFolderPath = this.textBoxOutputFolder.Text;
			RandomizedQuizGenerator quizGenerator = new RandomizedQuizGenerator(this);
			quizGenerator.GenerateQuiz(inputFilePath, outputFolderPath);
		}

		private void buttonChooseInputFile_Click(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog
			{
				InitialDirectory = Path.GetDirectoryName(this.textBoxInputFile.Text),
				FileName = this.textBoxInputFile.Text,
				Filter = "MS Word files (*.docx)|*.docx|All files (*.*)|*.*",
				CheckFileExists = true,
				FilterIndex = 1,
				RestoreDirectory = true
			};

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				this.textBoxInputFile.Text = openFileDialog.FileName;
			}
		}

		private void buttonChooseFolder_Click(object sender, EventArgs e)
		{
			FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog()
			{
				InitialDirectory = this.textBoxOutputFolder.Text
			};

			if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
			{
				this.textBoxOutputFolder.Text = folderBrowserDialog.SelectedPath;
			}
		}
	}
}