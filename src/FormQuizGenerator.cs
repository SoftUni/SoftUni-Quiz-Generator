using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace SoftUniQuizGenerator
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

		private void LogNewLine()
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
			GenerateQuiz(inputFilePath, outputFolderPath);
		}

		private void GenerateQuiz(string inputFilePath, string outputFolderPath)
		{
			if (KillAllProcesses("WINWORD"))
				Console.WriteLine("MS Word (WINWORD.EXE) is still running -> process terminated.");

			this.Log("Quiz generation started.");
			var wordApp = new Word.Application();
			var doc = wordApp.Documents.Open(inputFilePath);
			try
			{
				this.Log("Parsing the input document: " + inputFilePath);
				QuizParser quizParser = new QuizParser(this);
				QuizDocument quiz = quizParser.Parse(doc);
				this.Log("Input document parsed successfully.");
				quizParser.LogQuiz(quiz);
			}
			catch (Exception ex)
			{
				this.LogException(ex);
			}
			finally
			{
				doc.Close();
				wordApp.Quit();
			}
		}

		public bool KillAllProcesses(string processName)
		{
			Process[] processes = Process.GetProcessesByName(processName);
			int killedProcessesCount = 0;
			foreach (Process process in processes)
			{
				try
				{
					process.Kill();
					killedProcessesCount++;
					this.Log($"Process {processName} ({process.Id}) is stopped.");
				}
				catch
				{
					this.LogError($"Process {processName} ({process.Id}) is running, but cannot be stopped.");
				}
			}
			return (killedProcessesCount > 0);
		}
	}
}