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

		public static string TruncateString(string? input, int maxLength = 100)
		{
			if (input == null)
				input = string.Empty;
			input = input.Replace("\r", " ").Trim();
			if (input.Length <= maxLength)
			{
				return input;
			}
			else if (maxLength <= 3)
			{
				return input.Substring(0, maxLength);
			}
			else
			{
				int midLength = maxLength - 3;
				int halfLength = midLength / 2;
				string start = input.Substring(0, halfLength);
				string end = input.Substring(input.Length - halfLength);
				return $"{start}...{end}";
			}
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

		private void buttonGenerate_Click(object sender, EventArgs e)
		{
			if (KillAllProcesses("WINWORD"))
				Console.WriteLine("MS Word (WINWORD.EXE) is still running -> process terminated.");

			this.Log("Quiz generation started.");
			var wordApp = new Word.Application();
			string inputFilePath = this.textBoxInputFile.Text;
			var doc = wordApp.Documents.Open(inputFilePath);
			try
			{
				this.Log("Parsing the input document: " + inputFilePath);
				QuizParser quizParser = new QuizParser(this);
				QuizDocument quiz = quizParser.Parse(doc);
				this.Log("Input document parsed successfully.");
				LogQuiz(quiz);
			}
			catch (Exception ex)
			{
				this.LogError(ex.Message);
			}
			finally
			{
				doc.Close();
				wordApp.Quit();
			}
		}

		private void LogQuiz(QuizDocument quiz)
		{
			this.Log("Quiz document:");
			string quizHeaderText = TruncateString(quiz.ContentBeforeQuestions.Text);
			this.Log($"Quiz header: {quizHeaderText}", 1);
			this.Log($"Question groups: {quiz.QuestionGroups.Count}", 1);
			foreach (var group in quiz.QuestionGroups)
			{
				string groupHeaderText = TruncateString(group.ContentBeforeQuestions?.Text);
				this.Log($"Group header: {groupHeaderText}", 2);
				this.Log($"Questions: {group.Questions.Count}", 2);
				foreach (var question in group.Questions)
				{
					string questionHeaderText = TruncateString(question.ContentBeforeAnswers?.Text);
					this.Log($"Question content: {questionHeaderText}", 3);
					this.Log($"Answers: {question.Answers.Count}", 3);
					foreach (var answer in question.Answers)
					{
						string prefix = answer.IsCorrect ? "Correct answer" : "Wrong answer";
						string answerText = TruncateString(answer.Content.Text);
						this.Log($"{prefix}: {answerText}", 4);
					}
					string questionFooterText = TruncateString(question.ContentAfterAnswers?.Text);
					if (questionFooterText != null)
						this.Log($"Question footer: {questionFooterText}", 3);
				}
				string groupFooterText = TruncateString(group.ContentAfterQuestions?.Text);
				if (groupFooterText != "")
					this.Log($"Group footer: {groupFooterText}", 2);
			}
			string quizFooterText = TruncateString(quiz.ContentAfterQuestions?.Text);
			this.Log($"Quiz footer: {quizFooterText}", 1);
		}
	}
}