using static QuizGenerator.Core.StringUtils;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace QuizGenerator.Core
{
	public class RandomizedQuizGenerator
	{
		private ILogger logger;
		private Word.Application wordApp;

		public RandomizedQuizGenerator(ILogger logger)
		{
			this.logger = logger;
		}

		public void GenerateQuiz(string inputFilePath, string outputFolderPath)
		{
			this.logger.Log("Quiz generation started.");
			this.logger.LogNewLine();

			if (KillAllProcesses("WINWORD"))
				Console.WriteLine("MS Word (WINWORD.EXE) is still running -> process terminated.");

			// Start MS Word and open the input file
			this.wordApp = new Word.Application();
			this.wordApp.Visible = false; // Show / hide MS Word app window
			this.wordApp.ScreenUpdating = false; // Enable / disable screen updates after each change
			var inputDoc = this.wordApp.Documents.Open(inputFilePath);
			
			try
			{
				// Parse the input MS Word document
				this.logger.Log("Parsing the input document: " + inputFilePath);
				QuizParser quizParser = new QuizParser(this.logger);
				QuizDocument quiz = quizParser.Parse(inputDoc);
				this.logger.Log("Input document parsed successfully.");

				// Display the quiz content (question groups + questions + answers)
				quizParser.LogQuiz(quiz);

				// Generate the randomized quiz variants
				this.logger.LogNewLine();
				this.logger.Log("Generating quizes...");
				this.logger.Log($"  (output path = {outputFolderPath})");
				GenerateRandomizedQuizVariants(quiz, inputFilePath, outputFolderPath);
				
				this.logger.LogNewLine();
				this.logger.Log("Quiz generation completed.");
				this.logger.LogNewLine();
			}
			catch (Exception ex)
			{
				this.logger.LogException(ex);
			}
			finally
			{
				inputDoc.Close();
				this.wordApp.Quit();
			}
		}

		private void GenerateRandomizedQuizVariants(
			QuizDocument quiz, string inputFilePath, string outputFolderPath)
		{
			// Initialize the output folder (create it and ensure it is empty)
			this.logger.Log($"Initializing output folder: {outputFolderPath}");
			if (Directory.Exists(outputFolderPath))
			{
				Directory.Delete(outputFolderPath, true);
			}
			Directory.CreateDirectory(outputFolderPath);

			// Prepare the answer sheet for all variants
			List<List<char>> quizAnswerSheet = new List<List<char>>();	

			// Generate the requested randomized quiz variants, one by one
			for (int quizVariant = 1; quizVariant <= quiz.VariantsToGenerate; quizVariant++)
			{
				this.logger.LogNewLine();
				this.logger.Log($"Generating randomized quiz: variant #{quizVariant} out of {quiz.VariantsToGenerate} ...");
				string outputFilePath = outputFolderPath + Path.DirectorySeparatorChar +
					"quiz" + quizVariant.ToString("000") + ".docx";
				RandomizedQuiz randQuiz = RandomizedQuiz.GenerateFromQuizData(quiz);
				WriteRandomizedQuizToFile(
					randQuiz, quizVariant, inputFilePath, outputFilePath, quiz.LangCode);
				List<char> answers = ExtractAnswersAsLetters(randQuiz, quiz.LangCode);
				quizAnswerSheet.Add(answers);
				this.logger.Log($"Randomized quiz: variant #{quizVariant} out of {quiz.VariantsToGenerate} generated successfully.");
			}

			WriteAnswerSheetToHTMLFile(quizAnswerSheet, outputFolderPath);
		}

		private void WriteRandomizedQuizToFile(RandomizedQuiz randQuiz, 
			int quizVariant, string inputFilePath, string outputFilePath, string langCode)
		{
			File.Copy(inputFilePath, outputFilePath, true);

			// Open the output file in MS Word
			var outputDoc = this.wordApp.Documents.Open(outputFilePath);
			try
			{
				// Select all content in outputDoc and delete the seletion
				this.wordApp.Selection.WholeStory();
				this.wordApp.Selection.Delete();

				// Write the randomized quiz as MS Word document
				this.logger.Log($"Creating randomized quiz document: " + outputFilePath);
				WriteRandomizedQuizToWordDoc(randQuiz, quizVariant, langCode, outputDoc);
			}
			catch (Exception ex)
			{
				this.logger.LogException(ex);
			}
			finally
			{
				outputDoc.Save();
				outputDoc.Close();
			}
		}

		private void WriteRandomizedQuizToWordDoc(RandomizedQuiz quiz, int quizVariant,
			string langCode, Word.Document outputDoc)
		{
			// Print the quiz header in the output MS Word document
			string quizHeaderText = TruncateString(quiz.HeaderContent.Text);
			this.logger.Log($"Quiz header: {quizHeaderText}", 1);
			AppendRange(outputDoc, quiz.HeaderContent);

			// Replace all occurences of "# # #" with the variant number
			Word.Find find = this.wordApp.Selection.Find;
			find.Text = "# # #";
			find.Replacement.Text = quizVariant.ToString("000");
			find.Forward = true;
			find.Wrap = Word.WdFindWrap.wdFindContinue;
			object replaceAll = Word.WdReplace.wdReplaceAll;
			find.Execute(Replace: ref replaceAll);

			int questionNumber = 0;
			this.logger.Log($"Question groups = {quiz.QuestionGroups.Count}", 1);
			for (int groupIndex = 0; groupIndex < quiz.QuestionGroups.Count; groupIndex++)
			{
				this.logger.Log($"[Question Group #{groupIndex + 1}]", 1);
				QuizQuestionGroup group = quiz.QuestionGroups[groupIndex];
				string groupHeaderText = TruncateString(group.HeaderContent?.Text);
				this.logger.Log($"Group header: {groupHeaderText}", 2);
				if (!group.SkipHeader)
				{
					AppendRange(outputDoc, group.HeaderContent);
				}
				this.logger.Log($"Questions = {group.Questions.Count}", 2);
				for (int questionIndex = 0; questionIndex < group.Questions.Count; questionIndex++)
				{
					this.logger.Log($"[Question #{questionIndex + 1}]", 2);
					QuizQuestion question = group.Questions[questionIndex];
					string questionContent = TruncateString(question.HeaderContent?.Text);
					this.logger.Log($"Question content: {questionContent}", 3);
					questionNumber++;
					AppendText(outputDoc, $"{questionNumber}. ");
					AppendRange(outputDoc, question.HeaderContent);
					this.logger.Log($"Answers = {question.Answers.Count}", 3);
					char letter = GetStartLetter(langCode);
					foreach (var answer in question.Answers)
					{
						string prefix = answer.IsCorrect ? "Correct answer" : "Wrong answer";
						string answerText = TruncateString(answer.Content.Text);
						this.logger.Log($"{prefix}: {letter}) {answerText}", 4);
						AppendText(outputDoc, $"{letter}) ");
						AppendRange(outputDoc, answer.Content);
						letter++;
					}
					string questionFooterText = TruncateString(question.FooterContent?.Text);
					if (questionFooterText == "")
						questionFooterText = "(empty)";
					this.logger.Log($"Question footer: {questionFooterText}", 3);
					AppendRange(outputDoc, question.FooterContent);
				}
			}
			string quizFooterText = TruncateString(quiz.FooterContent?.Text);
			this.logger.Log($"Quiz footer: {quizFooterText}", 1);
			AppendRange(outputDoc, quiz.FooterContent);
		}

		public void AppendRange(Word.Document targetDocument, Word.Range sourceRange)
		{
			if (sourceRange != null)
			{
				// Get the range at the end of the target document
				Word.Range targetRange = targetDocument.Content;
				object wdColapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
				targetRange.Collapse(ref wdColapseEnd);

				// Insert the source range of formatted text to the target range
				targetRange.FormattedText = sourceRange.FormattedText;
			}
		}

		public void AppendText(Word.Document targetDocument, string text)
		{
			// Get the range at the end of the target document
			Word.Range targetRange = targetDocument.Content;
			object wdColapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
			targetRange.Collapse(ref wdColapseEnd);

			// Insert the source range of formatted text to the target range
			targetRange.Text = text;
		}

		private List<char> ExtractAnswersAsLetters(RandomizedQuiz randQuiz, string langCode)
		{
			char startLetter = GetStartLetter(langCode);

			List<char> answers = new List<char>();
			foreach (var question in randQuiz.AllQuestions)
			{
				int correctAnswerIndex = FindCorrectAnswerIndex(question.Answers);
				char answer = (char)(startLetter + correctAnswerIndex);
				answers.Add(answer);
			}

			return answers;
		}

		private static char GetStartLetter(string langCode)
		{
			char startLetter;
			if (langCode == "EN")
				startLetter = 'a'; // Latin letter 'a'
			else if (langCode == "BG")
				startLetter = 'а'; // Cyrillyc letter 'а'
			else
				throw new Exception("Unsupported language: " + langCode);
			return startLetter;
		}

		private int FindCorrectAnswerIndex(List<QuestionAnswer> answers)
		{
			for (int index = 0; index < answers.Count; index++)
			{
				if (answers[index].IsCorrect)
					return index;
			}

			// No correct answer found in the list of answers
			return -1;
		}

		private void WriteAnswerSheetToHTMLFile(
			List<List<char>> quizAnswerSheet, string outputFilePath)
		{
			string outputFileName = outputFilePath + Path.DirectorySeparatorChar + "answers.html";

			this.logger.LogNewLine();
			this.logger.Log($"Writing answers sheet: {outputFileName}");
			for (int quizIndex = 0; quizIndex < quizAnswerSheet.Count; quizIndex++)
			{
				List<char> answers = quizAnswerSheet[quizIndex];
				string answersAsString = $"Variant #{quizIndex + 1}: {string.Join(" ", answers)}";
				this.logger.Log(answersAsString, 1);
			}

			List<string> html = new List<string>();
			html.Add("<table border='1'>");
			html.Add("  <tr>");
			html.Add("    <td>Var</td>");
			for (int questionIndex = 0; questionIndex < quizAnswerSheet[0].Count; questionIndex++)
			{
				html.Add($"    <td>{questionIndex + 1}</td>");
			}
			html.Add("  </tr>");
			for (int quizIndex = 0; quizIndex < quizAnswerSheet.Count; quizIndex++)
			{
				html.Add("  <tr>");
				html.Add($"    <td>{(quizIndex + 1).ToString("000")}</td>");
				foreach (var answer in quizAnswerSheet[quizIndex])
				{
					html.Add($"    <td>{answer}</td>");
				}
				html.Add("  </tr>");
			}
			html.Add("</table>");

			File.WriteAllLines(outputFileName, html);
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
					this.logger.Log($"Process {processName} ({process.Id}) stopped.");
				}
				catch
				{
					this.logger.LogError($"Process {processName} ({process.Id}) is running, but cannot be stopped!");
				}
			}
			return (killedProcessesCount > 0);
		}
	}
}
