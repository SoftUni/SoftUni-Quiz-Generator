using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizParser
	{
		private const string QuizStartTag = "~~~ Quiz:";
		private const string QuizEndTag = "~~~ Quiz End ~~~";
		private const string QuestionGroupTag = "~~~ Question Group:";
		private const string QuestionTag = "~~~ Question ~~~";
		private const string CorrectAnswerTag = "Correct.";
		private const string WrongAnswerTag = "Wrong.";

		private ILogger logger;

		public QuizParser(ILogger logger)
		{ 
			this.logger = logger;
		}

		public QuizDocument Parse(Word.Document doc)
		{
			var quiz = new QuizDocument();
			quiz.QuestionGroups = new List<QuizQuestionGroup>();
			QuizQuestionGroup? group = null;
			QuizQuestion? question = null;
			int quizHeaderStartPos = 0;
			int groupHeaderStartPos = 0;
			int questionHeaderStartPos = 0;
			int questionFooterStartPos = 0;
			Word.Paragraph paragraph;

			for (int paragraphIndex = 1; paragraphIndex <= doc.Paragraphs.Count; paragraphIndex++)
			{
				paragraph = doc.Paragraphs[paragraphIndex];
				var text = paragraph.Range.Text.Trim();

				if (text.StartsWith(QuizStartTag))
				{
					// ~~~ Quiz: { "VariantsToGenerate": 5, "AnswersPerQuestion": 4 } ~~~
					var settings = ParseSettings(text, QuizStartTag);
					quiz.VariantsToGenerate = settings.VariantsToGenerate;
					quiz.AnswersPerQuestion = settings.AnswersPerQuestion;
					quizHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuestionGroupTag))
				{
					// ~~~ Question Group: { "QuestionsToGenerate": 1, "SkipHeader": true } ~~~
					SaveQuizHeader();
					SaveGroupHeader();
					SaveQuestionFooter();
					group = new QuizQuestionGroup();
					group.Questions = new List<QuizQuestion>();
					var settings = ParseSettings(text, QuestionGroupTag);
					group.QuestionsToGenerate = settings.QuestionsToGenerate;
					group.AnswersPerQuestion = settings.AnswersPerQuestion;
					if (group.AnswersPerQuestion == 0)
						group.AnswersPerQuestion = quiz.AnswersPerQuestion;
					quiz.QuestionGroups.Add(group);
					groupHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuestionTag))
				{
					// ~~~ Question ~~~
					SaveGroupHeader();
					SaveQuestionFooter();
					question = new QuizQuestion();
					question.Answers = new List<QuestionAnswer>();
					group.Questions.Add(question);
					questionHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(CorrectAnswerTag) || text.StartsWith(WrongAnswerTag))
				{
					// Wrong. Some wrong answer
					// Correct. Some correct answer
					SaveQuestionHeader();
					int answerStartRange;
					if (text.StartsWith(CorrectAnswerTag))
						answerStartRange = CorrectAnswerTag.Length;
					else
						answerStartRange = WrongAnswerTag.Length;
					if (text.Length > answerStartRange && text[answerStartRange] == ' ')
						answerStartRange++;
					question.Answers.Add(new QuestionAnswer
					{
						Content = doc.Range(
							paragraph.Range.Start + answerStartRange, 
							paragraph.Range.End),
						IsCorrect = text.StartsWith(CorrectAnswerTag)
					});
					questionFooterStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuizEndTag))
				{
					SaveGroupHeader();
					SaveQuestionFooter();

					// Take all following paragraphs to the end of the document
					var start = paragraph.Range.End;
					var end = doc.Content.End;
					quiz.ContentAfterQuestions = doc.Range(start, end);
					break;
				}
			}

			return quiz;

			void SaveQuizHeader()
			{
				if (quiz != null && 
					quiz.ContentBeforeQuestions == null && 
					quizHeaderStartPos != 0)
				{
					quiz.ContentBeforeQuestions =
						doc.Range(quizHeaderStartPos, paragraph.Range.Start - 1);
				}
				quizHeaderStartPos = 0;
			}

			void SaveGroupHeader()
			{
				if (group != null && 
					group.ContentBeforeQuestions == null &&
					groupHeaderStartPos != 0)
				{
					group.ContentBeforeQuestions =
						doc.Range(groupHeaderStartPos, paragraph.Range.Start - 1);
				}
				groupHeaderStartPos = 0;
			}

			void SaveQuestionHeader()
			{
				if (question != null && 
					question.ContentBeforeAnswers == null &&
					questionHeaderStartPos != 0)
				{
					question.ContentBeforeAnswers =
						doc.Range(questionHeaderStartPos, paragraph.Range.Start - 1);
				}
				questionHeaderStartPos = 0;
			}

			void SaveQuestionFooter()
			{
				if (question != null && 
					question.ContentAfterAnswers == null &&
					questionFooterStartPos != 0 &&
					questionFooterStartPos < paragraph.Range.Start)
				{
					question.ContentAfterAnswers =
						doc.Range(questionFooterStartPos, paragraph.Range.Start - 1);
				}
				questionFooterStartPos = 0;
			}
		}

		private static QuizSettings ParseSettings(string text, string tag)
		{
			var json = text.Substring(tag.Length).Trim();
			json = json.Replace("~~~", "").Trim();
			if (string.IsNullOrEmpty(json))
				json = "{}";
			QuizSettings settings = JsonConvert.DeserializeObject<QuizSettings>(json);
			return settings;
		}

		public static string TruncateString(string? input, int maxLength = 150)
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

		public void LogQuiz(QuizDocument quiz)
		{
			this.logger.Log($"Quiz document:");
			this.logger.Log($" - VariantsToGenerate: {quiz.VariantsToGenerate}");
			this.logger.Log($" - TotalAvailableQuestions: {quiz.TotalAvailableQuestions}");
			this.logger.Log($" - AnswersPerQuestion: {quiz.AnswersPerQuestion}");
			string quizHeaderText = TruncateString(quiz.ContentBeforeQuestions.Text);
			this.logger.Log($"Quiz header: {quizHeaderText}", 1);
			this.logger.Log($"Question groups = {quiz.QuestionGroups.Count}", 1);
			for (int groupIndex = 0; groupIndex < quiz.QuestionGroups.Count; groupIndex++)
			{
				this.logger.Log($"[Question Group #{groupIndex+1}]", 1);
				QuizQuestionGroup group = quiz.QuestionGroups[groupIndex];
				string groupHeaderText = TruncateString(group.ContentBeforeQuestions?.Text);
				this.logger.Log($"Group header: {groupHeaderText}", 2);
				this.logger.Log($"Questions = {group.Questions.Count}", 2);
				for (int questionIndex = 0; questionIndex < group.Questions.Count; questionIndex++)
				{
					this.logger.Log($"[Question #{questionIndex+1}]", 2);
					QuizQuestion question = group.Questions[questionIndex];
					string questionContent = TruncateString(question.ContentBeforeAnswers?.Text);
					this.logger.Log($"Question content: {questionContent}", 3);
					this.logger.Log($"Answers = {question.Answers.Count}", 3);
					foreach (var answer in question.Answers)
					{
						string prefix = answer.IsCorrect ? "Correct answer" : "Wrong answer";
						string answerText = TruncateString(answer.Content.Text);
						this.logger.Log($"{prefix}: {answerText}", 4);
					}
					string questionFooterText = TruncateString(question.ContentAfterAnswers?.Text);
					if (questionFooterText == "")
						questionFooterText = "(empty)";
					this.logger.Log($"Question footer: {questionFooterText}", 3);
				}
			}
			string quizFooterText = TruncateString(quiz.ContentAfterQuestions?.Text);
			this.logger.Log($"Quiz footer: {quizFooterText}", 1);
		}
	}
}
