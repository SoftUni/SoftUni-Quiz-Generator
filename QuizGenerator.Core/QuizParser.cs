using static QuizGenerator.Core.StringUtils;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;

namespace QuizGenerator.Core
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
					// ~~~ Quiz: {"VariantsToGenerate":5, "AnswersPerQuestion":4, "Lang":"BG"} ~~~
					this.logger.Log("Parsing: " + text, 1);
					var settings = ParseSettings(text, QuizStartTag);
					quiz.VariantsToGenerate = settings.VariantsToGenerate;
					quiz.AnswersPerQuestion = settings.AnswersPerQuestion;
					quiz.LangCode = settings.Lang;
					quizHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuestionGroupTag))
				{
					// ~~~ Question Group: { "QuestionsToGenerate": 1, "SkipHeader": true } ~~~
					this.logger.Log("Parsing: " + text, 1);
					SaveQuizHeader();
					SaveGroupHeader();
					SaveQuestionFooter();
					group = new QuizQuestionGroup();
					group.Questions = new List<QuizQuestion>();
					var settings = ParseSettings(text, QuestionGroupTag);
					group.QuestionsToGenerate = settings.QuestionsToGenerate;
					group.SkipHeader = settings.SkipHeader;
					group.AnswersPerQuestion = settings.AnswersPerQuestion;
					if (group.AnswersPerQuestion == 0)
						group.AnswersPerQuestion = quiz.AnswersPerQuestion;
					quiz.QuestionGroups.Add(group);
					groupHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuestionTag))
				{
					// ~~~ Question ~~~
					this.logger.Log("Parsing: " + text, 1);
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
					quiz.FooterContent = doc.Range(start, end);
					break;
				}
			}

			return quiz;

			void SaveQuizHeader()
			{
				if (quiz != null && 
					quiz.HeaderContent == null && 
					quizHeaderStartPos != 0)
				{
					quiz.HeaderContent =
						doc.Range(quizHeaderStartPos, paragraph.Range.Start);
				}
				quizHeaderStartPos = 0;
			}

			void SaveGroupHeader()
			{
				if (group != null && 
					group.HeaderContent == null &&
					groupHeaderStartPos != 0)
				{
					group.HeaderContent =
						doc.Range(groupHeaderStartPos, paragraph.Range.Start);
				}
				groupHeaderStartPos = 0;
			}

			void SaveQuestionHeader()
			{
				if (question != null && 
					question.HeaderContent == null &&
					questionHeaderStartPos != 0)
				{
					question.HeaderContent =
						doc.Range(questionHeaderStartPos, paragraph.Range.Start);
				}
				questionHeaderStartPos = 0;
			}

			void SaveQuestionFooter()
			{
				if (question != null && 
					question.FooterContent == null &&
					questionFooterStartPos != 0 &&
					questionFooterStartPos < paragraph.Range.Start)
				{
					question.FooterContent =
						doc.Range(questionFooterStartPos, paragraph.Range.Start);
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

		public void LogQuiz(QuizDocument quiz)
		{
			this.logger.LogNewLine();
			this.logger.Log($"Parsed quiz document (from the input MS Word file):");
			this.logger.Log($" - LangCode: {quiz.LangCode}");
			this.logger.Log($" - VariantsToGenerate: {quiz.VariantsToGenerate}");
			this.logger.Log($" - TotalAvailableQuestions: {quiz.TotalAvailableQuestions}");
			this.logger.Log($" - AnswersPerQuestion: {quiz.AnswersPerQuestion}");
			string quizHeaderText = TruncateString(quiz.HeaderContent.Text);
			this.logger.Log($"Quiz header: {quizHeaderText}", 1);
			this.logger.Log($"Question groups = {quiz.QuestionGroups.Count}", 1);
			for (int groupIndex = 0; groupIndex < quiz.QuestionGroups.Count; groupIndex++)
			{
				this.logger.Log($"[Question Group #{groupIndex+1}]", 1);
				QuizQuestionGroup group = quiz.QuestionGroups[groupIndex];
				string groupHeaderText = TruncateString(group.HeaderContent?.Text);
				this.logger.Log($"Group header: {groupHeaderText}", 2);
				this.logger.Log($"Questions = {group.Questions.Count}", 2);
				for (int questionIndex = 0; questionIndex < group.Questions.Count; questionIndex++)
				{
					this.logger.Log($"[Question #{questionIndex+1}]", 2);
					QuizQuestion question = group.Questions[questionIndex];
					string questionContent = TruncateString(question.HeaderContent?.Text);
					this.logger.Log($"Question content: {questionContent}", 3);
					this.logger.Log($"Answers = {question.Answers.Count}", 3);
					foreach (var answer in question.Answers)
					{
						string prefix = answer.IsCorrect ? "Correct answer" : "Wrong answer";
						string answerText = TruncateString(answer.Content.Text);
						this.logger.Log($"{prefix}: {answerText}", 4);
					}
					string questionFooterText = TruncateString(question.FooterContent?.Text);
					if (questionFooterText == "")
						questionFooterText = "(empty)";
					this.logger.Log($"Question footer: {questionFooterText}", 3);
				}
			}
			string quizFooterText = TruncateString(quiz.FooterContent?.Text);
			this.logger.Log($"Quiz footer: {quizFooterText}", 1);
		}
	}
}
