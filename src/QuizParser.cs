using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;

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
			var quizDocument = new QuizDocument();
			quizDocument.Settings = new QuizSettings()
			{
				// Initially, assign default settings
				VariantsToGenerate = 5,
				QuestionsCount = 5,
				AnswersPerQuestion = 4
			};
			quizDocument.QuestionGroups = new List<QuizQuestionGroup>();
			QuizQuestionGroup currentGroup = null;
			QuizQuestion currentQuestion = null;
			int quizHeaderStartPos = 0;

			foreach (Word.Paragraph paragraph in doc.Paragraphs)
			{
				var text = paragraph.Range.Text.Trim();

				if (text.StartsWith(QuizStartTag))
				{
					var quizSettings = ParseSettings(text, QuizStartTag);
					quizDocument.Settings = MergeSettings(quizDocument.Settings, quizSettings);
					quizHeaderStartPos = paragraph.Range.End;
				}
				else if (text.StartsWith(QuestionGroupTag))
				{
					if (quizDocument.ContentBeforeQuestions == null)
					{
						quizDocument.ContentBeforeQuestions =
							doc.Range(quizHeaderStartPos, paragraph.Range.Start - 1);
					}
					var groupSettings = ParseSettings(text, QuestionGroupTag);
					currentGroup = new QuizQuestionGroup
					{
						Settings = MergeSettings(quizDocument.Settings, groupSettings),
						Questions = new List<QuizQuestion>()
					};
					quizDocument.QuestionGroups.Add(currentGroup);
				}
				else if (text.StartsWith(QuestionTag))
				{
					currentQuestion = new QuizQuestion
					{
						Answers = new List<QuestionAnswer>()
					};
					if (currentGroup == null)
					{
						currentGroup = new QuizQuestionGroup
						{
							Settings = quizDocument.Settings
						};
					}
					currentGroup.Questions.Add(currentQuestion);
					if (quizDocument.ContentBeforeQuestions == null)
					{
						quizDocument.ContentBeforeQuestions = 
							doc.Range(quizHeaderStartPos, paragraph.Range.Start - 1);
					}
				}
				else if (text.StartsWith(CorrectAnswerTag) || text.StartsWith(WrongAnswerTag))
				{
					if (currentQuestion == null)
					{
						currentQuestion = new QuizQuestion
						{
							Answers = new List<QuestionAnswer>()
						};
					}
					int answerStartRange = 0;
					if (text.StartsWith(CorrectAnswerTag))
						answerStartRange = CorrectAnswerTag.Length;
					else
						answerStartRange = WrongAnswerTag.Length;
					if (text.Length > answerStartRange && text[answerStartRange] == ' ')
						answerStartRange++;
					currentQuestion.Answers.Add(new QuestionAnswer
					{
						Content = doc.Range(
							paragraph.Range.Start + answerStartRange, 
							paragraph.Range.End),
						IsCorrect = text.StartsWith(CorrectAnswerTag)
					});
				}
				else if (text.StartsWith(QuizEndTag))
				{
					// Take all following paragraphs to the end of the document
					var start = paragraph.Range.End;
					var end = doc.Content.End;
					quizDocument.ContentAfterQuestions = doc.Range(start, end);
					break;
				}
			}

			return quizDocument;
		}

		private static QuizSettings MergeSettings(
			QuizSettings parentSettings, QuizSettings childSettings)
		{
			return new QuizSettings
			{
				VariantsToGenerate = childSettings.VariantsToGenerate != 0 ? childSettings.VariantsToGenerate : parentSettings.VariantsToGenerate,
				QuestionsCount = childSettings.QuestionsCount != 0 ? childSettings.QuestionsCount : parentSettings.QuestionsCount,
				AnswersPerQuestion = childSettings.AnswersPerQuestion != 0 ? childSettings.AnswersPerQuestion : parentSettings.AnswersPerQuestion
			};
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
	}
}
