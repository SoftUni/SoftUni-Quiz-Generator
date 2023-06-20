using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class RandomizedQuiz
	{
		public Word.Range HeaderContent { get; set; }

		public List<QuizQuestionGroup> QuestionGroups { get; set; }

		public Word.Range FooterContent { get; set; }

		public IEnumerable<QuizQuestion> AllQuestions
			=> QuestionGroups.SelectMany(g => g.Questions);

		public static RandomizedQuiz GenerateFromQuizData(QuizDocument quizData)
		{
			// Clone the quiz header, question groups and footer
			RandomizedQuiz randQuiz = new RandomizedQuiz();
			randQuiz.HeaderContent = quizData.HeaderContent;
			randQuiz.FooterContent = quizData.FooterContent;
			randQuiz.QuestionGroups = new List<QuizQuestionGroup>();

			int questionGroupIndex = 1;
			foreach (var questionGroupData in quizData.QuestionGroups)
			{
				// Copy question groups from quiz data to randQuiz
				questionGroupIndex++;
				var randQuestionGroup = new QuizQuestionGroup();
				randQuiz.QuestionGroups.Add(randQuestionGroup);
				randQuestionGroup.HeaderContent = questionGroupData.HeaderContent;
				randQuestionGroup.SkipHeader = questionGroupData.SkipHeader;
				randQuestionGroup.Questions = new List<QuizQuestion>();

				// Check if QuestionsToGenerate is valid number
				if (questionGroupData.QuestionsToGenerate > questionGroupData.Questions.Count)
				{
					throw new Exception("QuestionsToGenerate > questions in group " +
						$"#{questionGroupIndex}` - {questionGroupData.HeaderContent?.Text}`");
				}

				// Generate a randomized subset of questions from the current group
				List<int> randomQuestionIndexes = 
					Enumerable.Range(0, questionGroupData.Questions.Count).ToList();
				randomQuestionIndexes = RandomizeList(randomQuestionIndexes);
				randomQuestionIndexes = randomQuestionIndexes.Take(
					questionGroupData.QuestionsToGenerate).ToList();
				foreach (int randQuestionIndex in randomQuestionIndexes)
				{
					var questionData = questionGroupData.Questions[randQuestionIndex];
					QuizQuestion randQuestion = new QuizQuestion();
					randQuestionGroup.Questions.Add(randQuestion);
					randQuestion.HeaderContent = questionData.HeaderContent;
					randQuestion.FooterContent = questionData.FooterContent;

					// Generate a randomized subset of answers (1 correct + several wrong)
					var correctAnswers = RandomizeList(questionData.CorrectAnswers);
					var wrongAnswers = RandomizeList(questionData.WrongAnswers);

					// Check if CorrectAnswers is valid number
					if (correctAnswers.Count == 0)
					{
						throw new Exception($"Question `{randQuestion.HeaderContent.Text}` --> " +
							$"at least 1 correct answer is required!");
					}

					// Check if WrongAnswers is valid number
					if (wrongAnswers.Count() == 0 ||
						wrongAnswers.Count() + 1 < questionGroupData.AnswersPerQuestion)
					{
						throw new Exception($"Question `{randQuestion.HeaderContent.Text}` --> " +
							$"wrong answers are less than required!");
					}

					List<QuestionAnswer> randAnswers =
						wrongAnswers.Take(questionGroupData.AnswersPerQuestion - 1)
						.Append(correctAnswers.First())
						.ToList();
					randQuestion.Answers = RandomizeList(randAnswers);
				}
			}

			return randQuiz;
		}

		private static List<T> RandomizeList<T>(IEnumerable<T> inputList)
		{
			// Randomize the list using Fisher-Yates shuffle algorithm
			List<T> list = new List<T>(inputList);
			Random rand = new Random();
			int lastIndex = list.Count;
			while (lastIndex > 1)
			{
				lastIndex--;
				int randIndex = rand.Next(lastIndex + 1);
				T value = list[randIndex];
				list[randIndex] = list[lastIndex];
				list[lastIndex] = value;
			}

			return list;
		}
	}
}
