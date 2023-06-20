using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizQuestion
	{
		public Word.Range HeaderContent { get; set; }

		public List<QuestionAnswer> Answers { get; set; }

		public IEnumerable<QuestionAnswer> CorrectAnswers =>
			this.Answers.Where(a => a.IsCorrect);

		public IEnumerable<QuestionAnswer> WrongAnswers =>
			this.Answers.Where(a => !a.IsCorrect);

		public Word.Range FooterContent { get; set; }
	}
}