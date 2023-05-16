using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizQuestion
	{
		public Word.Range ContentBeforeAnswers { get; set; }
		public List<QuestionAnswer> Answers { get; set; }
		public Word.Range ContentAfterAnswers { get; set; }
	}
}