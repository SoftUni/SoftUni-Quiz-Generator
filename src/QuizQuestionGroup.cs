using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizQuestionGroup
	{
		public QuizSettings Settings { get; set; }
		public Word.Range ContentBeforeQuestions { get; set; }
		public List<QuizQuestion> Questions { get; set; }
		public Word.Range ContentAfterQuestions { get; set; }
	}
}