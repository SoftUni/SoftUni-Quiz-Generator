using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizQuestionGroup
	{
		public int QuestionsToGenerate { get; set; }
		public int AnswersPerQuestion { get; set; }
		public bool SkipHeader { get; set; }
		public Word.Range ContentBeforeQuestions { get; set; }
		public List<QuizQuestion> Questions { get; set; }
	}
}