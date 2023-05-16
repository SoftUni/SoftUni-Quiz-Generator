using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuestionAnswer
	{
		public Word.Range Content { get; set; }
        public bool IsCorrect { get; set; }
    }
}
