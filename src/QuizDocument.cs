using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizDocument
	{
        public QuizSettings Settings { get; set; }
        public Word.Range ContentBeforeQuestions { get; set; }
        public List<QuizQuestionGroup> QuestionGroups { get; set; }
        public Word.Range ContentAfterQuestions { get; set; }
    }
}
