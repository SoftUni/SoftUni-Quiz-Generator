using Word = Microsoft.Office.Interop.Word;

namespace SoftUniQuizGenerator
{
	class QuizDocument
	{
		public int VariantsToGenerate { get; set; }

		public int AnswersPerQuestion { get; set; }

		public int TotalAvailableQuestions
			=> QuestionGroups.Sum(g => g.Questions.Count);

		public int TotalQuestionsToGenerate
			=> QuestionGroups.Sum(g => g.QuestionsToGenerate);

		public Word.Range ContentBeforeQuestions { get; set; }

        public List<QuizQuestionGroup> QuestionGroups { get; set; }

        public Word.Range ContentAfterQuestions { get; set; }
    }
}
