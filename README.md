# SoftUni Quiz Generator

Generates randomized quizes, intended for paper-based skill testing.
  - Paper is used to reduce the chances for the students to cheat.
The generator app runs in Windows. It was written in C# with Windows Forms as GUI framework.

The quiz input (source code) comes from and MS Word document, which holds:
  - quiz header text
  - question groups
    - question group header text
    - questions
      - question text content
      - answers
        - correct / wrong answer
  - quiz footer text

The output is a set of MS Word documents:
  - Variant1.docx, Variant2.docx, ... - randomized quiz variant
  - Answers.docs - answer sheet for each variant

The quiz generator takes a randomized subset of questions from each question group.
It takes randomized subset of {maxAnswers+1} answers for each question: several wrong + 1 correct answer.
