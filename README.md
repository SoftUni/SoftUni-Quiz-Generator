# SoftUni Quiz Generator

Generates **randomized quizes**, intended for **paper-based skill testing**.
  - Paper is used to reduce the chances for the students to cheat (avoid plagiarism)
  - The quiz generator app runs in **Windows**. It was written in **C#** with **Windows Forms** as GUI framework.

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/87fa0165-7546-4809-baff-2df306947786)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/cc21f9f0-3dbb-44a7-a974-6582c68bb25f)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/365297af-9cd0-4aec-bfab-e9348be9a7a2)

## Quiz Input

The **quiz input** (source code) comes from and **MS Word document**, which holds:
  - Quiz header text
  - Question groups
    - Question group header text
    - Questions
      - Question text content
      - Answers
        - Correct / Wrong answer
  - Quiz footer text

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/55eb0432-6b84-4714-a87d-ffd25426c1f1)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/6f5f7671-b7dd-4279-8090-392546d52762)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/d39aaa10-143e-4ea1-8fb9-5528200fadb7)

## Output: A Set of Random Quizes + Answer Sheet

The output is a **set of MS Word documents**:
  - `quiz001.docx`, `quiz002.docx`, ... - randomized quiz variants
  - `answers.html` - answer sheet (the correct answers for each quiz variant)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/5766f497-1613-465f-b27f-9a9057ca0af5)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/74353996-cf73-4703-9667-48337615377f)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/c3aed2a5-ef52-4558-a8f1-ec1d73138d2a)

## How Does It Work?

The quiz generator takes a **randomized subset of questions** from each question group.
  - The order of the questions in each group is random
  - For each question, the quiz generator takes **1 correct** + **several wrong answers** (randomly ordered subset).

## How to Check the Answers?

### Step 1: Prepare the Answers Sheet

  - Copy/paste the **answer sheet** from `answers.html` to the answers sheet **MS Word template** `answers-template.docx`.
  
  - This is because the length of each answer letters row (in centimeters) should be the same in the generated quiz and in the answer sheet.

  - This is how the answer sheet may look like in MS Word:

    ![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/0be2798b-8231-4de5-b0b9-0e9222fc36f1)
    
  - This is how the answer sheet may look like printed on a sheet of paper:

    ![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/fcce427a-649c-49e3-9417-cc8aab42c1fd)

### Step 2: Students Fill Their Answers
  - **Print** the generated quizes on **paper** and give them to the students.
  - **Students fill letters only**: students should **fill the answers section** at the top of the page (fill only letters: `a` / `b` / `c` / `d` / ...).
  - **No digital devices**: ensure that nobody has access to any digital devices (e.g. laptop / smartphone / smart watch)

![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/decc24de-66fb-44aa-9369-7b3738049967)

### Step 3: Check Student's Answers
  - **Print the answer sheet** from MS Word and use it to **check student's answers**.
  - Use can fold it at the correct row, depending on the variant you check:

    ![image](https://github.com/SoftUni/SoftUni-Quiz-Generator/assets/1689586/c41d4467-3d98-476a-8226-57326e2784ff)

    
