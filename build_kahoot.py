import os
import sys
import openpyxl
import fnmatch

START_ROW = 9   # First question on this line
COLS = {'question':'B', 'ans1':'C', 'ans2':'D', 'ans3':'E', 'ans4':'F', 'time':'G', 'true':'H'}
TIME = 120

kahootFile = 'KahootQuizTemplate.xlsx'

class KahootQuestion:
    question = ""
    answers = ["", "", "", ""]
    trueAnswers = ""

    def __init__(self, question, answers, trueAnswers):
        self.question = question
        for i, answer in enumerate(answers):
            self.answers[i] = answer
        self.trueAnswers = trueAnswers

def extractQuestions(questionFiles):
    questions = []
    # Process each question file
    for questionFile in questionFiles:
        question = ""
        answers = []
        true = ""
        qfile = open(questionFile)
        lines = qfile.readlines()
        for i, line in enumerate(lines):
            # Sort line content
            if line.startswith('[Question'):
                question = lines[i+1]
            elif line.startswith('[Answer'):
                answers.append(lines[i+1])
            elif line.startswith('[True'):
                true = lines[i+1]
        
        qfile.close()
        questions.append(KahootQuestion(question, answers, true))
    
    return questions

# Make method for reading files??
def buildKahootSheet(questions):
    kahootWb = openpyxl.load_workbook(kahootFile)
    kahootSheet = kahootWb['Sheet1']
    rowToWrite = START_ROW
    # Write to kahoot excel file
    for question in questions:
        row = str(rowToWrite)
        questionCell = COLS['question'] + row
        kahootSheet[questionCell] = question.question
        ans1Cell = COLS['ans1'] + row
        kahootSheet[ans1Cell] = question.answers[0]
        ans2Cell = COLS['ans2'] + row
        kahootSheet[ans2Cell] = question.answers[1]
        ans3Cell = COLS['ans3'] + row
        kahootSheet[ans3Cell] = question.answers[2]
        ans4Cell = COLS['ans4'] + row
        kahootSheet[ans4Cell] = question.answers[3]
        timeCell = COLS['time'] + row
        kahootSheet[timeCell] = TIME
        trueCell = COLS['true'] + row
        kahootSheet[trueCell] = question.trueAnswers

        rowToWrite += 1

    kahootWb.save(kahootFile)

if __name__ == "__main__":
    # Find all question files
    questionFiles = []
    questionData = []

    for root, dirs, files in os.walk("."):
        for file in fnmatch.filter(files, "*.txt"):
            # Each file is a question
            questionFiles.append(os.path.join(root, file))
    
    questionData = extractQuestions(questionFiles)
    buildKahootSheet(questionData)
