import openpyxl 
from random import randint, choice
from persian_names import fullname_en
import radar
import string
import json


class Sheet:

    def __init__(self, questions_count=100):
        self.questions_count = questions_count

    def generate_random_answer_sheet(self):
        sheet = list()
        for i in range(self.questions_count):
            sheet.append(randint(0, 4))
        return json.dumps(sheet)

    def generate_random_name(self,gender="r"):
        random_name = fullname_en(gender)
        return random_name

    def generate_random_date(self):
        random_date = radar.random_date()
        return str(random_date)


    def generate_random_exam_name(self):
        random_letter = choice(string.ascii_letters).upper()
        random_numbers = randint(1000,9999)
        
        random_exam_name = random_letter + str(random_numbers)
        return random_exam_name

    def generate_random_duration(self):
        random_duration = randint(1, 7200)
        return random_duration

    def generate_random_ip(self):
        return '.'.join(
            str(randint(0, 255)) for _ in range(4)
        )

    def generate_random_list_of_ips(self):
        ip_count = randint(1, 10)
        ips = list()
        for i in range(ip_count):
            ips.append(Sheet().generate_random_ip())
        ips = json.dumps(ips)
        return ips


    def generate_online_question_solve_duration(self):
        durations = list()
        for i in range(self.questions_count):
            duration = randint(0, 15000)
            duration = duration / 100
            durations.append(duration)
        durations = json.dumps(durations)
        return durations

class Exam:
    def generate_random_face_to_face_student_exam(exam_name=None, exam_date=None):
        row = list()
        if exam_name == None:
            exam_name = Sheet().generate_random_exam_name()
        if exam_date == None:
            exam_date = Sheet().generate_random_date()
        row.append(exam_name)
        row.append(Sheet().generate_random_name())
        row.append(Sheet().generate_random_answer_sheet())
        row.append(exam_date)
        row.append(Sheet().generate_random_duration())
        return row


    def generate_face_to_face_exam(students_count=20, exam_name=None, exam_date=None):
        studemt_rows = list()
        if exam_name == None:
            exam_name = Sheet().generate_random_exam_name()
        if exam_date == None:
            exam_date = Sheet().generate_random_date()
        exam_answer_row = [exam_name, Sheet().generate_random_answer_sheet()]
        for i in range(students_count):
            studemt_rows.append(Exam.generate_random_face_to_face_student_exam(exam_name, exam_date))
        return studemt_rows, exam_answer_row


    def generate_random_online_student_exam(exam_name=None, exam_date=None):
        row = list()
        if exam_name == None:
            exam_name = Sheet().generate_random_exam_name()
        if exam_date == None:
            exam_date = Sheet().generate_random_date()
        row.append(exam_name)
        row.append(Sheet().generate_random_name())
        row.append(Sheet().generate_random_answer_sheet())
        row.append(exam_date)
        row.append(Sheet().generate_online_question_solve_duration())
        row.append(Sheet().generate_random_list_of_ips())
        return row


    def generate_online_exam(students_count=20, exam_name=None, exam_date=None):
        studemt_rows = list()
        if exam_name == None:
            exam_name = Sheet().generate_random_exam_name()
        if exam_date == None:
            exam_date = Sheet().generate_random_date()
        exam_answer_row = [exam_name, Sheet().generate_random_answer_sheet()]
        for i in range(students_count):
            studemt_rows.append(Exam.generate_random_online_student_exam(exam_name, exam_date))
        return studemt_rows, exam_answer_row


class Excel:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.sheet = self.wb.active

    def insert_rows(self, rows):
        for row_index in range(len(rows)):
            for value_index in range(len(rows[row_index])):
                self.sheet.cell(row = row_index + 1, column = value_index + 1).value = rows[row_index][value_index]
        return True
        
    def save_file(self, path):
        self.wb.save(path)
        return path

def generate_database(exam_count=5, students_count=20, file_name="sample.xlsx"):
    """
    colnames are ExamName, StudentName, AnswerSheet, Date, Duration
    colnames are ExamName, StudentName, AnswerSheet, Date, DurationPerQuestion, IPs
    """

    face_to_face_student_excel = Excel()
    online_student_excel = Excel()
    exam_answers_excel = Excel()
    all_face_to_face_student_rows = list()
    all_online_student_rows = list()
    all_exam_rows = list()
    for i in range(exam_count):
        exam_name = Sheet().generate_random_exam_name()
        exam_date = Sheet().generate_random_date()
        student_rows, exam_answer_row = Exam.generate_face_to_face_exam(students_count, exam_name, exam_date)
        all_face_to_face_student_rows += student_rows

        student_rows, exam_answer_row = Exam.generate_online_exam(students_count, exam_name, exam_date)
        all_online_student_rows += student_rows
        
        all_exam_rows.append(exam_answer_row) 


    face_to_face_student_excel.insert_rows(all_face_to_face_student_rows)
    online_student_excel.insert_rows(all_online_student_rows)
    exam_answers_excel.insert_rows(all_exam_rows)
    face_to_face_student_excel.save_file("FTFStudents.xlsx")
    online_student_excel.save_file("OStudents.xlsx")
    exam_answers_excel.save_file("Exams.xlsx")


generate_database()