import csv

ros_path = "C:\\Users\\Pearson\\Downloads\\"
students = []

class Student:
    def __init__(self, student_data):
        self.first_name,self.last_name,self.middle_name,self.student_id,\
        self.email,self.phone,self.birthday,self.school,self.grade = student_data
        self.classes = []

    def enroll(self, course):
        self.classes.append(course)

    def is_enrolled(self, course):
        if course in self.classes:
            return True
        else:
            return False

    def drop_class(self, course):
        self.classes.remove(course)

    def print(self):
        slug = ','.join(self.first_name,self.last_name,self.middle_name,self.student_id,
                        self.email,self.phone,self.birthday,self.school,self.grade)
        updates = ','.join(self.classes)

        return slug+"("+updates+")"

class Course:
    def __init__(self, course_data):
        self.title,self.code,self.time_period,self.instructor = course_data
        self.students = []
        self.wait_list = []

    def enroll(self, student):
        self.students.append(student)

    def add_to_waitlist(self, student):
        self.wait_list.append(student)

    def check_enrollment(self, student):
        if student in self.students:
            return True
        else:
            return False

    def drop_student(self, student):
        self.students.remove(student)


student_list = {}
course_list = {}
with open(ros_path+"enrollment.csv") as csvfile:
    roster = csv.reader(csvfile)
    next(roster) # Skip
    for row in roster:
        if row[3] not in student_list.keys():
            student = Student(row[0:9])
            student_list[student.student_id] = student
        else:
            student = student_list[row[3]]
        if row[10] not in course_list.keys():
            course = Course(row[9:13])
            course_list[course.code] = course
        else:
            course = course_list[row[10]]

        status = row[13]
        if status == "ENROLLED":
            course.enroll(student)
            student.enroll(course)
        elif status == "WAIT":
            course.add_to_waitlist(student)
        else: # Student dropped
            if course.check_enrollment(student):
                course.drop_student(student)
            if student.is_enrolled(course):
                student.drop_class(course)


for course in course_list.values():
    print(course.title+": "+str(len(course.students)))
