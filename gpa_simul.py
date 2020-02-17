
from Tkinter import *
import ttk, tkMessageBox, re, math
from tkFileDialog import askopenfilename
from xlrd import open_workbook, XLRDError
from copy import deepcopy

letter_grades_table = {
    "A+": 4.1,
    "A": 4.0,
    "A-": 3.7,
    "B+": 3.3,
    "B": 3.0,
    "B-": 2.7,
    "C+": 2.3,
    "C": 2,
    "C-": 1.7,
    "D+": 1.3,
    "D": 1.5,
    "D-": 0.5,
    "F": 0.0,
    "S": 1.0,
    "": -1
}
numerical_grades_table = {
    4.1: "A+",
    4.0: "A",
    3.7: "A-",
    3.3: "B+",
    3.0: "B",
    2.7: "B-",
    2.3: "C+",
    2.0: "C",
    1.7: "C-",
    1.3: "D+",
    1.5: "D",
    0.5: "D-",
    0.0: "F",
    1.0: "S",
    -1: ""
}
letter_grades = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "F", "S"]

class GPA_Simulation_Tool(Frame):
    # we start to take this part from miniprojectone solution
    # Initializing the Class
    def __init__(self, parent, color):
        self.parent = parent
        self.color = color
        self.parent['bg'] = self.color
        self.courses_data = {}  # {course title: [course code, grade, credit]}
        self.original_courses_data = {}  # {course title: [course code, grade, credit]} to check if there is a change
        self.courses_by_semester = {}  # {semester: [course title, ] }
        self.combobox_grades = {}  # {combobox: course title}  the selected grade is fetched from the combobox itself.
        self.course_widgets = {}  # {course title: [label widget, combobox]}  the selected grade is fetched from the combobox itself
        self.edit_index = {}  # {Course title: index in edits listbox}
        self.taken_credits = 0
        self.gpa_numerator = 0
        self.displayed_year = 1
        self.immutable_original_gpa = 0
        self.initGUI()


    def font(self, size=15, bold="bold"):  # Default font size is 18
        return 'Calibri %d %s' % (size, bold)

    def initGUI(self):
        # HEADER ----------------------------------------------------------------
        Label(self.parent, text='SEHIR GPA SIMULATION TOOL V1.0', font=self.font(26), fg='white', bg='deepskyblue4')\
            .pack(fill=X)

        # Uploading data frame --------------------------------------------------
        uploading_data_frame = Frame(self.parent, bg=self.color)
        uploading_data_frame.pack(side=TOP, fill=X, pady=15)
        Label(uploading_data_frame, text="Courses File:", font=self.font(), bg=self.color).grid(row=1, column=1, padx=(320, 0))
        ttk.Button(uploading_data_frame, text="Choose", command=self.load_data).grid(row=1, column=2, padx=(40, 0))

        # Displaying semesters frame--------------------------------------------------
        displaying_semesters_frame = Frame(self.parent, bg=self.color)
        displaying_semesters_frame.pack(side=TOP, fill=X, padx=30)
        Button(displaying_semesters_frame, text="<<", height=12, command=self.previous_year).grid(row=1, rowspan=2, column=1, padx=(30, 0))
        # --- First semester frame ----------------------------------------------------
        self.left_semester = Label(displaying_semesters_frame, bg=self.color, font=self.font(15,""))
        self.left_semester.grid(row=0, column=2)
        self.lst_semester = Frame(displaying_semesters_frame, bg=self.color,
                                  highlightbackground="deepskyblue4", highlightthickness=2, width=400, height=200)
        self.lst_semester.grid(row=1, column=2, padx=20, sticky=N)
        """grid_propagate(0) since it is given a 0 flag,
            it prevents the frame from shrinking/swelling to wrap its content.
            If we were using pack, it would be pack_propagate(0)"""
        self.lst_semester.grid_propagate(0)
        # ------ courses will be added dynamically here when data is loaded -----------

        Button(displaying_semesters_frame, text=">>", height=12, command=self.next_year)\
            .grid(row=1, rowspan=2, column=4, padx=(12, 0))  # Primitive buttons, stay tuned
        self.right_semester = Label(displaying_semesters_frame, bg=self.color, font=self.font(15, ""))
        self.right_semester.grid(row=0, column=3)
        # --- Second semester frame ---------------------------------------------------
        self.second_semester = Frame(displaying_semesters_frame, bg=self.color,
                                     highlightbackground="deepskyblue4", highlightthickness=2, width=400, height=200)
        self.second_semester.grid(row=1, column=3, padx=(20,5), sticky=N)
        self.second_semester.grid_propagate(0)
        # ------ courses will be added dynamically here when data is loaded -----------

        # History frame
        history_frame = Frame(self.parent, bg=self.color)
        history_frame.pack(side=TOP, fill=X, pady=(20, 0), padx=(50,0))
        Label(history_frame, text="Edited Courses:", bg=self.color, font=self.font()).grid(row=1, column=1, sticky=W)
        self.history_listbox = Listbox(history_frame, height=8, width=80, font=self.font(12))
        self.history_listbox.grid(row=2, column=1, columnspan=2, rowspan=2)
        history_listbox_scrollbar = ttk.Scrollbar(history_frame, command=self.history_listbox.yview)
        self.history_listbox['xscrollcommand'] = history_listbox_scrollbar.set
        history_listbox_scrollbar.grid(row=2, rowspan=2, column=2, sticky=S+N+E)
        ttk.Button(history_frame, text="Remove", command=self.remove_change).grid(row=2, column=3, padx=(30, 0))  # ttk button just looks neat
        ttk.Button(history_frame, text="Edit", command=self.edit_change).grid(row=3, column=3, padx=(30, 0))

        # Edit GPA from a history (write your own history :D)
        self.edit_history_frame = Frame(self.parent, bg=self.color, height=25)
        self.edit_history_frame.pack_propagate(0)
        self.edit_history_frame.pack(side=TOP, fill=X)
        # ---- If a course is selected from the edit history, edit_history_frame will show its widgets

        # GPA data frame
        GPA_data_frame = Frame(self.parent, bg=self.color)
        GPA_data_frame.pack_propagate(0)
        GPA_data_frame.pack(side=TOP, fill=X, padx=(50,0),pady=20)
        Label(GPA_data_frame, font=self.font(), text="Current GPA:", bg=self.color).grid(row=0, column=0, padx=(20, 10))
        self.GPA_label = Label(GPA_data_frame, font=self.font(12), text="No data", bg=self.color)
        self.GPA_label.grid(row=0, column=1)
        Label(GPA_data_frame, font=self.font(), text="Original GPA:", bg=self.color).grid(row=0, column=2, padx=(80, 10))
        self.original_GPA = Label(GPA_data_frame, font=self.font(12), text="No data", bg=self.color)
        self.original_GPA.grid(row=0, column=3)
        Label(GPA_data_frame, font=self.font(), text="Change:", bg=self.color).grid(row=0, column=4,padx=(80, 10))
        self.GPA_change = Label(GPA_data_frame, font=self.font(12), text="No data", bg=self.color)
        self.GPA_change.grid(row=0, column=5)

    def edit_selected_grade(self, course_name):
        self.selected_course = Label(self.edit_history_frame, font=self.font(12, ""), text=course_name, bg=self.color)
        self.selected_course.grid(row=0, column=1, padx=(50, 10))
        self.new_grade = Entry(self.edit_history_frame, font=self.font(12, ""), width=3)
        self.new_grade.grid(row=0, column=2)
        ttk.Button(self.edit_history_frame, text="Save", command=self.save_edit).grid(row=0, column=3, padx=(50, 0))

    def letter2float(self, letter_grade):
        return letter_grades_table[letter_grade]

    def float2letter(self, numerical_grade):
        return numerical_grades_table[numerical_grade]

    def next_year(self):
        if self.displayed_year < 4:
            self.displayed_year += 1
            self.display_courses()

    def previous_year(self):
        if self.displayed_year > 1:
            self.displayed_year -= 1
            self.display_courses()

    def load_data(self):
        """ self.courses_by_semester = {semester: [course title, ] }
            self.courses_data = {course title: [course code, grade, credit] }"""
        # clearing display frames and data structures
        for frame in [self.lst_semester, self.second_semester]:
            for widget in frame.winfo_children():
                widget.destroy()
        self.courses_by_semester = {}
        try:
            data_sheet = open_workbook(askopenfilename()).sheet_by_index(0)
        except XLRDError:
            tkMessageBox.showerror("Invalid data file", "The file you have chosen is not a valid data file")
            return
        except IOError:
            tkMessageBox.showerror("Invalid data file", "Please select the data file")
            return
        for row in range(1, data_sheet.nrows):
            # The semester is in the 5th column in the sample datafile
            semester = int(re.findall(r'\d+',data_sheet.cell_value(row, 4))[0])
            course_code = data_sheet.cell_value(row, 0)
            course_title = data_sheet.cell_value(row, 1)
            grade = self.letter2float(data_sheet.cell_value(row, 2))
            credit = data_sheet.cell_value(row, 3)
            if grade > -1:
                self.taken_credits += credit
                self.gpa_numerator += grade*credit
            self.courses_by_semester.setdefault(semester, [])
            self.courses_by_semester[semester].append(course_title)
            self.courses_data[course_title] = [course_code, grade, credit]
        gpa = self.calculateGPA()
        self.immutable_original_gpa = gpa
        self.original_GPA['text'] = "%.2f" % gpa
        self.GPA_label['text'] = "%.2f" % gpa
        self.original_courses_data = deepcopy(self.courses_data)
        self.display_courses()

    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.destroy()

    def display_courses(self):
        # Clearing the semesters courses frames
        self.clear_frame(self.lst_semester)
        self.clear_frame(self.second_semester)
        """self.courses_by_semester = {semester: [course title, ] }
           self.courses_data = {course title: [course code, grade, credit] }
           self.course_widgets = {}  # {course title: [label, combobox]}  """
        lst_semester_row = 0
        second_semester_row = 0

        for semester in self.courses_by_semester:
            if self.displayed_year == math.ceil(semester/2.0):
                if semester % 2 == 1:
                    self.left_semester['text'] = "Semester %d" % semester
                    for course_title in self.courses_by_semester[semester]:
                        course_label = Label(self.lst_semester, text=course_title, bg=self.color, font=self.font(10))
                        course_label.grid(row=lst_semester_row, column=0, sticky=W)
                        grade = self.float2letter(self.courses_data[course_title][1])
                        grade_combobox = ttk.Combobox(self.lst_semester, font=self.font(10),width=3, values=letter_grades)
                        grade_combobox.grid(row=lst_semester_row, column=1, padx=10)
                        grade_combobox.set(grade)
                        grade_combobox.bind('<<ComboboxSelected>>', self.on_changing_grade)
                        self.combobox_grades[grade_combobox] = course_title
                        self.course_widgets[course_title] = [course_label, grade_combobox]
                        Label(self.lst_semester, bg=self.color, text=self.courses_data[course_title][2]).grid(row=lst_semester_row, column=2)
                        lst_semester_row += 1
                else:  # i.e: semester % 2 == 0
                    self.right_semester['text'] = "Semester %d" % semester
                    for course_title in self.courses_by_semester[semester]:
                        course_label = Label(self.second_semester, text=course_title, bg=self.color, font=self.font(10))
                        course_label.grid(row=second_semester_row, column=0, sticky=W)
                        grade = self.float2letter(self.courses_data[course_title][1])
                        grade_combobox = ttk.Combobox(self.second_semester, font=self.font(10), width=3, values=letter_grades)
                        grade_combobox.grid(row=second_semester_row, column=1, padx=10)
                        grade_combobox.set(grade)
                        self.combobox_grades[grade_combobox] = course_title
                        self.course_widgets[course_title] = [course_label, grade_combobox]
                        grade_combobox.bind('<<ComboboxSelected>>', self.on_changing_grade)
                        Label(self.second_semester, bg=self.color, text=self.courses_data[course_title][2]).grid(row=second_semester_row,column=2)
                        second_semester_row += 1

    def update_index_dictionary(self, shift, index):
        for course in self.edit_index:
            if self.edit_index[course] > index:
                self.edit_index[course] += shift


    def on_changing_grade(self, event):
        """self.courses_by_semester = {semester: [course title, ]}
           self.courses_data = {course title: [course code, grade, credit] }
           self.combobox_grades = {combobox : course title} """
        # Fetching course data from self.combobox_grades dictionary
        course_title = self.combobox_grades[event.widget]
        credit = self.courses_data[course_title][2]
        old_grade = self.courses_data[course_title][1]
        new_letter_grade = event.widget.get()
        new_grade = float(self.letter2float(new_letter_grade))
        if old_grade == new_grade: return  # Avoiding repeated consecutive choices from the combobox

        # _________________ Adding changes to the edit-history listbox ______________________________________
        self.edit_index.setdefault(course_title, self.history_listbox.size())
        index = self.edit_index[course_title]
        if self.original_courses_data[course_title][1] == new_grade:
            self.course_widgets[course_title][0]['bg'] = self.color
            self.history_listbox.delete(index)
            self.update_index_dictionary(-1, index)

        # Changing the course label background to yellow
        else:
            self.course_widgets[course_title][0].configure(bg="yellow")
            if index != self.history_listbox.size():
                self.history_listbox.delete(index)
            self.history_listbox.insert(index, self.write_edit_line(course_title, new_letter_grade))
        self.update_index_dictionary(1, index)
        # Calculating the new GPA
        try:
            self.gpa_numerator = self.gpa_numerator + credit * (new_grade - old_grade)
        except TypeError:
            return
        new_gpa = self.calculateGPA()
        self.GPA_label['text'] = "%.2f" % new_gpa
        # Updating the data dictionary
        self.courses_data[course_title][1] = new_grade

    def write_edit_line(self, course_title, new_grade):
        original_grade = self.float2letter(self.original_courses_data[course_title][1])
        return course_title+" "*(50-len(course_title))+">>>"+" "*10+original_grade+" "*10+">>>"+" "*10+new_grade

    def remove_change(self):
        """self.courses_data = {course title: [course code, grade, credit]}
            course data row:
                course_title+" "*15+">>>"+" "*12+self.float2letter(old_grade)+" "*12+">>>"+" "*12+new_grade"""
        selected_course = re.split(r">>>", self.history_listbox.get(ACTIVE))
        course_title = selected_course[0].rstrip()
        old_grade = self.letter2float(selected_course[1].strip())  # The original grade before the selected change
        new_grade = self.letter2float(selected_course[2].strip())  # The modified grade
        if old_grade == self.original_courses_data[course_title][1]:
            self.course_widgets[course_title][0]['bg'] = self.color
        else:
            self.course_widgets[course_title][0]['bg'] = "yellow"
        self.course_widgets[course_title][1].set(self.float2letter(old_grade))
        # updating courses grades dictionary
        credit = self.courses_data[course_title][2]
        self.courses_data[course_title][1] = old_grade
        self.gpa_numerator = self.gpa_numerator + (old_grade - new_grade)*credit
        self.GPA_label['text'] = "%.2f" % self.calculateGPA()
        # Deleting the selected change from the history listbox
        self.history_listbox.delete(ACTIVE)


    def edit_change(self):
        """self.courses_data = {course title: [course code, grade, credit]}
            course data row:
                course_title+" "*15+">>>"+" "*12+self.float2letter(old_grade)+" "*12+">>>"+" "*12+new_grade"""
        selected_course = re.split(r">>>", self.history_listbox.get(ACTIVE))
        course_title = selected_course[0].rstrip()
        new_grade = self.letter2float(selected_course[2].strip())  # The last modified grade
        self.clear_frame(self.edit_history_frame)
        self.edit_selected_grade(course_title)

    def save_edit(self):
        """self.courses_data = {course title: [course code, grade, credit]}"""
        course_title = self.selected_course['text']
        old_grade = self.courses_data[course_title][1]
        credit = self.courses_data[course_title][2]
        new_grade = self.letter2float(self.new_grade.get())
        if new_grade == self.original_courses_data[course_title][1]:
            self.course_widgets[course_title][0]['bg'] = self.color
        else:
            self.course_widgets[course_title][0]['bg'] = "yellow"
        self.course_widgets[course_title][1].set(self.float2letter(new_grade))
        self.gpa_numerator = self.gpa_numerator + (new_grade - old_grade)*credit
        self.GPA_label['text'] = "%.2f" % self.calculateGPA()
        self.history_listbox.insert(self.history_listbox.index(ACTIVE), self.write_edit_line(course_title, self.float2letter(new_grade)))
        self.history_listbox.delete(ACTIVE)

    def calculateGPA(self):
        gpa = self.gpa_numerator/float(self.taken_credits)
        try:
            self.GPA_change['text'] = "%.2f%%" % (((gpa - self.immutable_original_gpa)) / 0.041)
        except ValueError:
            self.GPA_change['text'] = 0.0
        return gpa

if __name__ == '__main__':
    root = Tk()
    o = GPA_Simulation_Tool(root, 'honeydew')
    root.geometry('1050x650')
    root.mainloop()