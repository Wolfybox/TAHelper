import xlrd


class StudentGradeHelper:
    def __init__(self, my_stu_dir, rc_dir, disc_dir, me_dir, fa_dir):
        self.my_stu = self.read_my_student(my_stu_dir)
        self.rain_clr = self.read_rain_clr(rc_dir)
        self.disc = self.read_discuss_grade(disc_dir)
        self.mooc_exam = self.read_mooc_exam(me_dir)
        self.final_exam = self.read_final_exam(fa_dir)
        print('-' * 50)
        print(f'我负责的学生共有：{len(self.my_stu)}人\t其中有：')
        print(f'{len(self.rain_clr)}人 加入了雨课堂并有期末成绩')
        print(f'{len(self.disc)}人 参与了MOOC讨论')
        print(f'{len(self.mooc_exam)}人 参加了MOOC测验')
        print(f'{len(self.final_exam)}人 参加了期末考试')

    def read_spoc_student(self, spoc_dir):
        spoc_data = xlrd.open_workbook(spoc_dir)
        spoc_sheet = spoc_data.sheet_by_index(0)
        # print(f'Total Students in Spoc: {spoc_sheet.nrows}')
        spoc_stu_list = []
        for row_index in range(1, spoc_sheet.nrows):
            row_ele_list = spoc_sheet.row_values(row_index)
            nickname, real_name, stu_no, college, group_name = row_ele_list
            spoc_stu_list.append(real_name)
        return spoc_stu_list

    def read_my_student(self, stu_info_dir):
        stu_data = xlrd.open_workbook(stu_info_dir)
        stu_sheet = stu_data.sheet_by_index(0)
        print(f'Total Students: {stu_sheet.nrows}')
        target_stu_dict = {}
        for row_index in range(stu_sheet.nrows):
            row_ele_list = stu_sheet.row_values(row_index)
            row_id, stu_id, stu_name, stu_acade, _, stu_class, _, _, ta_name = row_ele_list
            target_stu_dict[stu_name] = stu_id
        return target_stu_dict

    def read_discuss_grade(self, disc_dir):
        disc_data = xlrd.open_workbook(disc_dir)
        disc_sheet = disc_data.sheet_by_index(0)
        # print(f'Total discuss students:{disc_sheet.nrows}')
        grade_dict = {}
        for row_index in range(1, disc_sheet.nrows):
            row_ele_list = disc_sheet.row_values(row_index)
            stu_name, disc_grade = row_ele_list[1], row_ele_list[8]
            if stu_name in self.my_stu.keys():
                grade_dict[stu_name] = float(disc_grade)
        return grade_dict

    def read_mooc_exam(self, mooc_exam_dir):
        mooc_data = xlrd.open_workbook(mooc_exam_dir)
        mooc_sheet = mooc_data.sheet_by_index(0)
        # print(f'Total mooc exam students:{mooc_sheet.nrows}')
        grade_dict = {}
        for row_index in range(mooc_sheet.nrows):
            row_ele_list = mooc_sheet.row_values(row_index)
            stu_name, grades = row_ele_list[1], row_ele_list[4:]
            if stu_name in self.my_stu.keys():
                grade_dict[stu_name] = grades
        return grade_dict

    def read_rain_clr(self, rain_clr_dir):
        rain_data = xlrd.open_workbook(rain_clr_dir)
        rain_sheet = rain_data.sheet_by_index(0)
        # print(f'Total rain clr students:{rain_sheet.nrows}')
        rain_grad_dict = {}
        for row_index in range(2, rain_sheet.nrows):
            row_ele_list = rain_sheet.row_values(row_index)
            stu_name, grades = row_ele_list[1], row_ele_list[10:]
            if stu_name in self.my_stu.keys():
                rain_grad_dict[stu_name] = grades
        return rain_grad_dict

    def read_final_exam(self, fa_dir):
        fa_data = xlrd.open_workbook(fa_dir)
        fa_sheet = fa_data.sheet_by_index(0)
        fa_grad_dict = {}
        for row_index in range(2, fa_sheet.nrows):
            row_ele_list = fa_sheet.row_values(row_index)
            stu_name, grades = row_ele_list[2], row_ele_list[10]
            if stu_name in self.my_stu.keys():
                fa_grad_dict[stu_name] = grades
        return fa_grad_dict

    def check_miss_disc_stu(self):
        target_stu_list = set(self.my_stu.keys())
        dis_stu_list = set(self.disc.keys())
        diff = list(target_stu_list - dis_stu_list)
        print('-' * 50)
        print('未参与讨论的学生为：')
        diff = ' | '.join(diff)
        print(diff)

    def check_unfinished_mooc_exam(self):
        res_stu = {}
        for stu_name, stu_grade in self.mooc_exam.items():
            if '-' in stu_grade:
                res_stu[stu_name] = []
            for gi, g in enumerate(stu_grade):
                if g == '-':
                    res_stu[stu_name].append(str(gi + 1))

        print('-' * 50)
        print('未完成MOOC所有测验的学生：')
        for stuname, unf_chap in res_stu.items():
            print(f'{stuname} 未完成章节：{"-".join(unf_chap)}')


if __name__ == "__main__":
    pass
    # disc_dir = r'data\spoc\20210125-20210426大学计算机—计算思维导论MOOC成绩.xls'
    # read_discuss_grade(disc_dir)

    # mooc_exam_dir = r'data\spoc\20210220-202104262021春大学计算机SPOC学生单项成绩汇总.xls'
    # read_mooc_exam(mooc_exam_dir)

    # rain_clr_dir = r'data\rain_classroom\“大学计算机I-2021春-20级自动化5班,20级自动化6班,20级自动化7班,20级自动化8班”成绩单202104261208.xlsx'
    # read_rain_clr(rain_clr_dir)
