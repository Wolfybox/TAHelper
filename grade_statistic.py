import argparse

import xlwt
from utils import StudentGradeHelper

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('--my-stu-info', type=str, default=r'data/general/我负责的同学.xlsx', help='我负责的学生（自行修改）')
    argparser.add_argument('--rain-clr-grade', type=str,
                           default=r'data/rain_classroom/“大学计算机I-2021春-20级自动化5班,20级自动化6班,20级自动化7班,20级自动化8班”成绩单202104280909.xlsx',
                           help='雨课堂成绩单')
    argparser.add_argument('--disc-grade', type=str, default='data/spoc/20210125-20210428大学计算机—计算思维导论MOOC成绩.xls',
                           help='讨论成绩表')
    argparser.add_argument('--mooc-exam', type=str, default='data/spoc/20210220-202104282021春大学计算机SPOC学生单项成绩汇总.xls',
                           help='MOOC测验成绩表')
    argparser.add_argument('--final-exam', type=str, default='data/general/B班期末考试成绩.xlsx', help='期末考试成绩表')
    argparser.add_argument('--out-dir', type=str, default=r'data/general/负责学生成绩汇总.xls', help='合并所有分数输出到一个文件中')
    config = argparser.parse_args()

    stu_grade_helper = StudentGradeHelper(my_stu_dir=config.my_stu_info, rc_dir=config.rain_clr_grade,
                                          disc_dir=config.disc_grade, me_dir=config.mooc_exam, fa_dir=config.final_exam)

    stu_grade_helper.check_miss_disc_stu()
    stu_grade_helper.check_unfinished_mooc_exam()

    # # write grades to xls file
    wb = xlwt.Workbook()
    ws = wb.add_sheet('成绩汇总')
    # writing top columns
    col_names = ['姓名', '学号', '雨课堂-课堂成绩', '雨课堂-作业成绩', 'MOOC讨论成绩', 'MOOC测验总成绩', '期末考试成绩']
    for i, coln in enumerate(col_names):
        ws.write(0, i, coln)

    for i, stu_item in enumerate(stu_grade_helper.my_stu.items()):
        stu_name, stu_no = str(stu_item[0]), str(int(stu_item[1]))
        print(f'\rprocessing student {i + 1}/{len(stu_grade_helper.my_stu)} {stu_name} {stu_no}', end='')
        rc_cls, rc_hw = float(stu_grade_helper.rain_clr[stu_name][0]), float(stu_grade_helper.rain_clr[stu_name][1])
        try:
            dis_grade = float(stu_grade_helper.disc[stu_name])
        except Exception:
            dis_grade = .0
        mooc_grade = stu_grade_helper.mooc_exam[stu_name]
        sum_mooc_grade = sum([float(g) for g in mooc_grade if g != '-'])
        fa_grade = float(stu_grade_helper.final_exam[stu_name])
        row_elements = [stu_name, stu_no, rc_cls, rc_hw, dis_grade, sum_mooc_grade, fa_grade]
        for j, ele in enumerate(row_elements):
            ws.write(i + 1, j, ele)

    wb.save(config.out_dir)
