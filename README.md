# TA Helper

Script for TA: Automatic seats-map generation &amp; grading

## Announcement

This project is only for educational purpose.

## Setup

```python
pip install xlrd openpyxl re xlwt
```

Prepare students' info and put them in the correct folder.  Considering data privacy, contact the repo author for templates of students' data. 

## Quick start

For seats-map generation:

 ```
python generate_seat.py --sign-info 签到查询.xls --stu-info 大学计算机B班（自动化5-8班）.xlsx --save-dir 周三56节T3201座位表.xlsx --template-dir 附件3：固定座位图模板.xlsx
 ```

For  automatic grading:

```
python grade_statistics.py --my-stu-info 我负责的同学.xlsx --rain-clr-grade 大学计算机I-2021春-20级自动化5班,20级自动化6班,20级自动化7班,20级自动化8班”成绩单202104280909.xlsx --disc-grade 20210125-20210428大学计算机—计算思维导论MOOC成绩.xls --mooc-exam 20210220-202104282021春大学计算机SPOC学生单项成绩汇总.xls --final-exam B班期末考试成绩.xlsx --out-dir 负责学生成绩汇总.xls
```
