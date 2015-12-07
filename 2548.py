import warnings
warnings.filterwarnings("ignore")
from openpyxl import load_workbook
from openpyxl import Workbook

def call_excel(year):
    sum_men_die = 0
    sum_men = 0
    sum_women_die = 0
    sum_women = 0
    wb = load_workbook(year + '.xlsx')
    sh = wb[year]
    for i in range(7, 83):
        sum_men += sh['D'+str(i)].value # Find sum of men who live.
        sum_women += sh['E'+str(i)].value # Find sum of women who live.
        sum_men_die += sh['G'+str(i)].value # Find sum of men who dead form suicide.
        sum_women_die += sh['H'+str(i)].value # Find sum of women who dead form suicide.
    total_people = sum_men + sum_women
    total_die = sum_men_die + sum_women_die
    mean_men_die = sum_men_die * 1000000 / sum_men
    mean_women_die = sum_women_die * 1000000 / sum_women
    mean_total_die = total_die * 1000000 / total_people
    return  mean_men_die, mean_women_die, mean_total_die

def write_excel():
    wb = Workbook()
    sh = wb.active
    sh.merge_cells('B1:D1')
    sh['B1'] = "จำนวนคนฆ่าตัวตายต่อประชากรล้านคน"
    sh['B2'] = "ผู้ชาย"
    sh['A2'] = "ปี(ชาย)"
    sh['D2'] = "ผู้หญิง"
    sh['C2'] = "ปี(หญิง)"
    sh['F2'] = "รวม"
    sh['E2'] = "ปี(รวม)"
    years = ["2548", "2549", "2550", "2551", "2552", "2553", "2554", "2555", "2556", "2557"]
    collum = 3
    for year in years: # Append the result values to each cell.
        mean_men_die, mean_women_die, mean_total_die = call_excel(year)
        print(mean_men_die, mean_women_die, mean_total_die)
        sh['A'+str(collum)] = year
        sh['C'+str(collum)] = year
        sh['E'+str(collum)] = year
        sh['B'+str(collum)] = "%.2f" % mean_men_die
        sh['D'+str(collum)] = "%.2f" % mean_women_die
        sh['F'+str(collum)] = "%.2f" % mean_total_die
        collum += 1
    wb.save('result.xlsx')

write_excel()
##print(call_excel("2555"))
