'''
Program: PAYE Income Tax Calculator-Tax Year 16/17
Author: Ranbir Dixit
Date: 17/03/2016
Description: Calculates UK Net Salaries after deducting PAYE, Pension and Class 1 NIC for 2016/17 Tax Year and allows
the results to be written out to Microsoft-Excel for further analysis and charting
'''
import openpyxl
'''Define Tax Brackets and Rates'''
taxFreeBracket=11000.0
taxBracketBasicRateLower=11001.0
taxBracketBasicRateHigher=32000.0
taxBracketHigherRateLower=32001.0
taxBracketAdditionalRateLower=150000.0
taxBracketBasic=0.2
taxBracketHigher=0.4
taxBracketAdditional=0.45
'''Define Class 1 National Insurance Brackets and Rates'''
NIThreshold=8060.0
NIBasicRateUpperLimit=43000.0
NIBasicRate=0.12
NIAboveBasicRate=0.02
'''Calculate Tax Free Personal Allowance'''
def personalAllowance(salary):
    reductionfactor = (salary-100000)/2.0
    if salary>100000:
        personalallowance=taxFreeBracket-reductionfactor
        return max(0,personalallowance)
    else:
        return taxFreeBracket
'''Calculate Class 1 National Insurance'''
def nationalInsurance(salary):
    if salary <= NIThreshold:
        NInsurance = 0.0
        return NInsurance
    elif salary>NIThreshold and salary<=NIBasicRateUpperLimit:
        NInsurance=(salary-NIThreshold)*NIBasicRate
        return NInsurance
    elif salary>NIBasicRateUpperLimit:
        NInsuranceBasic=(NIBasicRateUpperLimit-NIThreshold)*NIBasicRate
        NIInsuranceAboveBasic=(salary-NIBasicRateUpperLimit)*NIAboveBasicRate
        NInsurance=NInsuranceBasic+NIInsuranceAboveBasic
        return NInsurance
'''Calculate netSalary: grossSalary - pension - NInsurance - incomeTax'''
def netSalary(grossAnnualSalary,pensionPercent):
    # grossAnnualSalary=float(raw_input("please enter gross annual salary:\n"))
    # pensionPercent=float(raw_input("please enter percentage of pension paid: \n "))
    pension=grossAnnualSalary*pensionPercent
    taxableSalary=grossAnnualSalary-pension
    taxableTotal=taxableSalary-personalAllowance(grossAnnualSalary)
    taxesBasic=taxBracketBasicRateHigher*taxBracketBasic #applies to last two cases
    if taxableTotal <= 0:
        totalIncomeTax = 0
        return grossAnnualSalary,pensionPercent,pension,nationalInsurance(grossAnnualSalary),totalIncomeTax,taxableSalary-nationalInsurance(grossAnnualSalary)
    elif taxableTotal > 0 and taxableTotal<taxBracketHigherRateLower:
        totalIncomeTax=taxableTotal*taxBracketBasic
        return grossAnnualSalary,pensionPercent,pension,nationalInsurance(grossAnnualSalary),totalIncomeTax,taxableSalary-totalIncomeTax-nationalInsurance(grossAnnualSalary)
    elif taxableTotal>=taxBracketHigherRateLower and taxableTotal<=taxBracketAdditionalRateLower:
        taxableHigher=taxableTotal-taxBracketBasicRateHigher
        taxesHigher=taxableHigher*taxBracketHigher
        totalIncomeTax=taxesBasic+taxesHigher
        return grossAnnualSalary,pensionPercent,pension,nationalInsurance(grossAnnualSalary),totalIncomeTax,taxableSalary-totalIncomeTax-nationalInsurance(grossAnnualSalary)
    else:
        taxesHigher=(taxBracketAdditionalRateLower-taxBracketBasicRateHigher)*taxBracketHigher
        taxableAdditional=taxableTotal-taxBracketAdditionalRateLower
        taxesAdditional=taxableAdditional*taxBracketAdditional
        totalIncomeTax=taxesBasic+taxesHigher+taxesAdditional
        return grossAnnualSalary,pensionPercent,pension,nationalInsurance(grossAnnualSalary),totalIncomeTax,taxableSalary-totalIncomeTax-nationalInsurance(grossAnnualSalary)
'''Write the file out to Microsoft-Excel using the openpyxl library'''
def write_results():
    pens=0.125
    wb=openpyxl.Workbook()
    sheet=wb.get_sheet_by_name('Sheet')
    sheet['A1']='Gross Salary'
    sheet['B1']='Pension Percent'
    sheet['C1']='Pension'
    sheet['D1']='National Insurance'
    sheet['E1']='Income Tax'
    sheet['F1']='Net Salary'
    sheet['G1']='Monthly Net Salary'
    i=2
    for sal in range(10000,100000,1000):
        aftertaxsalary=netSalary(sal,pens)
        #print aftertaxsalary[4]
        sheet['A'+str(i)] = aftertaxsalary[0]
        sheet['B'+str(i)] = aftertaxsalary[1]
        sheet['C'+str(i)] = aftertaxsalary[2]
        sheet['D'+str(i)] = aftertaxsalary[3]
        sheet['E'+str(i)] = aftertaxsalary[4]
        sheet['F'+str(i)] = aftertaxsalary[5]
        sheet['G'+str(i)] = aftertaxsalary[5]/12.0
        i+=1
    wb.save('PAYE Analysis_100000_5%Pension.xlsx')
write_results()


