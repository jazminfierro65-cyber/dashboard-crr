from pathlib import Path

import openpyxl

files = [
    Path("/Users/jazz/Downloads/plantilla_carga_web_crr.xlsx"),
    Path("/Users/jazz/Documents/New project/outputs/crr_dashboard/plantilla_carga_web_crr.xlsx"),
]


def delay_formula(row):
    return (
        f'=IFERROR(IF(OR(Q{row}="",S{row}=""),"",'
        f'IF(IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row}))>=IF(ISNUMBER(Q{row}),MOD(Q{row},1),TIMEVALUE(Q{row})),'
        f'(IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row}))-IF(ISNUMBER(Q{row}),MOD(Q{row},1),TIMEVALUE(Q{row})))*1440,'
        f'(IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row}))+1-IF(ISNUMBER(Q{row}),MOD(Q{row},1),TIMEVALUE(Q{row})))*1440)),"")'
    )


def duration_formula(row):
    return (
        f'=IFERROR(IF(OR(S{row}="",T{row}=""),"",'
        f'IF(IF(ISNUMBER(T{row}),MOD(T{row},1),TIMEVALUE(T{row}))>=IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row})),'
        f'(IF(ISNUMBER(T{row}),MOD(T{row},1),TIMEVALUE(T{row}))-IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row})))*1440,'
        f'(IF(ISNUMBER(T{row}),MOD(T{row},1),TIMEVALUE(T{row}))+1-IF(ISNUMBER(S{row}),MOD(S{row},1),TIMEVALUE(S{row})))*1440)),"")'
    )


def difference_formula(row):
    return f'=IFERROR(IF(OR(W{row}="",X{row}=""),"",W{row}-X{row}),"")'


def accuracy_formula(row):
    return f'=IFERROR(IF(Y{row}="","",IF(Y{row}>15,"Subestimacion",IF(Y{row}<-15,"Sobrestimacion","Asertivo"))),"")'


for path in files:
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Carga Web Cirugias"]

    for row in range(4, 501):
        sheet[f"V{row}"] = delay_formula(row)
        sheet[f"W{row}"] = duration_formula(row)
        # X is intentionally manual: Tiempo programado de cirugía en tabla (Min)
        sheet[f"Y{row}"] = difference_formula(row)
        sheet[f"Z{row}"] = accuracy_formula(row)

        sheet[f"Q{row}"].number_format = "hh:mm"
        sheet[f"S{row}"].number_format = "hh:mm"
        sheet[f"T{row}"].number_format = "hh:mm"
        sheet[f"U{row}"].number_format = "hh:mm"
        for col in ("V", "W", "X", "Y"):
            sheet[f"{col}{row}"].number_format = "0"

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(path)
    print(path)
