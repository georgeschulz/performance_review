from reports.check_structure import check_structure
from reports.callbacks import callbacks_report
from reports.timesheets import timesheets_report
from reports.attendance import attendance_report
from reports.production import production_report
from reports.sales_data import sales_data_report    
from reports.work_completion import work_completion_report
from reports.tech_leads import tech_leads_report

if __name__ == "__main__":
    check_structure()
    callbacks_report()
    timesheets_report()
    attendance_report(
        eight_o_clock_starts=[
            "Kamaal Sherrod",
            "Hussam Olabi",
            "Bianca Ramirez"
        ],
    )
    production_report()
    sales_data_report()
    work_completion_report(excluded_techs=[
        "Ivan Chavez"
    ])
    tech_leads_report()