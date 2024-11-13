from reports.check_structure import check_structure
from reports.callbacks import callbacks_report
from reports.timesheets import timesheets_report
from reports.attendance import attendance_report
from reports.production import production_report
from reports.sales_data import sales_data_report    
from reports.work_completion import work_completion_report
from reports.tech_leads import tech_leads_report
from reports.unconfirmed_work import unconfirmed_work
from reports.work_added import work_added
from reports.ar_report import ar_report
from reports.price import price_report
from reports.rate_per_hour_report import rate_per_hour_report

if __name__ == "__main__":
    check_structure()
    callbacks_report()
    timesheets_report()
    unconfirmed_work()
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
    work_added()
    ar_report()
    price_report(
        salespeople=[
            "Kamaal Sherrod",
            "Hussam Olabi",
            "Rob Dively"
        ]
    )
    rate_per_hour_report(
        custom_joins=[
            ('Jasmine Wilkes', 'Jasmine Wilkins'),
            ('James Wilkes', 'James Wilkins')   
        ]
    )