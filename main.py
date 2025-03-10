from reports.check_structure import check_structure
from reports.callbacks import callbacks_report
from reports.timesheets import timesheets_report
from reports.attendance import attendance_report
from reports.sales_data import sales_data_report    
from reports.work_completion import work_completion_report
from reports.tech_leads import tech_leads_report
from reports.unconfirmed_work import unconfirmed_work
from reports.work_added import work_added
from reports.ar_report import ar_report
from reports.price import price_report
from reports.rate_per_hour_report import rate_per_hour_report
from reports.retention import retention_report
from reports.calls import calls_report
from reports.close_rate import close_rate
from reports.link_accounts import link_accounts
from reports.first_year_cancels import first_year_cancels
from reports.channel_stats import channel_stats
from reports.job_not_ready import job_not_ready_report

START_DATE = "2023-02-01"
END_DATE = "2025-02-28"

if __name__ == "__main__":
    check_structure()
    link_accounts()
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
        ],
        exclude_techs=[
            "Lawrence Partlow"
        ]
    )
    retention_report(
        start_date=START_DATE,
        end_date=END_DATE
    )
    calls_report(user_mappings=[
        ("101", "Brian Grumbine"),
        ("102", "Hussam Olabi"), 
        ("103", "Kamaal Sherrod"),
        ("104", "Cindy McKnight"),
        ("106", "Bianca Ramirez")
    ])
    close_rate(salespeople=[
        "Kamaal Sherrod",
        "Hussam Olabi",
    ],
    exclude_channels=[
        "Outbound",
        "Sentricon Lead"
    ])
    
    first_year_cancels(salespeople=[
        "Kamaal Sherrod",
        "Hussam Olabi",
    ])
    channel_stats()
    job_not_ready_report()