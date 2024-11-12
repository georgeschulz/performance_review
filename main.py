import pandas as pd
from reports.check_structure import check_structure
from reports.callbacks import callbacks_report

if __name__ == "__main__":
    check_structure()
    callbacks_report()