import datetime
import numpy as np
import pandas as pd
from pathlib import Path
import random


def create_test_excel_file(filename: str, sheet_name: str):
    filename = Path(filename)
    random.seed(365)
    np.random.seed(365)
    number_of_data_rows = 1000

    # create list of 31 dates
    dates = pd.bdate_range(
        datetime.datetime(2023, 6, 1), freq="1d", periods=31
    ).tolist()
    # Convert datetime to date
    dates = [dt.date() for dt in dates]
    data = {
        "date": [random.choice(dates) for _ in range(number_of_data_rows)],
        "expense": [
            random.choice(["business", "personal"]) for _ in range(number_of_data_rows)
        ],
        "products": [
            random.choice(["book", "ribeye", "coffee", "salmon", "alcohol", "pie"])
            for _ in range(number_of_data_rows)
        ],
        "price": np.random.normal(15, 5, size=(1, number_of_data_rows))[0],
    }

    pd.DataFrame(data).to_excel(
        filename, index=False, sheet_name=sheet_name, float_format="%.2f"
    )
