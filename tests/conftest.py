import pandas as pd
import pytest


@pytest.fixture()
def sample_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "date": ["2026-01-01", "2026-01-02"],
            "sku": ["SKU1", "SKU2"],
            "product": ["Widget", "Keyboard"],
            "quantity": [2, 1],
            "unit_price": [10, 20],
            "unit_cost": [6, 12],
        }
    )
