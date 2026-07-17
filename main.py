from pathlib import Path

from traffic_optimizer import TrafficOptimizer


BASE_DIR = Path(__file__).parent

TRAFFIC_FILE = BASE_DIR / "Input" / "cerfgsrcsadmin-20260717-7898-vendor-custom-summary-report.csv"

CLIENT_FILE = BASE_DIR / "Input" / "client_list.xlsx"

OUTPUT_FILE = BASE_DIR / "Output" / "Phase3_Output.xlsx"


optimizer = TrafficOptimizer(
    TRAFFIC_FILE,
    CLIENT_FILE
)

optimizer.load_files()

optimizer.clean_data()

optimizer.filter_clients()

optimizer.create_daily_pivot()

optimizer.create_summary()

optimizer.greedy_assignment()

optimizer.export(OUTPUT_FILE)