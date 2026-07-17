import pandas as pd


class TrafficOptimizer:

    def __init__(self, traffic_file, client_file):

        self.traffic_file = traffic_file
        self.client_file = client_file
        self.assignment_df = None

        self.raw_df = None
        self.client_df = None
        self.filtered_df = None
        self.daily_pivot = None
        self.summary_df = None

    # ==========================================================
    # LOAD FILES
    # ==========================================================

    def load_files(self):

        print("=" * 60)
        print("LOADING FILES")
        print("=" * 60)

        self.raw_df = pd.read_csv(self.traffic_file)

        self.client_df = pd.read_excel(self.client_file)

        self.raw_df.columns = (
            self.raw_df.columns
            .str.strip()
            .str.replace("\ufeff", "", regex=False)
        )

        self.client_df.columns = (
            self.client_df.columns
            .str.strip()
        )

        print(f"Traffic Records : {len(self.raw_df):,}")
        print(f"Clients Loaded  : {len(self.client_df):,}")

    # ==========================================================
    # CLEAN DATA
    # ==========================================================

    def clean_data(self):

        print("\nCleaning Data...")

        self.raw_df["receiveDate"] = pd.to_datetime(
            self.raw_df["receiveDate"],
            errors="coerce"
        )

        self.raw_df["userName"] = (
            self.raw_df["userName"]
            .astype(str)
            .str.strip()
        )

        self.raw_df["totalSubmit"] = pd.to_numeric(
            self.raw_df["totalSubmit"],
            errors="coerce"
        ).fillna(0)

        self.raw_df = self.raw_df.dropna(subset=["receiveDate"])

        print("Cleaning Completed")

    # ==========================================================
    # FILTER CLIENTS
    # ==========================================================

    def filter_clients(self):

        print("\nFiltering Required Clients...")

        self.client_list = (
            self.client_df.iloc[:, 0]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )

        self.filtered_df = self.raw_df[
            self.raw_df["userName"].isin(self.client_list)
        ].copy()

        print(f"Filtered Records : {len(self.filtered_df):,}")
        print(f"Clients with Traffic : {self.filtered_df['userName'].nunique()}")

    # ==========================================================
    # CREATE DAILY PIVOT
    # ==========================================================

    def create_daily_pivot(self):

        print("\nCreating Daily Pivot...")

        self.daily_pivot = (
            self.filtered_df
            .pivot_table(
                index="receiveDate",
                columns="userName",
                values="totalSubmit",
                aggfunc="sum",
                fill_value=0
            )
            .sort_index()
            .reset_index()
        )

        # --------------------------------------------------
        # Add Missing Clients
        # --------------------------------------------------

        for client in self.client_list:

            if client not in self.daily_pivot.columns:
                self.daily_pivot[client] = 0

        # Arrange Columns

        ordered_columns = ["receiveDate"] + self.client_list

        self.daily_pivot = self.daily_pivot.reindex(
            columns=ordered_columns,
            fill_value=0
        )

        print("Daily Pivot Created Successfully")
        print(f"Days : {len(self.daily_pivot)}")
        print(f"Clients : {len(self.client_list)}")

# ==========================================================
# CREATE CLIENT SUMMARY
# ==========================================================

    def create_summary(self):

        print("\nCreating Client Summary...")

        summary = []

        clients = self.daily_pivot.columns[1:]

        total_days = len(self.daily_pivot)

        for client in clients:

            traffic = self.daily_pivot[client]

            total = traffic.sum()
            avg = traffic.mean()
            maximum = traffic.max()
            minimum = traffic.min()

            std = round(traffic.std(), 2)

            active = (traffic > 0).sum()

            if total > 0:
                peak_percent = round((maximum / total) * 100, 2)
            else:
                peak_percent = 0

            if avg > 0:
                cv = round(std / avg, 2)
            else:
                cv = 0

            if active > 0:
                utilization = round((active / total_days) * 100, 2)
            else:
                utilization = 0

            summary.append({

                "Client": client,

                "Total Traffic": total,

                "Average Daily": round(avg, 2),

                "Maximum Daily": maximum,

                "Minimum Daily": minimum,

                "Std Dev": std,

                "Active Days": active,

                "Utilization %": utilization,

                "Peak %": peak_percent,

                "CV": cv

            })

        self.summary_df = (
            pd.DataFrame(summary)
            .sort_values(
                by="Total Traffic",
                ascending=False
            )
            .reset_index(drop=True)
        )

        print("Client Summary Created")

        print(self.summary_df.head(10))

# ==========================================================
# INITIAL GREEDY ASSIGNMENT
# ==========================================================

    def greedy_assignment(self):

        print("\nCreating Initial Line Assignment...")

        line1_total = 0
        line2_total = 0

        # Count zero traffic clients
        line1_zero = 0
        line2_zero = 0

        assignments = []

        # Summary is already sorted by Total Traffic (Descending)
        for _, row in self.summary_df.iterrows():

            client = row["Client"]
            total = row["Total Traffic"]
            avg = row["Average Daily"]
            maximum = row["Maximum Daily"]
            active = row["Active Days"]

            # ----------------------------
            # Zero Traffic Clients
            # ----------------------------
            if total == 0:

                if line1_zero <= line2_zero:
                    line = "Line 1"
                    line1_zero += 1
                else:
                    line = "Line 2"
                    line2_zero += 1

            # ----------------------------
            # Clients with Traffic
            # ----------------------------
            else:

                if line1_total <= line2_total:
                    line = "Line 1"
                    line1_total += total
                else:
                    line = "Line 2"
                    line2_total += total

            # Save Assignment
            assignments.append({

                "Client": client,
                "Assigned Line": line,
                "Total Traffic": total,
                "Average Daily": avg,
                "Maximum Daily": maximum,
                "Active Days": active

            })

        self.assignment_df = pd.DataFrame(assignments)

        print("\nAssignment Completed")
        print(f"Line 1 Total Traffic : {line1_total:,.0f}")
        print(f"Line 2 Total Traffic : {line2_total:,.0f}")
        print(f"Difference           : {abs(line1_total-line2_total):,.0f}")
        print(f"Zero Traffic Clients - Line 1 : {line1_zero}")
        print(f"Zero Traffic Clients - Line 2 : {line2_zero}")

    # ==========================================================
    # EXPORT
    # ==========================================================

    def export(self, output_file):

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

            self.daily_pivot.to_excel(
                writer,
                sheet_name="Daily Pivot",
                index=False
            )

            self.summary_df.to_excel(
                writer,
                sheet_name="Client Summary",
                index=False
            )

            self.assignment_df.to_excel(
                writer,
                sheet_name="Initial Assignment",
                index=False
            )

        print("\nOutput Saved Successfully")
