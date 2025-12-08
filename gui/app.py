# ───────────── CSV / HEADER HANDLING ───────────── #

    def _post_process_csv(self, csv_path: str):
        """
        After export, read the CSV and:
        - Skip row 1 (the single 'ViolationCTG' cell)
        - Use row 2 as headers
        - Treat row 3+ as data
        - Show header filter dialog and log filtered headers.
        """
        self.log("\nReading CSV to detect headers...")
        try:
            # Skip the first row because it only has "ViolationCTG" in one column
            raw = pd.read_csv(csv_path, header=None, skiprows=1)

            # Now row index 0 is the real header row (original row 2)
            if raw.shape[0] < 1:
                self.log("Not enough rows in CSV to extract headers (need at least 1).")
                return

            header_row = list(raw.iloc[0])
            self.log(f"Detected {len(header_row)} headers from row 2.")
            self.log("Header names:")
            for h in header_row:
                self.log(f"  - {h}")

            # Data rows are index >= 1 (original rows 3+)
            if raw.shape[0] > 1:
                data = raw.iloc[1:].copy()
                data.columns = header_row
                self.log("\nPreview of first few data rows (after header row):")
                preview = data.head(10).to_string(index=False)
                self.log(preview)

            # Open dialog so user can choose columns to filter out
            self.log(
                "\nOpening header filter dialog so you can choose which\n"
                "columns should be filtered out in future versions..."
            )
            HeaderFilterDialog(self, header_row, self.log)

        except Exception as e:
            self.log(f"(Could not read CSV for header inspection: {e})")