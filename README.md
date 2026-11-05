# Factory Management System Pro v11

A desktop application for small factories and workshops to record daily production, close weeks with PDF/Excel reports, and analyse output by article and by year.

Built with **Python**, **CustomTkinter**, **SQLite**, **ReportLab** and **openpyxl**.

## Features

- **Daily Entry**: add operations by date, choose the article from a catalog (price per meter fills in automatically), edit or delete operations.
- **Articles catalog**: manage articles and their price per meter, show or hide them in the entry list. A new price only applies to new operations.
- **Weekly Summary**: totals per article with a weighted average price. Closing a week creates the PDF and Excel reports first, then marks the operations as PAID in a single transaction.
- **Weeks History**: every closed week is saved. Reprint its PDF, export it to Excel, or compare two weeks.
- **General Archive**: full history with search, date range and status filters. Deleted operations can be restored, and operations from closed weeks are locked.
- **Annual Analysis**: totals for the selected year, monthly bar chart, best month and each article's share of revenue. Export the year to Excel.
- **Safety**: fixed database path next to the program, automatic backups (the last 30 are kept).

## Run

```bash
pip install -r requirements.txt
python main.py
```

On Windows, if `python` opens the Microsoft Store, use `py` instead.

## Version history

- **v11**: articles catalog, editing, soft delete and restore, weeks history, safe week closing, Excel export, per-year analysis, automatic backups.
- **v1.1**: first version (daily entry, inventory summary, records, analytics, PDF invoice).
