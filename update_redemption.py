#!/usr/bin/env python3
"""
Drop-in executable for updating 'Redemption Summary' in Monthly Redemption.xlsx.

Reads the current month and fund list from Input File.xlsx, then recomputes
36-month redemption statistics from the monthly 'YYYY MMM' outflow tabs.
"""

from redemption_summary import main, process_redemption_summary

if __name__ == "__main__":
    main()
