from __future__ import annotations

import calendar
import html
import re
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from tkinter import messagebox, ttk

import win32com.client


# ============================================================
# ADD YOUR LINKS HERE ONCE
# ============================================================

WORKING_FILE_LINK = (
    "https://ericsson-my.sharepoint.com/:f:/r/personal/chhavi_nagpal_ericsson_com/Documents/Invoice_testing?d=we323e30d11f846baa49a64e947956a94&csf=1&web=1&e=fhrAlg"
)
UPLOAD_PATH_LINK = (
    "https://ericsson-my.sharepoint.com/:f:/r/personal/chhavi_nagpal_ericsson_com/Documents/Invoice_testing?d=we323e30d11f846baa49a64e947956a94&csf=1&web=1&e=7lUT7J"
)


# ============================================================
# COMPANY NAMES
# ============================================================

RAW_COMPANIES = [
    "Alam Matrix Towers Sdn. Bhd.",
    "Alam Mindscape Sdn. Bhd.",
    "Alpha Orange Sdn Bhd",
    "ARRAY TECHNOLOGY SDN BHD",
    "BANGKIT SETIA SDN BHD",
    "Baycom Sdn Bhd",
    "Borneo Restu Sdn Bhd",
    "BORNEO TOWER TECHNOLOGIES SDN. BHD.",
    "BULLISH AIM SDN BHD",
    "D'HARMONI TELCO INFRA SDN BHD",
    "DESS Digital Sdn. Bhd.",
    "DTP SOLUTIONS SDN BHD",
    "EDOTCO MALAYSIA SDN BHD",
    "EXAREX SDN BHD",
    "GLOBAL NET INITIATIVES SDN BHD",
    "GTP NETWORK SDN BHD",
    "HRSB INFRACOMM SDN BHD",
    "INFORIENT INFRASTRUCTURE SDN BHD",
    "INFRA QUEST SDN BHD",
    "JUSTCLICK VISION (M) SDN BHD",
    "KS IT SOLUTIONS SDN BHD",
    "Masjid Al-Khasyi'in",
    "MELAKA ICT HOLDINGS SDN BHD",
    "NETWORK FACILITIES SDN BHD",
    "OCK TELCO INFRA SDN BHD",
    "Omnitel Network Sdn. Bhd.",
    "ORISSA WICOMM (M) SDN BHD",
    "OSI TECHNOLOGY SDN BHD",
    "PDC TELECOMMUNICATION SERVICES SDN BHD",
    "PERLIS COMM SDN BHD",
    "Promajadi Sdn Bhd",
    "Prosontech Sdn Bhd",
    "QEW COMMUNICATIONS SDN BHD",
    "R & R ENGINEERING SUPPLY SDN BHD",
    "REACH TEN COMMUNICATION SDN BHD",
    "REDPYNE SDN BHD",
    "RIYYALCOMM SDN BHD",
    "SKYLINE TECHNOLOGY (M) SDN BHD",
    "SOMNIUM TECHNOLOGIES SDN BHD",
    "Standard Tower Sdn Bhd",
    "SUNWAY DIGITAL WAVE SDN. BHD.",
    "TELE-FLOW CORPORATION SDN BHD",
    "TELEKOM MALAYSIA BERHAD",
    "TERAS MILLENNIUM SDN BHD",
    "THREESIXTY TECHNOLOGIES SDN BHD",
    "TT dotCom Sdn Bhd",
    "Valsa (Sarawak) Sdn Bhd",
    "Ventrix Telecommunications Sdn Bhd",
    "Verticom Sdn Bhd",
    "VOLKSBAHN TECHNOLOGIES",
    "XMT Technologies Sdn Bhd",
    "YIKED BINA SDN BHD",
    "ZETA IOS SDN BHD",
]


def standardize_company_names(names: list[str]) -> list[str]:
    """
    Convert every company name to uppercase,
    remove extra spaces, remove duplicates,
    and sort alphabetically.
    """
    standardized_names = {
        " ".join(name.strip().upper().split())
        for name in names
        if name.strip()
    }

    return sorted(
        standardized_names,
        key=str.casefold,
    )


COMPANIES = standardize_company_names(RAW_COMPANIES)

MONTHS = list(calendar.month_name)[1:]
CURRENT_YEAR = datetime.now().year
YEARS = [str(year) for year in range(2015, CURRENT_YEAR + 1)]
DAYS = [str(day) for day in range(1, 32)]


@dataclass
class EmailDetails:
    company_name: str
    client_email: str
    year: int
    month: str
    day: int


class EmailGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Client Reconciliation Email Generator")
        self.root.geometry("760x500")
        self.root.resizable(False, False)

        today = datetime.now()

        self.company_var = tk.StringVar(value=COMPANIES[0])
        self.client_email_var = tk.StringVar()
        self.year_var = tk.StringVar(value=str(today.year))
        self.month_var = tk.StringVar(value=today.strftime("%B"))
        self.day_var = tk.StringVar(value=str(today.day))

        self.create_form()

    def create_form(self) -> None:
        frame = ttk.Frame(self.root, padding=25)
        frame.pack(fill="both", expand=True)

        title_label = ttk.Label(
            frame,
            text="CLIENT RECONCILIATION EMAIL GENERATOR",
            font=("Segoe UI", 16, "bold"),
        )
        title_label.grid(
            row=0,
            column=0,
            columnspan=2,
            pady=(0, 25),
        )

        ttk.Label(
            frame,
            text="Company Name:",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=10,
        )

        ttk.Combobox(
            frame,
            textvariable=self.company_var,
            values=COMPANIES,
            state="readonly",
            width=55,
        ).grid(
            row=1,
            column=1,
            sticky="w",
            pady=10,
        )

        ttk.Label(
            frame,
            text="Client Email:",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=10,
        )

        ttk.Entry(
            frame,
            textvariable=self.client_email_var,
            width=58,
        ).grid(
            row=2,
            column=1,
            sticky="w",
            pady=10,
        )

        ttk.Label(
            frame,
            text="Year:",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=10,
        )

        ttk.Combobox(
            frame,
            textvariable=self.year_var,
            values=YEARS,
            state="readonly",
            width=18,
        ).grid(
            row=3,
            column=1,
            sticky="w",
            pady=10,
        )

        ttk.Label(
            frame,
            text="Month:",
        ).grid(
            row=4,
            column=0,
            sticky="w",
            pady=10,
        )

        ttk.Combobox(
            frame,
            textvariable=self.month_var,
            values=MONTHS,
            state="readonly",
            width=18,
        ).grid(
            row=4,
            column=1,
            sticky="w",
            pady=10,
        )

        ttk.Label(
            frame,
            text="Date/Day:",
        ).grid(
            row=5,
            column=0,
            sticky="w",
            pady=10,
        )

        ttk.Combobox(
            frame,
            textvariable=self.day_var,
            values=DAYS,
            state="readonly",
            width=18,
        ).grid(
            row=5,
            column=1,
            sticky="w",
            pady=10,
        )

        ttk.Label(
            frame,
            text=(
                "The Working File and UPI upload links are already "
                "configured in the Python code."
            ),
            foreground="gray",
            wraplength=650,
        ).grid(
            row=6,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(25, 15),
        )

        ttk.Button(
            frame,
            text="GENERATE OUTLOOK EMAIL",
            command=self.generate_email,
        ).grid(
            row=7,
            column=0,
            columnspan=2,
            pady=15,
        )

    def generate_email(self) -> None:
        # Force the selected company name to uppercase.
        company_name = self.company_var.get().strip().upper()
        client_email = self.client_email_var.get().strip()
        year_text = self.year_var.get().strip()
        month = self.month_var.get().strip()
        day_text = self.day_var.get().strip()

        if not company_name:
            messagebox.showerror(
                "Missing Information",
                "Please select a company.",
            )
            return

        if not self.is_valid_email(client_email):
            messagebox.showerror(
                "Invalid Email",
                "Please enter a valid client email address.",
            )
            return

        try:
            year = int(year_text)
            day = int(day_text)
            month_number = MONTHS.index(month) + 1
        except ValueError:
            messagebox.showerror(
                "Invalid Date",
                "Please select a valid year, month, and date.",
            )
            return

        last_day_of_month = calendar.monthrange(
            year,
            month_number,
        )[1]

        if day > last_day_of_month:
            messagebox.showerror(
                "Invalid Date",
                f"{month} {year} has only {last_day_of_month} days.",
            )
            return

        if not self.is_valid_link(WORKING_FILE_LINK):
            messagebox.showerror(
                "Working File Link Missing",
                (
                    "Please add your real OneDrive or SharePoint "
                    "working-file link at the top of the code."
                ),
            )
            return

        if not self.is_valid_link(UPLOAD_PATH_LINK):
            messagebox.showerror(
                "UPI Upload Link Missing",
                (
                    "Please add your real UPI or SharePoint upload "
                    "link at the top of the code."
                ),
            )
            return

        details = EmailDetails(
            company_name=company_name,
            client_email=client_email,
            year=year,
            month=month,
            day=day,
        )

        try:
            self.create_outlook_email(details)

            messagebox.showinfo(
                "Success",
                "The Outlook email draft has been created.",
            )

        except Exception as error:
            messagebox.showerror(
                "Outlook Error",
                f"Unable to create the Outlook email.\n\n{error}",
            )

    @staticmethod
    def create_outlook_email(details: EmailDetails) -> None:
        # The company name is uppercase before being inserted into
        # the Outlook subject and body.
        company_name = html.escape(
            details.company_name.upper()
        )

        working_file_link = html.escape(
            WORKING_FILE_LINK,
            quote=True,
        )

        upload_path_link = html.escape(
            UPLOAD_PATH_LINK,
            quote=True,
        )

        period = f"{details.month} {details.year}"

        subject = (
            "RECONCILIATION DATA REQUEST - "
            f"{details.company_name.upper()} - {period}"
        )

        email_body = f"""
<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">

<p>Dear Client,</p>

<p>
As part of the reconciliation process for
<b>{company_name}</b> for <b>{period}</b>, please provide the required
reconciliation information and supporting documents.
</p>

<p><b>WORKING FILE</b></p>

<p>
Please open the working file using the link below and enter the required
information:
</p>

<p>
<a href="{working_file_link}">OPEN WORKING FILE</a>
</p>

<p><b>SUPPORTING DOCUMENTS</b></p>

<p>
Please upload the following supporting documents using the UPI upload path:
</p>

<ul>
    <li>Invoice(s)</li>
    <li>TNB Bill(s)</li>
    <li>Credit Note(s), if applicable</li>
</ul>

<p>
<a href="{upload_path_link}">OPEN UPI UPLOAD PATH</a>
</p>

<p>
Please ensure that all required documents are uploaded for
<b>{period}</b>.
</p>

<p>
Regards,<br>
ERICSSON RECONCILIATION TEAM
</p>

</body>
</html>
"""

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = details.client_email
        mail.Subject = subject
        mail.HTMLBody = email_body

        # Opens the email as a draft for review.
        # It does not send the email automatically.
        mail.Display()

    @staticmethod
    def is_valid_email(email_address: str) -> bool:
        pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
        return re.match(pattern, email_address) is not None

    @staticmethod
    def is_valid_link(link: str) -> bool:
        return link.lower().startswith(
            ("http://", "https://")
        )


def main() -> None:
    root = tk.Tk()
    EmailGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()