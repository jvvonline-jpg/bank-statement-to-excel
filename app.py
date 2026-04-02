import streamlit as st
import re
import io
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="AUB Statement to Excel", layout="wide")
st.title("AUB Bank Statement to Excel Converter")
st.write(
    "Upload an Atlantic Union Bank **statement PDF** to convert it into a "
    "structured Excel file with deposits, withdrawals, and running balances."
)

# ── Apple-inspired color palette ──────────────────────────────────────────────
APPLE_BLUE = "0071E3"
APPLE_DEEP_BLUE = "0055CC"
APPLE_BRIGHT_BLUE = "1A8CFF"
APPLE_NEAR_BLACK = "1D1D1F"
APPLE_LIGHT_GRAY = "F5F5F7"
APPLE_MID_GRAY = "D2D2D7"
APPLE_SEC_GRAY = "6E6E73"
WHITE = "FFFFFF"

# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_noise(text):
    """Remove barcode/OCR noise lines from extracted text."""
    MONTHS_PAT = r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
    lines = text.split("\n")
    cleaned = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # If a line has barcode noise BEFORE a transaction, extract just the transaction
        # e.g. "AHPEHKGK... Jan 12 UNITEDHEALTHCARE/BILLING 460.26 92,517.59"
        if not re.match(rf'^{MONTHS_PAT}\s+\d{{1,2}}\s+', stripped):
            embedded = re.search(rf'({MONTHS_PAT}\s+\d{{1,2}}\s+.+)', stripped)
            if embedded and re.search(r'[\d,]+\.\d{2}', embedded.group(1)):
                stripped = embedded.group(1).strip()

        # Skip pure barcode-like strings (long alphanumeric with no spaces)
        if re.match(r'^[A-Z]{10,}$', stripped):
            continue
        if re.match(r'^[A-Z ]{30,}$', stripped) and not any(c.isdigit() for c in stripped):
            if not any(kw in stripped for kw in [
                'TRANSACTION', 'CHECK', 'BALANCE', 'DEPOSIT', 'FUND', 'TRANSFER',
                'CORNERSTONES', 'FAIRFAX', 'PAYCOM', 'FLORES', 'SCHWAB', 'FIDELITY',
                'RELATIONSHIP', 'FLEX BUSINESS', 'ACCOUNT', 'ENDING', 'BEGINNING',
                'TRNSFR', 'MERCHAN', 'FUNDRAISE', 'BIDKIT', 'MONTHLY', 'INTEREST',
                'FEE RECAP', 'SUMMARY', 'SERVICE', 'AVERAGE', 'MINIMUM',
                'COMMONWEALTH', 'SENTARA', 'RAISERIGHT', 'STARKIST', 'AMERICA',
                'DISTRICT', 'UNITEDHEALTHCARE', 'JOHN HANCOCK', 'ROBERT', 'PAYPAL',
                'NETWORK', 'BBGF', 'BTI360', 'NEW YORK', 'WEX INC', 'METKC',
                'UWNCA', 'JK GROUP', 'M&T COMM', 'PPAY'
            ]):
                continue
        # Skip page reference lines
        if re.match(r'^\d{8}\s+\d{7}\s+\d{4}-\d{4}', stripped):
            continue
        if re.match(r'^\d{8}$', stripped):
            continue
        if re.match(r'^\d{7}$', stripped):
            continue
        if re.match(r'^\d{4}-\d{4}$', stripped):
            continue
        if re.match(r'^\d{8}\s+M\d+DDA', stripped):
            continue
        # Clean inline barcode noise (space-separated single chars like "A A L B H N...")
        stripped = re.sub(r'(?:\b[A-Z]\s){5,}[A-Z]\b', '', stripped).strip()
        if not stripped:
            continue
        cleaned.append(stripped)
    return "\n".join(cleaned)


def parse_amount(text):
    """Parse a dollar amount string, returning float. Handles $1,234.56 format."""
    if not text:
        return None
    text = text.strip().replace("$", "").replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def extract_balance_summary(full_text):
    """Extract key balance summary fields from the statement."""
    info = {}

    m = re.search(r'Account Number[:\s]+(\d+)', full_text)
    if m:
        info['account_number'] = m.group(1)

    m = re.search(r'Account Owner\(s\):\s*(.+)', full_text)
    if m:
        info['account_name'] = m.group(1).strip()
    elif re.search(r'CORNERSTONES,?\s*INC', full_text):
        info['account_name'] = 'CORNERSTONES, INC.'

    m = re.search(r'Statement Date\s+(\d{2}/\d{2}/\d{4})', full_text)
    if m:
        info['statement_date'] = m.group(1)

    m = re.search(r'Statement Thru Date\s+(\d{2}/\d{2}/\d{4})', full_text)
    if m:
        info['statement_thru'] = m.group(1)

    m = re.search(r'Account Type.*?Account Number.*?\n(.+?)\s+\d{10}', full_text)
    if m:
        info['account_type'] = m.group(1).strip()
    else:
        m2 = re.search(r'FLEX BUSINESS CKING PLUS', full_text)
        if m2:
            info['account_type'] = 'FLEX BUSINESS CKING PLUS'

    m = re.search(r'Beginning Balance as of\s*(\d{2}/\d{2}/\d{4})\s+\$([\d,]+\.\d{2})', full_text)
    if m:
        info['begin_date'] = m.group(1)
        info['begin_balance'] = parse_amount(m.group(2))

    m = re.search(r'Ending Balance as of\s*(\d{2}/\d{2}/\d{4})\s+\$([\d,]+\.\d{2})', full_text)
    if m:
        info['end_date'] = m.group(1)
        info['end_balance'] = parse_amount(m.group(2))

    m = re.search(r'\+\s*Deposits and Credits\s+\((\d+)\)\s+\$([\d,]+\.\d{2})', full_text)
    if m:
        info['deposit_count'] = int(m.group(1))
        info['deposit_total'] = parse_amount(m.group(2))

    m = re.search(r'-\s*Withdrawals and Debits\s+\((\d+)\)\s+\$([\d,]+\.\d{2})', full_text)
    if m:
        info['withdrawal_count'] = int(m.group(1))
        info['withdrawal_total'] = parse_amount(m.group(2))

    return info


def parse_transactions(full_text):
    """Parse all transactions from the TRANSACTION DETAIL sections."""
    transactions = []

    # Find all TRANSACTION DETAIL sections (skip CHECK TRANSACTION SUMMARY)
    sections = re.split(r'CHECK TRANSACTION SUMMARY', full_text)[0]

    # Match transaction lines: Date Description [Deposit] [Withdrawal] Balance
    # Date format: "Jan DD" or "Feb DD"
    # Amounts: comma-separated with decimals
    lines = sections.split("\n")

    MONTH_MAP = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
    }

    # Pattern for a transaction line starting with a date
    date_pat = re.compile(
        r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+'
    )

    # Pattern for amounts at the end of a line
    # Could be: deposit balance | withdrawal balance | just balance
    amount_pat = re.compile(
        r'([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$'
    )
    single_amount_pat = re.compile(
        r'([\d,]+\.\d{2})\s*$'
    )

    i = 0
    year = None

    # Determine year from statement date
    m = re.search(r'Statement Date\s+\d{2}/\d{2}/(\d{4})', full_text)
    if m:
        year = int(m.group(1))
    else:
        year = 2026

    while i < len(lines):
        line = lines[i].strip()

        # Skip headers, footers, and non-transaction lines
        if not line or line.startswith('Account Number') or line.startswith('Statement'):
            i += 1
            continue
        if line.startswith('Page ') or line.startswith('TRANSACTION DETAIL'):
            i += 1
            continue
        if line.startswith('Date') and 'Description' in line:
            i += 1
            continue
        if 'BEGINNING BALANCE' in line:
            i += 1
            continue
        if 'ENDING BALANCE' in line:
            i += 1
            continue
        if line.startswith('RELATIONSHIP') or line.startswith('Balance Summary'):
            i += 1
            continue
        if line.startswith('Beginning Balance as of') or line.startswith('+ Deposits'):
            i += 1
            continue
        if line.startswith('- Withdrawals') or line.startswith('Ending Balance'):
            i += 1
            continue
        if line.startswith('Service Charges') or line.startswith('Average'):
            i += 1
            continue
        if line.startswith('Minimum Balance') or line.startswith('Interest'):
            i += 1
            continue
        if line.startswith('Annual Percentage') or line.startswith('Number of Days'):
            i += 1
            continue
        if line.startswith('Earnings Summary') or line.startswith('Interest Paid'):
            i += 1
            continue
        if line.startswith('Customer Service') or line.startswith('Check/Items'):
            i += 1
            continue
        if any(line.startswith(x) for x in [
            'PO Box', 'Customer Care', 'Monday', 'Saturday', 'Mailing',
            'Glen Allen', 'Visit Us', 'Atlantic', 'Follow us'
        ]):
            i += 1
            continue

        date_match = date_pat.match(line)
        if not date_match:
            i += 1
            continue

        month_str = date_match.group(1)
        day = int(date_match.group(2))
        month = MONTH_MAP.get(month_str, 1)

        # Determine year for this transaction
        txn_year = year
        # If statement is for January but month is Dec, it's previous year
        if m_stmt := re.search(r'Statement Date\s+(\d{2})', full_text):
            stmt_month = int(m_stmt.group(1))
            if month > stmt_month + 1:
                txn_year = year - 1

        rest = line[date_match.end():]

        # Try to find two amounts at end (deposit/withdrawal + balance)
        two_amt = amount_pat.search(rest)
        one_amt = single_amount_pat.search(rest)

        description = ""
        deposit = None
        withdrawal = None
        balance = None

        if two_amt:
            desc_part = rest[:two_amt.start()].strip()
            amt1 = parse_amount(two_amt.group(1))
            amt2 = parse_amount(two_amt.group(2))

            description = desc_part
            balance = amt2
            # Store amount as deposit initially; fix_deposit_withdrawal_classification
            # will use the running balance chain to correctly classify each transaction
            deposit = amt1

        elif one_amt:
            # Only balance (like BEGINNING BALANCE line) - skip these
            desc_part = rest[:one_amt.start()].strip()
            # This might be a line with only a balance and no amount
            # Or it could be that deposit/withdrawal is the single amount
            # Check if it's really just a balance line
            if 'BALANCE' in desc_part.upper():
                i += 1
                continue
            # Single amount could mean it's a deposit with no separate balance shown
            # This is unusual in this format - skip
            i += 1
            continue
        else:
            i += 1
            continue

        # Collect continuation description lines
        j = i + 1
        while j < len(lines):
            next_line = lines[j].strip()
            if not next_line:
                j += 1
                continue
            # If next line starts with a date, stop
            if date_pat.match(next_line):
                break
            # If next line is a header/footer, stop
            if next_line.startswith('Account Number') or next_line.startswith('Page '):
                break
            if next_line.startswith('TRANSACTION DETAIL'):
                break
            if next_line.startswith('Date') and 'Description' in next_line:
                break
            if next_line.startswith('CHECK TRANSACTION'):
                break
            if re.match(r'^\d{8}\s', next_line):
                j += 1
                continue
            # Check if it's a continuation line (no amounts, just text)
            if not re.search(r'[\d,]+\.\d{2}', next_line):
                # It's a description continuation
                description += " " + next_line
                j += 1
            else:
                break
        i = j

        date_str = f"{month:02d}/{day:02d}/{txn_year}"

        transactions.append({
            'date': date_str,
            'description': description.strip(),
            'deposit': deposit,
            'withdrawal': withdrawal,
            'balance': balance,
        })

    return transactions


def fix_deposit_withdrawal_classification(transactions, begin_balance):
    """
    Use the running balance to correctly classify deposits vs withdrawals.
    If balance increases from previous, the amount is a deposit.
    If balance decreases, it's a withdrawal.
    """
    prev_balance = begin_balance
    for txn in transactions:
        amt = txn['deposit'] if txn['deposit'] is not None else txn['withdrawal']
        if amt is None:
            prev_balance = txn['balance'] if txn['balance'] is not None else prev_balance
            continue

        if txn['balance'] is not None and prev_balance is not None:
            diff = round(txn['balance'] - prev_balance, 2)
            if abs(diff - amt) < 0.02:
                # It's a deposit (balance went up by amt)
                txn['deposit'] = amt
                txn['withdrawal'] = None
            elif abs(diff + amt) < 0.02:
                # It's a withdrawal (balance went down by amt)
                txn['withdrawal'] = amt
                txn['deposit'] = None
            elif diff > 0:
                txn['deposit'] = amt
                txn['withdrawal'] = None
            else:
                txn['withdrawal'] = amt
                txn['deposit'] = None

        prev_balance = txn['balance'] if txn['balance'] else prev_balance

    return transactions


def build_excel(transactions, balance_info):
    """Build Apple-styled Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bank Statement"

    # ── Styles ────────────────────────────────────────────────────────────────
    header_fill = PatternFill('solid', fgColor=APPLE_BLUE)
    header_font = Font(name='Helvetica', bold=True, size=11, color=WHITE)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    title_font = Font(name='Helvetica', bold=True, size=16, color=APPLE_DEEP_BLUE)
    subtitle_font = Font(name='Helvetica', size=11, color=APPLE_SEC_GRAY)
    info_label_font = Font(name='Helvetica', bold=True, size=10, color=APPLE_SEC_GRAY)
    info_value_font = Font(name='Helvetica', bold=True, size=10, color=APPLE_NEAR_BLACK)

    data_font = Font(name='Helvetica', size=10, color=APPLE_NEAR_BLACK)
    data_font_bold = Font(name='Helvetica', bold=True, size=10, color=APPLE_NEAR_BLACK)
    money_fmt = '#,##0.00'
    date_font = Font(name='Helvetica', size=10, color=APPLE_NEAR_BLACK)

    light_gray_fill = PatternFill('solid', fgColor=APPLE_LIGHT_GRAY)
    alt_row_fill = PatternFill('solid', fgColor='F8F9FA')
    totals_fill = PatternFill('solid', fgColor=APPLE_LIGHT_GRAY)

    thin_border = Border(
        bottom=Side(style='thin', color=APPLE_MID_GRAY)
    )
    header_border = Border(
        bottom=Side(style='medium', color=APPLE_DEEP_BLUE)
    )

    # ── Title Section ─────────────────────────────────────────────────────────
    row = 1
    ws.merge_cells('A1:F1')
    ws['A1'].value = "Atlantic Union Bank — Statement Detail"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(vertical='center')

    row = 2
    acct_name = balance_info.get('account_name', '')
    acct_num = balance_info.get('account_number', '')
    acct_type = balance_info.get('account_type', '')
    ws.merge_cells('A2:F2')
    ws['A2'].value = f"{acct_name}  |  Account: ***{acct_num[-4:] if len(acct_num) >= 4 else acct_num}  |  {acct_type}"
    ws['A2'].font = subtitle_font

    # ── Balance Summary ───────────────────────────────────────────────────────
    row = 4
    labels = ['Beginning Balance:', 'Ending Balance:', 'Deposits:', 'Withdrawals:']
    values = [
        f"${balance_info.get('begin_balance', 0):,.2f}  ({balance_info.get('begin_date', '')})",
        f"${balance_info.get('end_balance', 0):,.2f}  ({balance_info.get('end_date', '')})",
        f"{balance_info.get('deposit_count', 0)} items  —  ${balance_info.get('deposit_total', 0):,.2f}",
        f"{balance_info.get('withdrawal_count', 0)} items  —  ${balance_info.get('withdrawal_total', 0):,.2f}",
    ]
    for idx, (lbl, val) in enumerate(zip(labels, values)):
        r = row + idx
        ws.cell(row=r, column=1, value=lbl).font = info_label_font
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(row=r, column=2, value=val).font = info_value_font

    # ── Column Headers ────────────────────────────────────────────────────────
    row = 9
    headers = ['Date', 'Description', 'Deposits', 'Withdrawals', 'Balance', 'Status']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = header_border

    # ── Beginning Balance Row ─────────────────────────────────────────────────
    row = 10
    begin_bal = balance_info.get('begin_balance', 0)
    ws.cell(row=row, column=1, value=balance_info.get('begin_date', '')).font = data_font_bold
    ws.cell(row=row, column=2, value='BEGINNING BALANCE').font = data_font_bold
    ws.cell(row=row, column=5, value=begin_bal).font = data_font_bold
    ws.cell(row=row, column=5).number_format = money_fmt
    for c in range(1, 7):
        ws.cell(row=row, column=c).fill = light_gray_fill
        ws.cell(row=row, column=c).border = thin_border

    # ── Transaction Data ──────────────────────────────────────────────────────
    data_start = 11
    for idx, txn in enumerate(transactions):
        r = data_start + idx
        ws.cell(row=r, column=1, value=txn['date']).font = date_font
        ws.cell(row=r, column=2, value=txn['description']).font = data_font

        if txn['deposit'] is not None:
            ws.cell(row=r, column=3, value=txn['deposit']).font = data_font
            ws.cell(row=r, column=3).number_format = money_fmt

        if txn['withdrawal'] is not None:
            ws.cell(row=r, column=4, value=txn['withdrawal']).font = data_font
            ws.cell(row=r, column=4).number_format = money_fmt

        # Balance: use formula referencing previous balance + deposit - withdrawal
        if idx == 0:
            ws.cell(row=r, column=5).value = f'=E10+C{r}-D{r}'
        else:
            ws.cell(row=r, column=5).value = f'=E{r-1}+C{r}-D{r}'
        ws.cell(row=r, column=5).number_format = money_fmt
        ws.cell(row=r, column=5).font = data_font

        ws.cell(row=r, column=6, value='').font = data_font

        # Alternating row colors
        if idx % 2 == 1:
            for c in range(1, 7):
                ws.cell(row=r, column=c).fill = alt_row_fill

        for c in range(1, 7):
            ws.cell(row=r, column=c).border = thin_border

    last_data = data_start + len(transactions) - 1

    # ── Totals Row ────────────────────────────────────────────────────────────
    totals_row = last_data + 1
    ws.cell(row=totals_row, column=1, value='TOTALS').font = Font(
        name='Helvetica', bold=True, size=11, color=APPLE_DEEP_BLUE
    )
    ws.cell(row=totals_row, column=3).value = f'=SUM(C{data_start}:C{last_data})'
    ws.cell(row=totals_row, column=3).number_format = money_fmt
    ws.cell(row=totals_row, column=3).font = data_font_bold
    ws.cell(row=totals_row, column=4).value = f'=SUM(D{data_start}:D{last_data})'
    ws.cell(row=totals_row, column=4).number_format = money_fmt
    ws.cell(row=totals_row, column=4).font = data_font_bold
    ws.cell(row=totals_row, column=5).value = f'=E{last_data}'
    ws.cell(row=totals_row, column=5).number_format = money_fmt
    ws.cell(row=totals_row, column=5).font = data_font_bold
    for c in range(1, 7):
        ws.cell(row=totals_row, column=c).fill = totals_fill
        ws.cell(row=totals_row, column=c).border = Border(
            top=Side(style='medium', color=APPLE_DEEP_BLUE),
            bottom=Side(style='medium', color=APPLE_DEEP_BLUE)
        )

    # ── Summary Section ───────────────────────────────────────────────────────
    sum_row = totals_row + 2
    ws.cell(row=sum_row, column=1, value='Total items:').font = info_label_font
    ws.cell(row=sum_row, column=3, value=len(transactions)).font = info_value_font
    sum_row += 1
    ws.cell(row=sum_row, column=1, value='Beginning balance:').font = info_label_font
    ws.cell(row=sum_row, column=5, value=begin_bal).font = info_value_font
    ws.cell(row=sum_row, column=5).number_format = money_fmt
    sum_row += 1
    ws.cell(row=sum_row, column=1, value='Ending balance:').font = info_label_font
    ws.cell(row=sum_row, column=5).value = f'=E{last_data}'
    ws.cell(row=sum_row, column=5).font = info_value_font
    ws.cell(row=sum_row, column=5).number_format = money_fmt

    # ── Verification Row ──────────────────────────────────────────────────────
    sum_row += 1
    ws.cell(row=sum_row, column=1, value='Statement ending balance:').font = info_label_font
    ws.cell(row=sum_row, column=5, value=balance_info.get('end_balance', 0)).font = info_value_font
    ws.cell(row=sum_row, column=5).number_format = money_fmt
    sum_row += 1
    ws.cell(row=sum_row, column=1, value='Difference:').font = Font(
        name='Helvetica', bold=True, size=10, color='FF3B30'
    )
    ws.cell(row=sum_row, column=5).value = f'=E{sum_row-1}-E{sum_row-2}'
    ws.cell(row=sum_row, column=5).font = Font(
        name='Helvetica', bold=True, size=10, color='FF3B30'
    )
    ws.cell(row=sum_row, column=5).number_format = money_fmt

    # ── Column Widths ─────────────────────────────────────────────────────────
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 16

    ws.freeze_panes = 'A10'

    return wb


# ── Streamlit UI ──────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("Upload Bank Statement PDF", type="pdf")

if uploaded_file:
    if st.button("Convert to Excel", type="primary"):
        with st.spinner("Extracting text from PDF..."):
            pdf_bytes = uploaded_file.read()
            full_text = ""
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += clean_noise(text) + "\n"

        with st.spinner("Parsing balance summary..."):
            balance_info = extract_balance_summary(full_text)

        with st.spinner("Parsing transactions..."):
            transactions = parse_transactions(full_text)
            begin_bal = balance_info.get('begin_balance', 0)
            transactions = fix_deposit_withdrawal_classification(transactions, begin_bal)

        # ── Display Balance Summary ───────────────────────────────────────
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric(
                "Beginning Balance",
                f"${balance_info.get('begin_balance', 0):,.2f}",
            )
        with col2:
            st.metric(
                "Ending Balance",
                f"${balance_info.get('end_balance', 0):,.2f}",
            )
        with col3:
            st.metric(
                "Deposits",
                f"{balance_info.get('deposit_count', 0)} items",
                f"${balance_info.get('deposit_total', 0):,.2f}",
            )
        with col4:
            st.metric(
                "Withdrawals",
                f"{balance_info.get('withdrawal_count', 0)} items",
                f"${balance_info.get('withdrawal_total', 0):,.2f}",
            )

        # ── Verification ──────────────────────────────────────────────────
        st.markdown("---")
        total_dep = sum(t['deposit'] for t in transactions if t['deposit'])
        total_wth = sum(t['withdrawal'] for t in transactions if t['withdrawal'])
        calc_ending = round(begin_bal + total_dep - total_wth, 2)
        stmt_ending = balance_info.get('end_balance', 0)
        diff = round(calc_ending - stmt_ending, 2)

        vcol1, vcol2, vcol3 = st.columns(3)
        with vcol1:
            st.metric("Calculated Ending Balance", f"${calc_ending:,.2f}")
        with vcol2:
            st.metric("Statement Ending Balance", f"${stmt_ending:,.2f}")
        with vcol3:
            if abs(diff) < 0.02:
                st.metric("Difference", f"${diff:,.2f}", delta="✓ Balanced")
            else:
                st.metric("Difference", f"${diff:,.2f}", delta=f"${diff:,.2f} off", delta_color="inverse")

        st.success(f"Parsed **{len(transactions)}** transactions.")

        # ── Preview ───────────────────────────────────────────────────────
        if transactions:
            preview = []
            for t in transactions[:15]:
                preview.append({
                    'Date': t['date'],
                    'Description': t['description'][:60],
                    'Deposit': f"${t['deposit']:,.2f}" if t['deposit'] else '',
                    'Withdrawal': f"${t['withdrawal']:,.2f}" if t['withdrawal'] else '',
                    'Balance': f"${t['balance']:,.2f}" if t['balance'] else '',
                })
            st.write("**Preview (first 15 transactions):**")
            st.table(preview)

        # ── Build and Download Excel ──────────────────────────────────────
        with st.spinner("Building Excel file..."):
            wb = build_excel(transactions, balance_info)
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

        acct = balance_info.get('account_number', 'statement')[-4:]
        stmt_date = balance_info.get('statement_date', '').replace('/', '-')
        filename = f"AUB_Statement_{acct}_{stmt_date}.xlsx"

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.ml",
            type="primary",
        )
