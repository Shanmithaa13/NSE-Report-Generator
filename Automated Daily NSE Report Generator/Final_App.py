import streamlit as st
import pandas as pd
import os
import re
import time
import logging
import shutil
import zipfile
import smtplib
import hashlib
import json
import datetime
import schedule
import threading
import plotly.graph_objects as go
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# File paths and constants
BASE_DOWNLOAD_DIR = r"C:\Users\Lenovo\Desktop\info inten"
LOG_FILE = os.path.join(BASE_DOWNLOAD_DIR, "reports.log")
OTHERS_FOLDER = os.path.join(BASE_DOWNLOAD_DIR, "Others")
USER_DB_FILE = os.path.join(BASE_DOWNLOAD_DIR, "users.json")
SCHEDULER_DB_FILE = os.path.join(BASE_DOWNLOAD_DIR, "schedulers.json")
ANALYSIS_FOLDER = os.path.join(BASE_DOWNLOAD_DIR, "Analysis")

# File Type Categories
FILE_TYPES = {
    "csv": "CSV",
    "xls": "XLS",
    "xlsx": "XLS",
    "txt": "TXT",
    "doc": "DOC",
    "docx": "DOC",
    "dat": "DAT",
    "bat": "BAT"
}

# Ensure all directories exist
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)
os.makedirs(OTHERS_FOLDER, exist_ok=True)
os.makedirs(ANALYSIS_FOLDER, exist_ok=True)

# Logging Setup
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(message)s")

# Initialize user database if it doesn't exist
if not os.path.exists(USER_DB_FILE):
    with open(USER_DB_FILE, 'w') as f:
        json.dump({}, f)

# Initialize scheduler database if it doesn't exist
if not os.path.exists(SCHEDULER_DB_FILE):
    with open(SCHEDULER_DB_FILE, 'w') as f:
        json.dump([], f)

# =====================================================================
# User Authentication Functions
# =====================================================================

def hash_password(password):
    """Hash a password for storing."""
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    """Load user database."""
    try:
        with open(USER_DB_FILE, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_users(users):
    """Save user database."""
    with open(USER_DB_FILE, 'w') as f:
        json.dump(users, f)

def register_user(username, email, password):
    """Register a new user."""
    users = load_users()
    
    # Check if username already exists
    if username in users:
        return False, "Username already exists"
    
    # Check if email already exists
    for user in users.values():
        if user['email'] == email:
            return False, "Email already exists"
    
    # Add new user
    users[username] = {
        'email': email,
        'password': hash_password(password),
    }
    save_users(users)
    return True, "Registration successful"

def verify_user(username, password):
    """Verify user credentials."""
    users = load_users()
    
    if username not in users:
        return False, "Username not found"
    
    if users[username]['password'] != hash_password(password):
        return False, "Incorrect password"
    
    return True, "Login successful"

def reset_password(username, email, new_password):
    """Reset user password."""
    users = load_users()
    
    if username not in users:
        return False, "Username not found"
    
    if users[username]['email'] != email:
        return False, "Email does not match"
    
    users[username]['password'] = hash_password(new_password)
    save_users(users)
    return True, "Password reset successful"

def get_user_email(username):
    """Get user email."""
    users = load_users()
    if username in users:
        return users[username]['email']
    return None

# =====================================================================
# Scheduler Functions
# =====================================================================

def load_schedulers():
    """Load scheduler database."""
    try:
        with open(SCHEDULER_DB_FILE, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []
def save_schedulers(schedulers):
    """Save scheduler database."""
    with open(SCHEDULER_DB_FILE, 'w') as f:
        json.dump(schedulers, f)

def add_scheduler(username, scheduler_name, execution_time, frequency):
    """Add a new scheduler."""
    schedulers = load_schedulers()
    
    # Create new scheduler
    new_scheduler = {
        'username': username,
        'name': scheduler_name,
        'execution_time': execution_time,
        'frequency': frequency,
        'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    
    schedulers.append(new_scheduler)
    save_schedulers(schedulers)
    return True, "Scheduler created successfully"

def get_user_schedulers(username):
    """Get schedulers for a user."""
    schedulers = load_schedulers()
    return [s for s in schedulers if s['username'] == username]

# =====================================================================
# Email Functions
# =====================================================================

def send_email(to_email, subject, body):
    """Send email to the user."""
    try:
        from_email = os.environ.get("EMAIL_ADDRESS", "shanmithaasrinivasan@gmail.com")
        password = os.environ.get("EMAIL_PASSWORD", "keqb quan qutc dnkd")
        smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Attach body
        msg.attach(MIMEText(body, 'plain'))
        
        # Connect to server and send
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()
        
        logging.info(f"Email sent to {to_email}")
        return True, "Email sent successfully"
    except Exception as e:
        logging.error(f"Failed to send email: {str(e)}")
        return False, f"Failed to send email: {str(e)}"

# =====================================================================
# NSE Report Download Functions
# =====================================================================

def log_message(message):
    """Logs a message to the log file and prints it."""
    logging.info(message)
    print(message)

def ensure_folder_structure(folder_path):
    """Ensure Folder Structure Exists"""
    os.makedirs(folder_path, exist_ok=True)

def handle_exception(error_message, exception_obj):
    """Handle Exceptions"""
    log_message(f"ERROR: {error_message} - {str(exception_obj)}")

def extract_date_from_filename(filename):
    """Extract Date from Filename"""
    match_ddmmyy = re.search(r'(\d{2})(\d{2})(\d{2})', filename)
    match_ddmmyyyy = re.search(r'(\d{2})(\d{2})(2025)', filename)
    match_yyyymmdd = re.search(r'(2025)(\d{2})(\d{2})', filename)
    match_mmddyyyy = re.search(r'(\d{2})(\d{2})(2025)', filename)

    if match_ddmmyyyy:
        day, month, year = match_ddmmyyyy.groups()
    elif match_yyyymmdd:
        year, month, day = match_yyyymmdd.groups()
    elif match_mmddyyyy:
        month, day, year = match_mmddyyyy.groups()
    elif match_ddmmyy:
        day, month, year = match_ddmmyy.groups()
        year = "2025"
    else:
        return None

    return f"{day}-{month}-{year}"

def get_file_type_folder(file_name):
    """Get File Type Folder"""
    file_extension = os.path.splitext(file_name)[1].lower().strip(".")
    return FILE_TYPES.get(file_extension, None)

def remove_duplicate_files(destination_folder, file_name):
    """Remove Duplicate Files"""
    file_base, file_ext = os.path.splitext(file_name)

    for existing_file in os.listdir(destination_folder):
        existing_base, existing_ext = os.path.splitext(existing_file)

        if existing_base.startswith(file_base) and existing_ext == file_ext:
            existing_path = os.path.join(destination_folder, existing_file)

            try:
                os.remove(existing_path)
                log_message(f"Removed duplicate file: {existing_path}")
            except Exception as e:
                handle_exception(f"Failed to remove duplicate file {existing_file}", e)

def move_file_to_folder(file_path):
    """Move File to Correct Folder"""
    try:
        file_name = os.path.basename(file_path)
        file_date = extract_date_from_filename(file_name)
        file_type = get_file_type_folder(file_name)

        if file_date == None:
            log_message(f"Warning: Unable to determine date for {file_name}, moving to Others folder.")
            file_date = "Others"

        destination_folder = os.path.join(BASE_DOWNLOAD_DIR, file_date, file_type) if file_type else OTHERS_FOLDER
        ensure_folder_structure(destination_folder)

        remove_duplicate_files(destination_folder, file_name)
        shutil.move(file_path, os.path.join(destination_folder, file_name))

        log_message(f"Moved: {file_name} -> {destination_folder}")

    except Exception as e:
        handle_exception(f"Failed to move file {file_path}", e)

def extract_zip(zip_path):
    """Extract ZIP & Move Contents"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            extracted_files = zip_ref.namelist()
            zip_ref.extractall(BASE_DOWNLOAD_DIR)

        for file_name in extracted_files:
            extracted_file_path = os.path.join(BASE_DOWNLOAD_DIR, file_name)

            if os.path.exists(extracted_file_path):
                move_file_to_folder(extracted_file_path)

        os.remove(zip_path)
        log_message(f"Deleted original ZIP: {zip_path}")

    except Exception as e:
        handle_exception(f"Error extracting {zip_path}", e)

def download_nse_reports(driver):
    """Download Reports from NSE"""
    url = "https://www.nseindia.com/all-reports"
    driver.get(url)
    wait = WebDriverWait(driver, 30)
    time.sleep(5)

    try:
        report_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".reportsDownload a")))
        log_message(f"Found {len(report_elements)} report(s).")

        for i in range(len(report_elements)):
            try:
                report_elements = driver.find_elements(By.CSS_SELECTOR, ".reportsDownload a")
                data_link = report_elements[i].get_attribute("data-link") or report_elements[i].get_attribute("href")

                if data_link and data_link != "javascript:;":
                    log_message(f"Downloading: {data_link}")
                    driver.execute_script("window.open(arguments[0], '_blank');", data_link)
                    time.sleep(5)
                    driver.switch_to.window(driver.window_handles[0])
                else:
                    log_message(f"Downloading Report {i+1}")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", report_elements[i])
                    time.sleep(1)
                    ActionChains(driver).move_to_element(report_elements[i]).click().perform()
                    time.sleep(5)

            except Exception as e:
                handle_exception(f"Error downloading report {i+1}", e)

    except Exception as e:
        handle_exception("Failed to fetch reports from NSE", e)

def setup_and_download():
    """Initialize WebDriver & Download Reports"""
    try:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        chrome_prefs = {
            "download.default_directory": BASE_DOWNLOAD_DIR,
            "download.prompt_for_download": False
        }
        chrome_options.add_experimental_option("prefs", chrome_prefs)

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        download_nse_reports(driver)
        driver.quit()

    except Exception as e:
        handle_exception("Failed to initialize WebDriver", e)

def process_downloaded_files():
    """Process Downloaded Files"""
    downloaded_files = os.listdir(BASE_DOWNLOAD_DIR)

    for file_name in downloaded_files:
        file_path = os.path.join(BASE_DOWNLOAD_DIR, file_name)

        if os.path.isfile(file_path) and file_name not in ["users.json", "schedulers.json", "reports.log"]:
            if file_name.endswith(".zip"):
                extract_zip(file_path)
            else:
                move_file_to_folder(file_path)

def cleanup_others_folder():
    """Cleanup Others Folder"""
    reports_log_path = os.path.join(OTHERS_FOLDER, "reports.log")
    txt_folder_path = os.path.join(OTHERS_FOLDER, "TXT")

    if os.path.exists(reports_log_path):
        os.remove(reports_log_path)
        log_message("Deleted reports.log from Others folder.")

    if os.path.exists(txt_folder_path):
        shutil.rmtree(txt_folder_path)
        log_message("Deleted TXT folder from Others folder.")

def run_download_process(username=None):
    """Run the complete download process"""
    log_message("Starting NSE Report Download & Processing...")
    setup_and_download()
    process_downloaded_files()
    cleanup_others_folder()
    log_message("Process Completed Successfully!")
    
    # Send email notification if username provided
    if username:
        email = get_user_email(username)
        if email:
            today = datetime.now().strftime("%Y-%m-%d")
            subject = f"NSE Reports Downloaded Successfully - {today}"
            body = f"""
            Dear {username},
            
            We're pleased to inform you that your scheduled NSE reports download has completed successfully.
            
            Date: {today}
            Time: {datetime.now().strftime("%H:%M:%S")}
            
            You can view the downloaded reports in your dashboard.
            
            Thank you for using NSEBot.
            """
            send_email(email, subject, body)

# =====================================================================
# CSV Analysis Functions
# =====================================================================

def create_analysis_log(log_file_path, message):
    """
    Create or append to the Data Analysis Log file.
    """
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

def analyze_csv(file_path, recursion=False):
    if not recursion:
        st.subheader("Stock Analysis")

        # Add a "Start Analysis" button
        if st.button("Start Analysis", key="start_analysis_button"):
            # Run analysis and store results in session state
            st.session_state["analysis_results"] = analyze_files_separately(BASE_DOWNLOAD_DIR)
            st.success("Analysis completed successfully!")

        # Check if analysis results are available in session state
        if "analysis_results" not in st.session_state:
            st.info("Click 'Start Analysis' to begin the analysis process.")
            return

        # Retrieve results from session state
        buy_prices, sell_prices, top_amt, top_qty, top_short_qty, columns = st.session_state["analysis_results"]

        # Add a unique key to the selectbox
        category = st.selectbox(
            "Select Analysis Category",
            ["Top Buy Stocks", "Top Sold Stocks", "Highest Quantity", "Highest Amount"],
            key="analysis_category_selectbox"
        )

        # Display results based on the selected category
        if category == "Top Buy Stocks":
            if not buy_prices.empty:
                st.title("Top Buy Stocks")
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=buy_prices, x=buy_prices.columns[1], y=buy_prices.columns[0], ax=ax, palette="Greens_d")
                plt.title("Top 5 Buy Stocks")
                st.pyplot(fig)
                st.write(buy_prices)
            else:
                st.info("No data available for Top 5 Buy Stocks")

        elif category == "Top Sold Stocks":
            if not sell_prices.empty:
                st.title("Top Sold Stocks")
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=sell_prices, x=sell_prices.columns[1], y=sell_prices.columns[0], ax=ax, palette="Reds_d")
                plt.title("Top 5 Sold Stocks")
                st.pyplot(fig)
                st.write(sell_prices)
            else:
                st.info("No data available for Top 5 Sold Stocks")

        elif category == "Highest Quantity":
            if not top_qty.empty:
                st.title("Highest Quantity")
                aggregated_qty = top_qty.groupby(top_qty.columns[0])[top_qty.columns[1]].sum().reset_index()
                aggregated_qty.columns = ["Stock Name", "Total Quantity"]
                aggregated_qty = aggregated_qty.sort_values(by="Total Quantity", ascending=False).head(5)
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=aggregated_qty, x="Total Quantity", y="Stock Name", ax=ax, palette="Purples_d")
                plt.title("Top 5 Stocks by Quantity")
                st.pyplot(fig)
                st.write(aggregated_qty)
            else:
                st.info("No data available for Highest Quantity")
        
        elif category == "Highest Amount":
            if not top_amt.empty:
                st.title("Highest Amount")
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(data=top_amt, x=top_amt.columns[1], y=top_amt.columns[0], ax=ax, palette="Blues_d")
                plt.title("Top 5 Amounts")
                st.pyplot(fig)
                st.write(top_amt)
            else:
                st.info("No data available for 5 Highest Amount")

    # Recursive call logic
    else:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None

        try:
            # Read the CSV file, skipping bad lines
            df = pd.read_csv(file_path, on_bad_lines='skip')
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None

        # Drop duplicate columns
        df = df.loc[:, ~df.columns.duplicated()]

        # Identify the required columns based on keywords
        name_col = None
        buy_sell_col = None
        price_col = None
        qty_col = None
        amt_col = None

        for col in df.columns:
            if 'client' in col.lower() or ('name' in col.lower() or 'symbol' in col.lower() or 'security' in col.lower()):
                name_col = col
            elif 'buy/sell' in col.lower() or 'transaction' in col.lower():
                buy_sell_col = col
            elif 'price' in col.lower() or 'trade price' in col.lower() or 'wght. avg. price' in col.lower():
                price_col = col
            elif 'qty' in col.lower() or 'quantity' in col.lower() or 'quantity traded' in col.lower():
                qty_col = col
            elif 'amt' in col.lower() or 'amount' in col.lower() or 'amt fin by all the members' in col.lower():
                amt_col = col

        # Check for column names inside the rows
        if not (name_col and (buy_sell_col and price_col) or (qty_col and amt_col and name_col) or ((name_col) and (qty_col or 'qty' in df.columns) and (price_col or amt_col))):
            for i, row in df.iterrows():
                if 'name' in row.astype(str).str.lower().values or 'qty' in row.astype(str).str.lower().values or 'amt' in row.astype(str).str.lower().values or 'security name' in row.astype(str).str.lower().values:
                    df.columns = row
                    df = df.drop(i)
                    break

            for col in df.columns:
                if 'client' in col.lower() or ('name' in col.lower() or 'symbol' in col.lower() or 'security' in col.lower()):
                    name_col = col
                elif 'buy/sell' in col.lower() or 'transaction' in col.lower():
                    buy_sell_col = col
                elif 'price' in col.lower() or 'trade price' in col.lower() or 'wght. avg. price' in col.lower():
                    price_col = col
                elif 'qty' in col.lower() or 'quantity' in col.lower() or 'quantity traded' in col.lower():
                    qty_col = col
                elif 'amt' in col.lower() or 'amount' in col.lower() or 'amt fin by all the members' in col.lower():
                    amt_col = col

        if not (name_col and (buy_sell_col and price_col) or (qty_col and amt_col and name_col) or ((name_col) and (qty_col or 'qty' in df.columns) and (price_col or amt_col))):
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), df.columns.tolist()

        # Convert 'Buy/Sell' column to string and normalize if it exists
        if buy_sell_col:
            df[buy_sell_col] = df[buy_sell_col].astype(str).str.strip().str.upper()

        # Filter buy and sell prices if columns exist
        buy_prices = pd.DataFrame()
        sell_prices = pd.DataFrame()
        if buy_sell_col and price_col:
            buy_prices = df[df[buy_sell_col] == 'BUY'][[name_col, price_col, qty_col]].sort_values(by=price_col, ascending=False).head(5)
            sell_prices = df[df[buy_sell_col] == 'SELL'][[name_col, price_col, qty_col]].sort_values(by=price_col, ascending=True).head(5)

        # Find top 5 amounts and quantities if columns exist
        top_amt = pd.DataFrame()
        top_qty = pd.DataFrame()
        if amt_col and name_col:
            top_amt = df[[name_col, amt_col]].sort_values(by=amt_col, ascending=False).head(5)
        if qty_col and name_col:
            top_qty = df[[name_col, qty_col]].sort_values(by=qty_col, ascending=False).head(5)

        # Find top 5 quantities and corresponding security names for shortselling
        top_short_qty = pd.DataFrame()
        if name_col and qty_col:
            top_short_qty = df[[name_col, qty_col]].sort_values(by=qty_col, ascending=False).head(5)

        return buy_prices, sell_prices, top_amt, top_qty, top_short_qty, None

def analyze_files_separately(directory):
    column_info = []
    all_buy_prices = pd.DataFrame()
    all_sell_prices = pd.DataFrame()
    all_top_amt = pd.DataFrame()
    all_top_qty = pd.DataFrame()
    all_top_short_qty = pd.DataFrame()

    # Create the Data Analysis Log file
    log_file_path = os.path.join(ANALYSIS_FOLDER, "data_analysis_log.txt")
    create_analysis_log(log_file_path, "Starting Data Analysis Process")

    for filename in os.listdir(directory):
        if filename.endswith(".csv"):
            file_path = os.path.join(directory, filename)
            create_analysis_log(log_file_path, f"Processing file: {filename}")

            try:
                df = pd.read_csv(file_path, on_bad_lines='skip')
                buy_prices, sell_prices, top_amt, top_qty, top_short_qty, columns = analyze_csv(file_path, recursion=True)

                if not buy_prices.empty or not sell_prices.empty or not top_amt.empty or not top_qty.empty or not top_short_qty.empty:
                    all_buy_prices = pd.concat([all_buy_prices, buy_prices])
                    all_sell_prices = pd.concat([all_sell_prices, sell_prices])
                    all_top_amt = pd.concat([all_top_amt, top_amt])
                    all_top_qty = pd.concat([all_top_qty, top_qty])
                    all_top_short_qty = pd.concat([all_top_short_qty, top_short_qty])
                    create_analysis_log(log_file_path, f"✅ Successfully processed file: {filename}")
                    print(f"✅ Successfully processed file: {filename}")
                elif columns is not None:
                    column_info.append((filename, ', '.join(columns)))
                    create_analysis_log(log_file_path, f"✅File {filename} does not contain the required columns.")
                    print(f"❌ File {filename} does not contain the required columns. Columns found: {columns}")

            except pd.errors.ParserError as e:
                create_analysis_log(log_file_path, f"Error reading {filename}: {e}")
                print(f"Error reading {filename}: {e}")
                continue

    # Save column information to a text file in table format
    with open(os.path.join(ANALYSIS_FOLDER, 'column_info.txt'), 'w') as f:
        f.write(f"{'File Name':<30} | Columns\n")
        f.write(f"{'-'*30} | {'-'*50}\n")
        for info in column_info:
            f.write(f"{info[0]:<30} | {info[1]}\n")

    # Reset index for all DataFrames to avoid duplicate labels
    all_buy_prices = all_buy_prices.reset_index(drop=True)
    all_sell_prices = all_sell_prices.reset_index(drop=True)
    all_top_amt = all_top_amt.reset_index(drop=True)
    all_top_qty = all_top_qty.reset_index(drop=True)
    all_top_short_qty = all_top_short_qty.reset_index(drop=True)

    create_analysis_log(log_file_path, "Data Analysis Process Completed")
    return all_buy_prices, all_sell_prices, all_top_amt, all_top_qty, all_top_short_qty, column_info

# =====================================================================
# Dashboard Functions
# =====================================================================

def count_file_types():
    """Count the number of files of each type in the download directory."""
    file_counts = {ext: 0 for ext in FILE_TYPES.keys()}
    
    # Walk through all directories and count files
    for root, _, files in os.walk(BASE_DOWNLOAD_DIR):
        for file in files:
            if file in ["users.json", "schedulers.json", "reports.log"]:
                continue
                
            ext = os.path.splitext(file)[1].lower().strip(".")
            if ext in file_counts:
                file_counts[ext] += 1

    return file_counts

def display_dashboard_content():
    """Display the dashboard content with pie chart and total reports."""
    # File type distribution data
    file_counts = count_file_types()
    total_files = sum(file_counts.values())
    
    # Remove empty categories and calculate percentages
    filtered_counts = {k: v for k, v in file_counts.items() if v > 0}
    file_types = list(filtered_counts.keys())
    counts = list(filtered_counts.values())
    
    if total_files > 0:
        percentages = [count / total_files * 100 for count in counts]
        
        # Create columns for pie chart and total reports
        left_col, right_col = st.columns([1, 1])
        
        # Left column with pie chart
        with left_col:
            st.markdown("<h3 style='color: white;'>File Type Distribution</h3>", unsafe_allow_html=True)
            
            # Create pie chart using plotly
            fig = px.pie(
                values=percentages,
                names=file_types,
                color_discrete_sequence=px.colors.sequential.Blues_r,
                hole=0.3,
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(
                showlegend=False,
                margin=dict(t=0, b=0, l=0, r=0),
                height=400,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='white', size=14)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display total reports
            st.markdown("<h3 style='color: white;'>Total Reports</h3>", unsafe_allow_html=True)
            st.markdown(f"<h1 style='color: white; font-size: 60px;'>{total_files}</h1>", unsafe_allow_html=True)
        
        # Right column with custom download schedule
        with right_col:
            st.markdown("<h2 style='color: white;'>Custom Download Schedule</h2>", unsafe_allow_html=True)
            
            # Create color legend for file types
            legend_col1, legend_col2, legend_col3 = st.columns([1, 1, 1])
            
            # Define colors for legend
            legend_colors = {
                "csv": "#75E6DA",
                "dat": "#189AB4",
                "txt": "#FF5E5B",
                "doc": "#FF9B54",
                "xls": "#8AC926",
                "xlsx": "#05445E"
            }
            
            # Create legend in columns
            legends_per_column = 2
            all_types = list(legend_colors.keys())
            
            for i, col in enumerate([legend_col1, legend_col2, legend_col3]):
                with col:
                    start_idx = i * legends_per_column
                    end_idx = min(start_idx + legends_per_column, len(all_types))
                    for file_type in all_types[start_idx:end_idx]:
                        color = legend_colors[file_type]
                        st.markdown(f"""
                        <div style='display: flex; align-items: center;'>
                            <div style='background-color: {color}; width: 15px; height: 15px; margin-right: 5px;'></div>
                            <span style='color: white;'>{file_type}</span>
                        </div>
                        """, unsafe_allow_html=True)
            
            # Date and time selectors
            st.markdown("<p style='color: white;'>Select the date</p>", unsafe_allow_html=True)
            selected_date = st.date_input("", datetime.now().date(), key="date_select", label_visibility="collapsed")
            
            st.markdown("<p style='color: white;'>Select the time (HH:MM)</p>", unsafe_allow_html=True)
            selected_time = st.text_input("", "16:16", key="time_select", label_visibility="collapsed")
            
            # Display scheduled time
            formatted_date = selected_date.strftime("%Y-%m-%d")
            st.markdown(f"<p style='color: white;'>Scheduled time: {formatted_date} {selected_time}</p>", unsafe_allow_html=True)
            
            # Action buttons
            save_col, download_col = st.columns([1, 1])
            
            with save_col:
                if st.button("Save Schedule"):
                    if st.session_state.get('logged_in', False):
                        username = st.session_state.get('username', "")
                        scheduler_name = f"Schedule {formatted_date} {selected_time}"
                        add_scheduler(username, scheduler_name, selected_time, "Daily")
                        st.success("Schedule saved successfully!")
                    else:
                        st.warning("Please log in to save schedules.")
            
            with download_col:
                if st.button("Start Downloading Reports Now"):
                    with st.spinner("Downloading reports... Please wait."):
                        if st.session_state.get('logged_in', False):
                            run_download_process(st.session_state.get('username', ""))
                        else:
                            run_download_process()
                        st.success("Download completed successfully!")
    else:
        st.info("No files have been downloaded yet. Use the 'Start Downloading Reports Now' button to download reports.")

def logout():
    """Logout the user."""
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.page = "login"
    st.success("You have been logged out.")
    st.rerun()

# =====================================================================
# Main Application Logic
# =====================================================================

def display_login_page():
    """Display the login page."""
    st.header("Login")
    
    username = st.text_input("Username", key="login_username")
    password = st.text_input("Password", type="password", key="login_password")
    
    if st.button("Login", key="login_button"):
        if username and password:
            success, message = verify_user(username, password)
            
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.page = "dashboard"
                st.success(message)
                st.rerun()
            else:
                st.error(message)
        else:
            st.error("Please enter both username and password")

def display_register_page():
    """Display the registration page."""
    st.header("Register")
    
    username = st.text_input("Username", key="register_username")
    email = st.text_input("Email", key="register_email")
    password = st.text_input("Password", type="password", key="register_password")
    confirm_password = st.text_input("Confirm Password", type="password", key="register_confirm_password")
    
    if st.button("Register", key="register_button"):
        if username and email and password and confirm_password:
            if password != confirm_password:
                st.error("Passwords do not match")
            else:
                success, message = register_user(username, email, password)
                
                if success:
                    st.success(message)
                    st.session_state.page = "login"
                    st.rerun()
                else:
                    st.error(message)
        else:
            st.error("Please fill in all fields")

def display_forgot_password_page():
    """Display the forgot password page."""
    st.header("Reset Password")
    
    username = st.text_input("Username", key="reset_username")
    email = st.text_input("Email", key="reset_email")
    new_password = st.text_input("New Password", type="password", key="reset_new_password")
    confirm_password = st.text_input("Confirm New Password", type="password", key="reset_confirm_password")
    
    if st.button("Reset Password", key="reset_button"):
        if username and email and new_password and confirm_password:
            if new_password != confirm_password:
                st.error("Passwords do not match")
            else:
                success, message = reset_password(username, email, new_password)
                
                if success:
                    st.success(message)
                    st.session_state.page = "login"
                    st.rerun()
                else:
                    st.error(message)
        else:
            st.error("Please fill in all fields")

def display_scheduler():
    """Display the scheduler page."""
    st.header("Report Download Scheduler")
    
    # Show existing schedulers
    user_schedulers = get_user_schedulers(st.session_state.username)
    
    if user_schedulers:
        st.subheader("Your Scheduled Downloads")
        for idx, scheduler in enumerate(user_schedulers):
            st.markdown(f"""
            {scheduler['name']}  
            Execution Time: {scheduler['execution_time']}  
            Frequency: {scheduler['frequency']}  
            Created: {scheduler['created_at']}
            """)
            st.divider()
    
    # Create new scheduler
    st.subheader("Create New Scheduler")
    
    scheduler_name = st.text_input("Scheduler Name", key="scheduler_name")
    execution_time = st.time_input("Execution Time", datetime.now().time(), key="execution_time")
    frequency = st.selectbox("Frequency", ["Daily", "Weekdays", "Weekly", "Monthly"], key="frequency")
    
    if st.button("Create Scheduler", key="create_scheduler_button"):
        if scheduler_name and execution_time and frequency:
            success, message = add_scheduler(
                st.session_state.username,
                scheduler_name,
                execution_time.strftime("%H:%M"),
                frequency
            )
            
            if success:
                st.success(message)
                
                # Schedule the task
                exec_time = execution_time.strftime("%H:%M")
                
                if frequency == "Daily":
                    schedule.every().day.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                elif frequency == "Weekdays":
                    schedule.every().monday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                    schedule.every().tuesday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                    schedule.every().wednesday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                    schedule.every().thursday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                    schedule.every().friday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                elif frequency == "Weekly":
                    schedule.every().monday.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                elif frequency == "Monthly":
                    # This is a simple approximation, as schedule doesn't directly support monthly
                    schedule.every(30).days.at(exec_time).do(
                        run_download_process, username=st.session_state.username
                    )
                
                st.rerun()
            else:
                st.error(message)
        else:
            st.error("Please fill in all fields")
    
    # Manual download button
    st.subheader("Manual Download")
    if st.button("Download Reports Now", key="download_now_button"):
        with st.spinner("Downloading reports from NSE... This may take a few minutes."):
            run_download_process(st.session_state.username)
        st.success("Reports downloaded successfully!")

def display_dashboard():
    """Display the complete dashboard with analytics and downloaded files."""
    
    # Create tab structure for dashboard
    tabs = st.tabs(["Overview", "Statistics", "Log History","Analysis"])
    
    with tabs[0]:
        display_dashboard_content()
        
        # Display downloaded reports section
        st.markdown("<h2 style='text-align: center; color: white; margin-top: 4rem;'>Downloaded Reports</h2>", unsafe_allow_html=True)
        
        # Get recent downloads from log
        recent_downloads = []
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r') as log_file:
                for line in log_file:
                    if "Moved:" in line:
                        parts = line.split("Moved:")
                        if len(parts) > 1:
                            file_info = parts[1].strip()
                            recent_downloads.append(file_info)
        
        # Display recent downloads in a clean table
        if recent_downloads:
            # Create a dataframe from recent downloads
            download_data = []
            for download in recent_downloads[-10:]:  # Show last 10 downloads
                parts = download.split("->")
                if len(parts) == 2:
                    file_name = parts[0].strip()
                    destination = parts[1].strip()
                    ext = os.path.splitext(file_name)[1].lower().strip(".")
                    if ext in ["csv", "dat", "txt", "doc", "xls"]:  # Only show specific file types
                        download_data.append({"File Name": file_name, "Destination": destination})
            
            if download_data:
                download_df = pd.DataFrame(download_data)
                st.dataframe(download_df, use_container_width=True)
        else:
            st.info("No download history available yet.")
    
    with tabs[1]:
        st.subheader("File Statistics")
        
        # Get file type counts
        file_counts = count_file_types()
        file_counts = {k: v for k, v in file_counts.items() if k in ["csv", "dat", "txt", "doc", "xls"] and v > 0}
        
        if file_counts:
            # Create bar chart of file types
            fig = px.bar(
                x=list(file_counts.keys()),
                y=list(file_counts.values()),
                labels={'x': 'File Type', 'y': 'Count'},
                color=list(file_counts.values()),
                color_continuous_scale=px.colors.sequential.Blues,
            )
            fig.update_layout(
                title="File Type Distribution",
                xaxis_title="File Type",
                yaxis_title="Count",
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0.1)',
                font=dict(color='white')
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Create dataframe for file counts
            file_df = pd.DataFrame({
                "File Type": list(file_counts.keys()),
                "Count": list(file_counts.values()),
                "Percentage": [f"{(count / sum(file_counts.values())) * 100:.2f}%" for count in file_counts.values()]
            })
            st.dataframe(file_df, use_container_width=True)
        else:
            st.info("No files have been downloaded yet.")

    with tabs[2]:
        st.subheader("Log History")
        
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r') as log_file:
                log_lines = log_file.readlines()
                for line in log_lines[-20:]:  # Show last 20 log entries
                    st.text(line.strip())
        else:
            st.info("No log history available.")

# =====================================================================
# Main Application Logic
# =====================================================================

def main():
    """Main function to run the Streamlit app."""
    st.set_page_config(page_title="NSE Report Downloader", page_icon="📈", layout="wide")
    
    # Initialize session state
    if 'page' not in st.session_state:
        st.session_state.page = "login"
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False

    # Set background color
    st.markdown(
        """
        <style>
        .reportview-container {
            background-color: #E6E6FA;  /* Light purple */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Sidebar navigation
    st.sidebar.title("Navigation")
    if st.session_state.logged_in:
        st.sidebar.button("Dashboard", on_click=lambda: setattr(st.session_state, 'page', 'dashboard'))
        #st.sidebar.button("Scheduler", on_click=lambda: setattr(st.session_state, 'page', 'scheduler'))
        st.sidebar.button("Logout", on_click=logout)
    else:
        st.sidebar.button("Login", on_click=lambda: setattr(st.session_state, 'page', 'login'))
        st.sidebar.button("Register", on_click=lambda: setattr(st.session_state, 'page', 'register'))
        st.sidebar.button("Forgot Password", on_click=lambda: setattr(st.session_state, 'page', 'forgot_password'))

    # Page routing
    if st.session_state.page == "login":
        display_login_page()
    elif st.session_state.page == "register":
        display_register_page()
    elif st.session_state.page == "forgot_password":
        display_forgot_password_page()
    elif st.session_state.page == "scheduler":
        display_scheduler()
    elif st.session_state.page == "dashboard":
        display_dashboard()

if _name_ == "_main_":
    main()