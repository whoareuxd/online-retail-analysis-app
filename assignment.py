import pandas as pd
from datetime import datetime
from colorama import Fore, Style, init

# فعال کردن رنگ‌ها در ویندوز و ترمینال‌های پشتیبانی‌شده
init(autoreset=True)

# بارگذاری فایل اکسل
try:
    df = pd.read_excel("onlineRetail.xlsx")
except FileNotFoundError:
    print(Fore.RED + "File 'onlineRetail.xlsx' not found. Please check the path and try again.")
    exit()

# تبدیل ستون تاریخ به datetime
df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce")

# حذف ردیف‌هایی که در ستون‌های مهم داده معتبری ندارند
df.dropna(subset=["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID"], inplace=True)

# فیلتر کردن داده‌ها: فقط ردیف‌هایی با تعداد مثبت
df = df[df["Quantity"] > 0]

# محاسبه درآمد هر ردیف (فروش هر محصول)
df["Revenue"] = df["Quantity"] * df["UnitPrice"]

# شروع منو و گزینه‌ها
def menu():
    print(Fore.LIGHTYELLOW_EX + "\n--- Menu ---")
    print(Fore.LIGHTCYAN_EX + "1. Top 5 cities with the most orders")
    print("2. Top 5 products by quantity sold")
    print("3. Top 5 brands by average unit price")
    print("4. Search for product (by ID or name)")
    print("5. Show orders by a specific customer (CustomerID)")
    print("6. Show products purchased on a specific date")
    print("7. Show total revenue of the store")
    print("8. Show most frequent day of week for orders")
    print("9. Exit\n")

def top_cities():
    # گروه‌بندی بر اساس کشور 
    city_counts = df.groupby("Country")["InvoiceNo"].nunique().sort_values(ascending=False).head(5)
    print(Fore.LIGHTGREEN_EX + "Top 5 countries by number of unique orders:")
    print(city_counts)

def top_products():
    product_qty = df.groupby("Description")["Quantity"].sum().sort_values(ascending=False).head(5)
    print(Fore.LIGHTGREEN_EX + "Top 5 products by quantity sold:")
    print(product_qty)

def top_brands():
    # فرض بر این است که برندها در ستون Brand وجود دارند
    if "Brand" not in df.columns:
        print(Fore.RED + "Brand information is not available in the dataset.")
        return
    avg_price = df.groupby("Brand")["UnitPrice"].mean().sort_values(ascending=False).head(5)
    print(Fore.LIGHTGREEN_EX + "Top 5 Brands by average unit price:")
    print(avg_price)

def search_product():
    print(Fore.LIGHTMAGENTA_EX + "Search product by:")
    print("1. StockCode (ID)")
    print("2. Product Name")
    choice = input("Your choice: ").strip()
    if choice == "1":
        code = input("Enter StockCode: ").strip()
        product_names = df[df["StockCode"] == code]["Description"].unique()
        if len(product_names) == 0:
            print(Fore.RED + "Product not found.")
        else:
            print(Fore.LIGHTGREEN_EX + "Product name(s):")
            for name in product_names:
                print(name)
    elif choice == "2":
        name = input("Enter part or full product name (case insensitive): ").strip().lower()
        matches = df[df["Description"].str.lower().str.contains(name)]["StockCode"].unique()
        if len(matches) == 0:
            print(Fore.RED + "No product matches found.")
        else:
            print(Fore.LIGHTGREEN_EX + "Matching StockCodes:")
            for c in matches:
                print(c)
    else:
        print(Fore.RED + "Invalid choice.")

def show_orders_customer():
    try:
        customer_id = int(input("Enter CustomerID: ").strip())
    except ValueError:
        print(Fore.RED + "Invalid CustomerID format.")
        return
    cust_orders = df[df["CustomerID"] == customer_id]
    if cust_orders.empty:
        print(Fore.RED + "No orders found for this CustomerID.")
    else:
        print(Fore.LIGHTGREEN_EX + f"Orders for CustomerID {customer_id}:")
        # نمایش فقط ستون‌های مهم
        print(cust_orders[["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice"]].to_string(index=False))

def show_products_by_date():
    date_str = input("Enter date (YYYY-MM-DD): ").strip()
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        print(Fore.RED + "Invalid date format.")
        return
    date_orders = df[df["InvoiceDate"].dt.date == date_obj.date()]
    if date_orders.empty:
        print(Fore.RED + f"No orders found on {date_str}.")
    else:
        print(Fore.LIGHTGREEN_EX + f"Orders on {date_str}:")
        print(date_orders[["InvoiceNo", "StockCode", "Description", "Quantity", "UnitPrice"]].to_string(index=False))

def total_revenue():
    total = df["Revenue"].sum()
    print(Fore.LIGHTGREEN_EX + f"Total revenue of the store: {total:,.2f}")

def frequent_day_of_week():
    # اضافه کردن ستون روز هفته (Monday=0,... Sunday=6)
    df["DayOfWeek"] = df["InvoiceDate"].dt.day_name()
    freq = df.groupby("DayOfWeek")["InvoiceNo"].nunique().sort_values(ascending=False)
    print(Fore.LIGHTGREEN_EX + "Number of unique orders by day of the week:")
    print(freq)

# اجرای منو و دریافت ورودی کاربر
while True:
    menu()
    choice = input(Fore.LIGHTBLUE_EX + "Choose an option: ").strip()
    print(Style.RESET_ALL)
    if choice == "1":
        top_cities()
    elif choice == "2":
        top_products()
    elif choice == "3":
        top_brands()
    elif choice == "4":
        search_product()
    elif choice == "5":
        show_orders_customer()
    elif choice == "6":
        show_products_by_date()
    elif choice == "7":
        total_revenue()
    elif choice == "8":
        frequent_day_of_week()
    elif choice == "9":
        print(Fore.YELLOW + "Exiting... Goodbye!")
        break
    else:
        print(Fore.RED + "Invalid choice, please try again.")
