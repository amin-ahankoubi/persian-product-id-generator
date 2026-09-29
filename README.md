# Persian Product ID Generator

A desktop application for generating and managing structured product identifiers with Persian/Jalali date support.

This application is designed to simplify the process of generating unique product identifiers, storing related descriptions, managing records, archiving old records, filtering data, and exporting information to Excel.

The application is built with **Python**, **Tkinter**, and **SQLite**, with support for Persian user interfaces and Jalali dates.

---

## ✨ Features

* Generate sequential product identifiers
* Support for Persian/Jalali dates
* Configurable product codes and names
* Local SQLite database
* Automatic sequential numbering per product and year
* Mandatory description for each generated identifier
* Search and filtering
* Filter by:

  * Identifier
  * Description
  * Date range
  * Product
* Archive records
* Restore archived records
* Copy generated identifiers to clipboard
* Copy multiple selected identifiers
* Export records to Excel
* Separate active and archived records
* Database reset and management options
* Persistent local application data
* Persian/Farsi user interface
* Windows desktop application support
* Compatible with PyInstaller-based executable builds

---

## 🖥️ Application Overview

The application generates identifiers using the following structure:

```text
YEAR / PRODUCT_CODE / SEQUENCE
```

For example:

```text
1405/110101/0001
1405/110101/0002
1405/210201/0001
```

The sequence number is maintained separately for each product and year.

The application prevents the sequence from exceeding `9999` for a product within the configured year.

---

## ⚙️ How It Works

When a user selects a product and requests a new identifier, the application:

1. Reads the configured system year.
2. Finds the latest sequence number for the selected product.
3. Generates the next sequence number.
4. Creates the complete identifier.
5. Requests a mandatory description.
6. Asks the user to confirm the registration.
7. Stores the record in the local SQLite database.

Example:

```text
Year:         1405
Product Code: 110101
Sequence:     0007

Generated ID:
1405/110101/0007
```

---

## 🗄️ Database

The application uses **SQLite** as its local database.

The database is automatically created on the user's system and contains the following main entities:

* `settings`
* `products`
* `records`

Application data is stored inside the user's local application data directory.

On Windows, the default location is:

```text
%LOCALAPPDATA%\PersianProductID\
```

The database file is:

```text
records.sqlite3
```

The database is created automatically when the application starts.

---

## 📊 Excel Export

Records can be exported to an `.xlsx` file.

The generated Excel file contains:

| Column  | Description          |
| ------- | -------------------- |
| شناسه   | Generated identifier |
| محصول   | Product name         |
| توضیحات | Record description   |
| تاریخ   | Creation date        |
| وضعیت   | Active or archived   |

The Excel worksheet is also configured for right-to-left presentation.

---

## 🗃️ Archive Management

Records can be moved from the main list to the archive.

Archived records can:

* Be searched and filtered
* Be exported to Excel
* Have their identifiers copied
* Be restored to the main list

This allows the application to maintain a clear separation between active and historical records.

---

## ⚙️ System Settings

The application provides a settings section where administrators can configure:

* System start year
* Product codes
* Product names

Product codes must be between 1 and 10 characters and product names cannot be empty.

---

## 🛠️ Technologies

| Technology  | Purpose                      |
| ----------- | ---------------------------- |
| Python      | Application development      |
| Tkinter     | Desktop GUI                  |
| SQLite      | Local database               |
| `jdatetime` | Jalali/Persian date handling |
| `openpyxl`  | Excel export                 |
| PyInstaller | Executable packaging         |

Python standard-library modules such as `sqlite3`, `os`, `sys`, and `tkinter` are also used.

---

## 📦 Installation

### Requirements

* Python 3.10 or newer
* Windows is currently the primary target platform

Clone the repository:

```bash
git clone https://github.com/amin-ahankoubi/persian-product-id-generator.git
```

Enter the project directory:

```bash
cd persian-product-id-generator
```

Install the required packages:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
python product_id_generator.py
```

---

## 📋 Dependencies

The project currently depends on:

```text
jdatetime
openpyxl
```

The following components are part of the Python standard library:

```text
os
sys
sqlite3
tkinter
```

---

## 🏗️ Project Structure

```text
persian-product-id-generator/
│
├── product_id_generator.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
└── logo.ico
```

> `logo.ico` is used as the application's window icon and should be included in the repository if it is part of the project distribution.

---

## 🔐 Security Note

The application currently contains an administrative reset mechanism.

Before publishing this repository as a public project, avoid storing passwords or other sensitive credentials directly in the source code.

For production use, the reset credential should be moved to a secure configuration mechanism or replaced with a safer authentication approach.

---

## 🚀 Future Improvements

Possible future improvements include:

* User authentication and role management
* Configurable application themes
* Improved Persian typography and UI
* Database backup and restore
* CSV import/export
* Advanced reporting
* Automatic database backup
* Configurable identifier formats
* Multi-user/network database support
* Application update mechanism
* Installer package for Windows
* Comprehensive logging
* Unit and integration tests

---

## 👨‍💻 Author

**Amin Ahankoubi**

Software Developer

GitHub: [amin-ahankoubi](https://github.com/amin-ahankoubi)

---

## 📄 License

This project is licensed under the MIT License.

See the [LICENSE](LICENSE) file for details.
