# Excel VBA Automation Projects

A collection of Excel VBA automation exercises completed as part of my coursework. The projects demonstrate how VBA macros can be used to automate spreadsheet formatting, create interactive controls, and organize dataset records.

## 📌 Project Overview

This project consists of two Excel VBA exercises based on structured dataset records.

The main objectives were to:

- Automate Excel formatting using VBA macros
- Create and execute macros through form-control buttons
- Format spreadsheet headers programmatically
- Sort records automatically using VBA
- Work with Excel's Developer tools and Form Controls

---

## 🛠️ Technologies & Tools

- **Microsoft Excel**
- **VBA (Visual Basic for Applications)**
- Excel Macros
- Form Controls
- Excel Developer Tools

---

# 📂 Project 1 — Header Automation

### Objective

Create a VBA macro that automatically generates and formats a header for a given dataset.

### Header Fields

The required header contains the following fields:

```text
show_id
type
title
director
cast
country
date_added
release_year
rating
duration
```

The header must be formatted using:

- **Font:** Calibri
- **Size:** 12
- **Style:** Bold

### Macro Functionality

The VBA macro automates the process of creating the required header and applying the specified formatting.

### Interactive Button

A Form Control button was also created to execute the macro.

The button can be configured through:

**Developer → Insert → Button (Form Control)**

The macro is then assigned to the button so that the operation can be performed without manually opening the VBA editor.

---

# 📊 Project 2 — Record Sorting Automation

### Objective

Create VBA procedures and corresponding buttons to automatically sort the dataset according to different criteria.

### Sorting Operations

#### 1. Sort by Release Year

Records are sorted:

**Descending by `release_year`**

This places the most recent release years first.

#### 2. Sort by Date Added

Records are sorted:

**Ascending by `date_added`**

This organizes the records from the oldest date added to the newest.

### Interactive Controls

A separate button is created for each sorting operation.

The buttons are named according to their respective functions, allowing the user to perform each sorting operation directly from the Excel worksheet.

---

## 📁 Project Structure

```text
VBA-Projects/
│
├── README.md
├── .gitignore
│
├── Project-1-Header-Automation/
│   ├── Add_header.xlsm
│   ├── Add_header_module.bas
│   └── screenshots/
│       ├── before.png
│       ├── after.png
│       └── button-setup.png
│
└── Project-2-Record-Sorting/
    ├── Sort_records.xlsm
    ├── Sort_records_module.bas
    └── screenshots/
        ├── sort-by-year.png
        ├── sort-by-date.png
        └── buttons-overview.png
```

The `.bas` files are plain-text exports of the VBA modules (via the VBA editor's **File → Export File**), included so the macro code can be reviewed directly on GitHub without opening Excel or enabling macros.

---

## 🎯 Skills Demonstrated

Through these exercises, I practiced:

- Excel VBA programming
- Macro creation
- Excel worksheet manipulation
- Automated cell formatting
- Data sorting with VBA
- Creating Excel Form Controls
- Assigning macros to buttons
- Automating repetitive spreadsheet operations
- Working with structured datasets

---

## 📚 Academic Context

This project was completed as part of my coursework in **VBA/Excel automation**.

The assignment provided practical exercises for applying VBA to common spreadsheet automation tasks, including formatting, user interaction, and data organization.

The VBA implementations in this repository represent my own work based on the requirements of the coursework.

---

## 🚀 Future Improvements

Possible improvements to the project include:

- Adding error handling to the VBA procedures
- Dynamically detecting the dataset range
- Creating a single user interface/dashboard for all operations
- Adding additional sorting and filtering options
- Automating data validation
- Adding logging or status messages
- Making the macros reusable with different datasets

---

## 👨‍💻 Author

**Habib Akinola**

This project forms part of my growing portfolio in **automation, programming, data handling, and IT systems**.
