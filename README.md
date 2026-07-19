# AutoAMLE

AutoAMLE is a Python-based desktop application developed to automate the preparation of AMLE deployment files for Nokia LTE networks.

The tool reads Nokia LTE MDB dump files, processes the required parameter families, and generates deployment-ready CSV files in under 30 seconds, significantly reducing manual effort and minimizing configuration errors.

---

## Features

### AMLEPR Generation
- Matches the correct AMLE profile for each cell according to configurable source-to-target frequency mappings.
- Prevents duplicate entries.
- Supports an unlimited number of mapping definitions.

### LNREL Generation
- Automatically generates intra-site neighbor relations.
- Configures AMLE-related parameters for each relation.
- Eliminates repetitive manual configuration.

### LC Generation
- Enables AMLE functionality on LTE cells.
- Produces deployment-ready CSV output.

### Intelligent Data Processing
- Reads Nokia LTE MDB dump files.
- Supports Excel-based input configuration.
- Automatically validates and processes the required parameter families.

### User-Friendly Interface
- Simple graphical interface built with Python.
- One-click generation of deployment files.
- Fast processing with clear progress indication.

---

## Technologies Used

- Python
- Tkinter
- Pandas
- PyODBC
- Microsoft Access (MDB)
- Excel Processing

---

## Workflow

```
Nokia LTE MDB Dump
        +
   Excel Configuration
        │
        ▼
     AutoAMLE
        │
        ▼
Deployment-ready CSV Files
├── AMLEPR
├── LNREL
└── LC
```

---

## Benefits

- Reduces deployment preparation time from manual work to under 30 seconds.
- Minimizes human errors.
- Ensures consistent deployment file generation.
- Supports multiple frequency mapping scenarios.
- Improves productivity for LTE optimization and deployment activities.

---

## Screenshots

> you will find the gui screenshot in GitHub 

Example:

- Main Application
- Input Configuration
- Generated CSV Output

---

## Notes

This repository demonstrates the software engineering and automation concepts behind the tool.

No customer data, proprietary configuration data, or confidential network information is included.

---

## Author

**Abdelrahman Mohamed Galal**

RF Optimization Engineer | Python Automation | Data Engineering Enthusiast

LinkedIn:
(https://www.linkedin.com/in/abdelrahman-mohamed-galal-1594201b4/?skipRedirect=true)

GitHub:
https://github.com/AbdelrahmanMohamed105
