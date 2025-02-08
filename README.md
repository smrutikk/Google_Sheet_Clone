
# Google Sheets Clone

## Description:

Google Sheets Clone is a desktop application built using the Electron framework, incorporating technologies like HTML, CSS, JavaScript, and EJS as a templating engine.  
Some key features of this project are:
- **Cell Operations:** Change cell size, detect cycles in formulas, and evaluate formulas using a stack-based approach.
- **File Management:** New, Open, Save functionality.
- **Formatting:** Customize font, size, bold, italic, underline, color, background color, and text alignment.
- **Workbook Management:** Add multiple sheets to a single workbook.
- **Mathematical Functions:** `SUM`, `AVERAGE`, `MAX`, `MIN`, and `COUNT`.
- **Data Quality Functions:** 
  - `TRIM`: Remove extra spaces.
  - `UPPER` and `LOWER`: Convert text to uppercase or lowercase.
  - `REMOVE_DUPLICATES`: Eliminate duplicate entries.
  - `FIND_AND_REPLACE`: Locate and replace values seamlessly.

---

## Screenshots
![](img/SS_1.png)
![](img/SS_2.png)

---

## Technical Stack:

### Frontend:
- **Programming Language:** JavaScript
- **Framework:** Electron
- **Templating Engine:** EJS
- **IDE:** VSCode

---

## Why Electron Framework?

The Electron framework was chosen for its ability to build cross-platform desktop applications using web technologies like HTML, CSS, and JavaScript. Key reasons include:
1. **Cross-Platform Compatibility:** Electron allows the application to run on Windows, macOS, and Linux without requiring separate development efforts for each platform.
2. **Web Technology Integration:** By leveraging web technologies, the development process was simplified while enabling the integration of rich UI features.
3. **Access to Native Features:** Electron provides access to native OS functionalities such as file handling, enabling seamless implementation of features like `New`, `Open`, and `Save`.
4. **Scalability:** Electron’s architecture allows easy addition of features like multiple sheets and advanced formula operations, ensuring scalability.

---

## Steps to Run Locally

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/google-sheets-clone.git
   cd google-sheets-clone
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Run the application:**
   ```bash
   npm start
   ```

4. **Build the application (optional):**
   ```bash
   npm run build
   ```

---
## Electron Set Up

1. **Ensure presence of Node.js**
   ```bash
   npm -v
   ```
2. **Install Electron Locally**
   ```bash
   npm install --save-dev electron
   ```
---

