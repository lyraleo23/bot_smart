# 🤖 Olist Tiny → Smartphar Integration Bot

## 📋 Description

This project aims to **automate the insertion of e-commerce orders into the Smartphar ERP**,  
replacing the manual order entry process.  

🔹 **Workflow:**  
1. The bot reads the **order spreadsheet from the Olist Tiny ERP**.  
2. It automatically inserts the orders into **Smartphar**, ensuring that production is scheduled correctly.  
3. Eliminates the need for manual input, optimizing time and reducing operational errors.  

## 🛠️ Technologies Used

- **Python** → Main programming language.  
- **Main Libraries**:
  - `pandas` → Data manipulation and analysis of the order spreadsheet.  
  - `pyautogui` → GUI automation to interact with the Smartphar system.  
  - `openpyxl` → Reading and writing Excel files.  
  - `tkinter` → Graphical user interface for better usability.  
  - `dotenv` → Environment variable management.  

## 🚀 How to Use the Project

### Step 1: Clone the Repository
```bash
git clone https://github.com/lyraleo23/bot_smart.git
cd bot_smart
```

### Step 2: Install Dependencies
Make sure Python is installed. Then, install the required packages:
```bash
pip install -r requirements.txt
```

### Step 3: Configure Environment Variables
Create a `.env` file and set up your Smartphar login credentials:
```bash
SMARTPHAR_USER=your_username
SMARTPHAR_PASSWORD=your_password
```

### Step 4: Run the Bot
```bash
python main.py
```

### Step 5: Verify the Orders
The orders will be automatically registered in Smartphar and available for production.


 ## 🧠 Applied Concepts

- **Process Automation** → Reducing manual work in order entry.  
- **GUI Automation with PyAutoGUI** → Simulating user interactions to enter orders in Smartphar.  
- **Data Processing** → Extracting information from the order spreadsheet.  
- **User Interface (Tkinter)** → Creating an interactive way to run the bot.  

## 📄 License

This project is licensed under the MIT License.

## 🤝 Contributions

Contributions are welcome! If you find improvements or need to report issues,  
feel free to open issues or pull requests.

## 📞 Contact

- **Author**: Leonardo Lyra  
- **GitHub**: [lyraleo23](https://github.com/lyraleo23)  
- **LinkedIn**: [Leonardo Lyra](https://www.linkedin.com/in/leonardo-lyra/)  

