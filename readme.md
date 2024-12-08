# Investment and Financial Tools

This repository contains a collection of Python scripts designed to help with financial planning, analysis, and decision-making. Each script is tailored for a specific purpose, ranging from loan calculations to retirement planning.

---

## Tools

### 1. **Auto Loan Calculator**  
**File:** `auto_loan.py`  
Calculate the monthly payment, total interest, and payoff schedule for an auto loan.  
**Features:**  
- Adjustable loan term and interest rate.  
- Detailed amortization schedule.

### 2. **Budget Planner**  
**File:** `budget_planner.py`  
Allocate income across spending categories and track actual spending against the budget.  
**Features:**  
- Visualizes spending vs. budget with pie and bar charts.  
- Provides detailed breakdown and insights in Excel.

### 3. **Compound Interest Calculator**  
**File:** `compound_interest.py`  
Analyze compound interest growth with options for inflation adjustment and contribution increases.  
**Features:**  
- Flexible contribution frequencies (e.g., monthly, bi-weekly).  
- Calculates real vs. nominal balance.  
- Embedded graphs in Excel output.

### 4. **Debt Payoff Planner**  
**File:** `debt_payoff.py`  
Compare the Snowball and Avalanche methods for paying off debts.  
**Features:**  
- Customizable payoff strategies.  
- Detailed debt schedules with total interest comparisons.

### 5. **Emergency Fund Calculator**  
**File:** `emergency_fund.py`  
Track progress toward building an emergency fund.  
**Features:**  
- Supports multiple contribution frequencies.  
- Calculates how long it will take to reach your goal.  
- Visualizes savings growth.

### 6. **Loan vs. Savings Comparison Tool**  
**File:** `loan_savings_comparison.py`  
Compare the costs and benefits of saving for an expense versus taking out a loan.  
**Features:**  
- Adjustable timeframes for savings and loan terms.  
- Inflation-adjusted savings targets.  
- Graphs comparing loan costs to savings growth.

### 7. **Mortgage Calculator**  
**File:** `mortgage.py`  
Calculate monthly payments, including property taxes, insurance, and PMI, for a mortgage.  
**Features:**  
- Optional extra payments for faster payoff.  
- Detailed amortization schedules.  
- Visualizes interest vs. principal over time.

### 8. **Personal Loan Calculator**  
**File:** `personal_loan.py`  
Analyze monthly payments, total interest, and payoff schedules for personal loans.  
**Features:**  
- Flexible loan term and interest rate.  
- Generates Excel breakdowns.

### 9. **Retirement Savings Planner**  
**File:** `retirement.py`  
Plan how much to save to reach your retirement goals.  
**Features:**  
- Accounts for inflation and annual contribution increases.  
- Breaks down savings progress by contribution frequency.  
- Generates graphs and Excel summaries.

### 10. **Savings Goal Planner**  
**File:** `savings_goal.py`  
Determine the contributions needed to achieve a specific financial goal by a set date.  
**Features:**  
- Flexible contribution frequencies.  
- Inflation-adjusted goals.  
- Provides a detailed progress report in Excel.

### 11. **Stock Growth Calculator**  
**File:** `stock_growth.py`  
Calculate stock growth with reinvested dividends and periodic contributions.  
**Features:**  
- Supports various contribution frequencies.  
- Visualizes investment growth over time.  
- Detailed breakdowns in Excel.

---

## Requirements

- **Python 3.10 or later**
- **Dependencies**:
  - `pandas`
  - `matplotlib`
  - `openpyxl`

Install dependencies using:
```
pip install -r requirements.txt
```

---

## Usage

1. Clone the repository:
```
git clone https://github.com/chiefgyk3d/investment_tools.git
cd investment_tools
```

2. Install dependencies:
```
pip install -r requirements.txt
```

3. Run any script from the command line. For example, to calculate compound interest:
```
python3 compound_interest.py
```

4. Follow the prompts to input your data. Each script will guide you step by step through the required inputs.

5. The results will be saved in an Excel file in the current directory. The file will include:
   - Detailed breakdowns of calculations.
   - Embedded graphs for visualization.

---

## Example

**Using the Budget Planner:**

Run the script:
```
python3 budget_planner.py
```

Follow the prompts:
- Enter your monthly income.
- Specify budget categories and amounts (e.g., "Housing: 1500").
- Enter your actual spending in each category.
- Provide a base name for the output files (e.g., "budget_report").

The output will include:
- An Excel file (`budget_report.xlsx`) with detailed budget vs. actual comparisons.
- Charts embedded in the file for easy visualization.

---

## Contributions

Feel free to contribute by submitting issues or pull requests. Suggestions for new tools or features are always welcome!

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Author

**ChiefGyk3D**  
For more of my work, visit [My GitHub Profile](https://github.com/chiefgyk3d).
