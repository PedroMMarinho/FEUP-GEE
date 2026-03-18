import pandas as pd
import matplotlib.pyplot as plt

# 1. Load ALL sheets from the Excel file
file_path = 'UV5545-XLS-ENG.xlsx'
all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)
bg_color = "#eef3db"

# 2. Bulletproof extraction function
def get_yearly_trend(sheets_dict, search_term):
    for sheet_name, df in sheets_dict.items():
        for col in df.columns:
            mask = df[col].astype(str).str.strip().str.match(search_term, case=False)
            if mask.any():
                row = df[mask].iloc[0]
                nums = pd.to_numeric(row, errors='coerce').dropna().values
                if len(nums) >= 4:
                    return [float(nums[i]) * 1000 for i in range(4)]
    raise ValueError(f"Could not find '{search_term}' with 4 years of data!")

# 3. Extracting the 4-year trends
years = ['2012', '2013', '2014', '2015']
cash = get_yearly_trend(all_sheets, r'^Cash')
inv = get_yearly_trend(all_sheets, r'^Inventory1')
ar = get_yearly_trend(all_sheets, r'^Accounts receivable')
ap = get_yearly_trend(all_sheets, r'^Accounts payable')
revenue = get_yearly_trend(all_sheets, r'^Revenue')

# 4. Slide-matched Color Palette 
colors = {
    'Cash': '#2e581d',        
    'Inventory': '#899e54',   
    'AR': '#2B5B84',          
    'AP': '#dcc09e',          
    'Revenue': '#435823'      
}

# Helper function to remove the top/right borders and clean up the chart
def clean_spines(ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cccccc')
    ax.spines['bottom'].set_color('#cccccc')
    ax.tick_params(axis='both', colors='#333333')

# ---------------------------------------------------------
# CHART 1: Current Assets 
# ---------------------------------------------------------
fig1, ax1 = plt.subplots(figsize=(10, 6))
fig1.patch.set_facecolor(bg_color)
ax1.set_facecolor(bg_color)
ax1.plot(years, inv, marker='o', markersize=8, linewidth=3, color=colors['Inventory'], label='Inventory')
ax1.plot(years, ar, marker='s', markersize=8, linewidth=3, color=colors['AR'], label='Accounts Receivable')
ax1.plot(years, cash, marker='D', markersize=8, linewidth=3, color=colors['Cash'], label='Cash Balance')

# Tighter text labels next to the points
for i in range(4):
    ax1.annotate(f"${int(inv[i]):,}", (years[i], inv[i]), textcoords="offset points", xytext=(0, 8), ha='center', fontsize=9, color=colors['Inventory'], fontweight='bold')
    if i == 2:
        offset1 = (0, -11) 
        offset2 = (0, 6)
    else:
        offset1 = (0, -16)
        offset2 = (0, 8)
    ax1.annotate(f"${int(ar[i]):,}", (years[i], ar[i]), textcoords="offset points", xytext=offset1, ha='center', fontsize=9, color=colors['AR'], fontweight='bold')
    ax1.annotate(f"${int(cash[i]):,}", (years[i], cash[i]), textcoords="offset points", xytext=offset2, ha='center', fontsize=9, color=colors['Cash'], fontweight='bold')

# Simplified Title and disabled grid
ax1.set_title('Current Assets', fontsize=16, fontweight='bold', pad=20)
ax1.set_ylabel('Dollars ($)', fontsize=12, fontweight='bold')
ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: f"${int(x):,}"))
ax1.grid(False) 
clean_spines(ax1)

# Clean legend without a harsh box
ax1.legend(fontsize=11, frameon=False)
plt.tight_layout()
fig1.savefig('current_assets_chart_clean.png', dpi=300)

# ---------------------------------------------------------
# CHART 2: Revenue vs. Payables
# ---------------------------------------------------------
fig2, ax2 = plt.subplots(figsize=(10, 6))
fig2.patch.set_facecolor(bg_color)
ax2.set_facecolor(bg_color)
ax2.plot(years, revenue, marker='*', markersize=12, linewidth=3, color=colors['Revenue'], label='Revenue')
ax2.plot(years, ap, marker='^', markersize=10, linewidth=3, color=colors['AP'], label='Accounts Payable')

# Tighter text labels next to the points
for i in range(4):
    ax2.annotate(f"${int(revenue[i]):,}", (years[i], revenue[i]), textcoords="offset points", xytext=(0, 8), ha='center', fontsize=9, color=colors['Revenue'], fontweight='bold')
    ax2.annotate(f"${int(ap[i]):,}", (years[i], ap[i]), textcoords="offset points", xytext=(0, -14), ha='center', fontsize=9, color=colors['AP'], fontweight='bold')

# Simplified Title and disabled grid
ax2.set_title('Revenue vs. Payables', fontsize=16, fontweight='bold', pad=20)
ax2.set_ylabel('Dollars ($)', fontsize=12, fontweight='bold')
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: f"${int(x):,}"))
ax2.grid(False)
clean_spines(ax2)

# Clean legend without a harsh box
ax2.legend(fontsize=11, frameon=False)
plt.tight_layout()
fig2.savefig('revenue_ap_chart_clean.png', dpi=300)


import matplotlib.pyplot as plt

# Data from Exhibit 2
years = ['2012', '2013', '2014', '2015']

inv_days = [424.2, 432.1, 436.5, 476.3]
inv_bench = 386.3

rec_days = [41.9, 45.0, 48.0, 50.9]
rec_bench = 21.8

pay_days = [15.6, 13.3, 10.2, 9.9]
pay_bench = 26.9

# Slide-matched Color Palette
colors = {
    'Inventory': '#899e54',   
    'AR': '#2B5B84',          
    'AP': '#dcc09e',          
    'Benchmark': '#E74C3C' 
}

def clean_spines(ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cccccc')
    ax.spines['bottom'].set_color('#cccccc')
    ax.tick_params(axis='both', colors='#333333')

# ---------------------------------------------------------
# CHART 1: Professional Asset Management
# ---------------------------------------------------------
fig1, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
fig1.patch.set_facecolor(bg_color)
ax1.set_facecolor(bg_color)
ax2.set_facecolor(bg_color)

# Subplot A: Inventory Days
ax1.bar(years, inv_days, color=colors['Inventory'], alpha=0.9)
ax1.axhline(y=inv_bench, color=colors['Benchmark'], linestyle='--', linewidth=2.5, label=f'Industry Benchmark ({inv_bench})')
for i, v in enumerate(inv_days):
    ax1.text(i, v + 5, f"{v:.1f}", ha='center', fontweight='bold', color='#333333')
ax1.set_title('Inventory Days', fontsize=14, fontweight='bold')
ax1.legend(loc='lower right')
ax1.set_ylabel('Days', fontsize=12, fontweight='bold')
clean_spines(ax1)

# Subplot B: Receivable Days
ax2.bar(years, rec_days, color=colors['AR'], alpha=0.9)
ax2.axhline(y=rec_bench, color=colors['Benchmark'], linestyle='--', linewidth=2.5, label=f'Industry Benchmark ({rec_bench})')
for i, v in enumerate(rec_days):
    ax2.text(i, v + 1, f"{v:.1f}", ha='center', fontweight='bold', color='#333333')
ax2.set_title('Receivable Days', fontsize=14, fontweight='bold')
ax2.legend(loc='lower right')
ax2.set_ylabel('Days', fontsize=12, fontweight='bold')
clean_spines(ax2)

# Professional Title
plt.suptitle('Working Capital Efficiency: Days Inventory & Receivables vs. 2014 Benchmark', fontsize=16, fontweight='bold', y=1.05)
plt.tight_layout()
fig1.savefig('days_asset_benchmark_professional.png', dpi=300, bbox_inches='tight')

# ---------------------------------------------------------
# CHART 2: Professional Liability Management
# ---------------------------------------------------------
fig2, ax3 = plt.subplots(figsize=(10, 5))
fig2.patch.set_facecolor(bg_color)
ax3.set_facecolor(bg_color)

ax3.bar(years, pay_days, color=colors['AP'], alpha=0.9, width=0.6)
ax3.axhline(y=pay_bench, color=colors['Benchmark'], linestyle='--', linewidth=2.5, label=f'Industry Benchmark ({pay_bench})')
for i, v in enumerate(pay_days):
    ax3.text(i, v + 0.3, f"{v:.1f}", ha='center', fontweight='bold', color='#333333')

# Professional Title
ax3.set_title('Liability Management: Accounts Payable Days vs. 2014 Benchmark', fontsize=16, fontweight='bold', pad=20)
ax3.set_ylabel('Days', fontsize=12, fontweight='bold')
ax3.legend(loc='upper right')
clean_spines(ax3)

plt.tight_layout()
fig2.savefig('days_payable_benchmark_professional.png', dpi=300)