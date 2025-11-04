# Correlation Analysis Script - User Guide

## Overview
This script performs Pearson's correlation test and creates Bland-Altman plots for all nerve level measurements in your Final_Fa_Vals spreadsheet.

## Requirements
```bash
pip install pandas numpy matplotlib scipy openpyxl
```

## Usage

### Basic Usage
```bash
python correlation_analysis.py <your_excel_file.xlsx>
```

Example:
```bash
python correlation_analysis.py Final_Fa_Vals_Complete.xlsx
```

This creates a folder called `correlation_results/` with all outputs.

### Specify Custom Output Directory
```bash
python correlation_analysis.py <your_excel_file.xlsx> <output_directory>
```

Example:
```bash
python correlation_analysis.py Final_Fa_Vals_Complete.xlsx my_analysis
```


For each measurement, it:
1. **Extracts valid data** (excludes 'x' values and patients 12, 14)
2. **Calculates Pearson's r** and p-value
3. **Creates a scatter plot** with correlation line
4. **Creates a Bland-Altman plot** showing agreement
5. **Compiles statistics** into reports

## Output Files

### In `correlation_results/` folder:

1. **Scatter Plots** (`*_scatter.png`)
   - Shows correlation between the two methods
   - Includes line of best fit
   - Displays r, p-value, and R² statistics

2. **Bland-Altman Plots** (`*_bland_altman.png`)
   - Shows agreement between methods
   - Mean difference line
   - 95% limits of agreement (±1.96 SD)

3. **correlation_summary.csv**
   - Quick reference table with all results
   - Columns: Measurement, n, r, p-value, r², interpretation

4. **correlation_report.txt**
   - Detailed text report
   - Summary table
   - Interpretation guide
   - Individual results with 95% limits of agreement

## Understanding the Results

### Pearson's r (correlation coefficient)
- **+1.0**: Perfect positive correlation
- **+0.8 to +1.0**: Very strong positive
- **+0.6 to +0.8**: Strong positive
- **+0.4 to +0.6**: Moderate positive
- **+0.2 to +0.4**: Weak positive
- **0.0 to +0.2**: Very weak
- **(Negative values indicate inverse relationships)**

### P-value (significance)
- **p < 0.001**: *** Highly significant
- **p < 0.01**: ** Very significant
- **p < 0.05**: * Significant
- **p ≥ 0.05**: ns (not significant)

### R² (coefficient of determination)
- Proportion of variance explained (0 to 1)
- Example: R² = 0.60 means 60% of variance is explained

### Bland-Altman Plot Interpretation
- **Points clustered near zero**: Good agreement
- **Points widely scattered**: Poor agreement
- **Points outside 95% limits**: Potential outliers
- **Systematic bias**: Mean difference far from zero

## Clinical Interpretation

**Strong correlations (r ≥ 0.60)**: The two measurement methods show good agreement and either can be used reliably.

**Moderate correlations (r = 0.40-0.59)**: Methods show some agreement but may reflect different aspects of the measurement or have greater variability.

**Weak correlations (r < 0.40)**: Methods may be measuring different things or have high measurement error. Use caution when comparing.

## Troubleshooting

### "Not enough valid data points"
- Need at least 3 patients with complete data for both methods
- Check for 'x' values (missing data markers)
- Check for NaN/blank values

### Script fails to run
- Make sure you have all required packages installed
- Check that the Excel file path is correct
- Ensure the Excel file has the expected structure

### Unexpected results
- Verify column names match expected format
- Check that row 0 contains headers
- Ensure data starts at row 1 (after header)


**Line 143-158**: Define which measurements to analyze
```python
measurements = [
    ('C5 whole cord', 'Unnamed: 2', 'C5_Whole', 'C5 Whole Cord'),
    ...
]
```

**get_valid_pairs()**: Controls which data is included/excluded

**create_scatter_plot()** and **create_bland_altman_plot()**: Customize visualizations


If using this analysis in publications, consider citing:
- Pearson, K. (1895). Note on regression and inheritance in the case of two parents. Proceedings of the Royal Society of London.
- Bland, J. M., & Altman, D. G. (1986). Statistical methods for assessing agreement between two methods of clinical measurement. The Lancet, 327(8476), 307-310.
