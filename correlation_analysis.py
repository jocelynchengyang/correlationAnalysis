"""
Pearson Correlation and Bland-Altman Analysis
Compares measurements from Values.xls vs combined_output.csv for all nerve levels

Usage: python correlation_analysis.py <Final_Fa_Vals_file.xlsx>
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import pearsonr
import sys
import os
from pathlib import Path

def get_valid_pairs(x_series, y_series, pt_id_series):
    """Extract valid numeric pairs from two series, excluding 'x' and NaN values."""
    valid_pairs = []
    pt_ids = []
    
    for i in range(len(x_series)):
        x_val = x_series.iloc[i]
        y_val = y_series.iloc[i]
        pt_id = pt_id_series.iloc[i]
        
        if pd.notna(x_val) and pd.notna(y_val):
            # Skip if x is 'x' (missing marker)
            if isinstance(x_val, str) and x_val.lower() == 'x':
                continue
            try:
                x_float = float(x_val)
                y_float = float(y_val)
                valid_pairs.append((x_float, y_float))
                pt_ids.append(int(pt_id))
            except:
                continue
    
    if len(valid_pairs) > 0:
        x_values = [pair[0] for pair in valid_pairs]
        y_values = [pair[1] for pair in valid_pairs]
        return x_values, y_values, pt_ids
    else:
        return None, None, None

def create_scatter_plot(x_values, y_values, pt_ids, r, p_value, title, x_label, y_label, filename):
    """Create scatter plot with correlation line."""
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # Plot data points
    ax.scatter(x_values, y_values, s=100, alpha=0.6, color='steelblue', 
               edgecolors='black', linewidth=1.5)
    
    # Add labels for each point
    for i, pt_id in enumerate(pt_ids):
        ax.annotate(f'{pt_id}', (x_values[i], y_values[i]), 
                    xytext=(5, 5), textcoords='offset points', fontsize=8, alpha=0.7)
    
    # Add line of best fit
    z = np.polyfit(x_values, y_values, 1)
    p = np.poly1d(z)
    x_line = np.linspace(min(x_values), max(x_values), 100)
    ax.plot(x_line, p(x_line), "r--", linewidth=2, alpha=0.8, label='Line of best fit')
    
    # Add identity line (y=x)
    min_val = min(min(x_values), min(y_values))
    max_val = max(max(x_values), max(y_values))
    ax.plot([min_val, max_val], [min_val, max_val], 'k:', linewidth=1.5, 
            alpha=0.5, label='Perfect agreement (y=x)')
    
    # Labels and title
    ax.set_xlabel(x_label, fontsize=11, fontweight='bold')
    ax.set_ylabel(y_label, fontsize=11, fontweight='bold')
    ax.set_title(title, fontsize=13, fontweight='bold', pad=20)
    
    # Add statistics box
    textstr = f'n = {len(pt_ids)}\nr = {r:.4f}\np = {p_value:.4f}\nR² = {r**2:.4f}'
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.8)
    ax.text(0.05, 0.95, textstr, transform=ax.transAxes, fontsize=10,
            verticalalignment='top', bbox=props, family='monospace')
    
    # Add grid and legend
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.legend(loc='lower right', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    plt.close()

def create_bland_altman_plot(x_values, y_values, pt_ids, title, filename):
    """Create Bland-Altman plot."""
    fig, ax = plt.subplots(figsize=(10, 8))
    
    differences = [y - x for x, y in zip(x_values, y_values)]
    averages = [(x + y) / 2 for x, y in zip(x_values, y_values)]
    mean_diff = np.mean(differences)
    sd_diff = np.std(differences, ddof=1)
    
    # Plot
    ax.scatter(averages, differences, s=100, alpha=0.6, color='coral', 
               edgecolors='black', linewidth=1.5)
    
    # Add labels
    for i, pt_id in enumerate(pt_ids):
        ax.annotate(f'{pt_id}', (averages[i], differences[i]), 
                    xytext=(5, 5), textcoords='offset points', fontsize=8, alpha=0.7)
    
    # Add mean line
    ax.axhline(mean_diff, color='blue', linestyle='-', linewidth=2, 
               label=f'Mean difference: {mean_diff:.4f}')
    
    # Add ±1.96 SD lines (95% limits of agreement)
    upper_limit = mean_diff + 1.96*sd_diff
    lower_limit = mean_diff - 1.96*sd_diff
    ax.axhline(upper_limit, color='red', linestyle='--', linewidth=1.5, 
               label=f'+1.96 SD: {upper_limit:.4f}')
    ax.axhline(lower_limit, color='red', linestyle='--', linewidth=1.5,
               label=f'-1.96 SD: {lower_limit:.4f}')
    
    # Zero line
    ax.axhline(0, color='gray', linestyle=':', linewidth=1, alpha=0.5)
    
    ax.set_xlabel('Average of Two Methods', fontsize=11, fontweight='bold')
    ax.set_ylabel('Difference (SCToolbox - Values.xls)', fontsize=11, fontweight='bold')
    ax.set_title(title, fontsize=13, fontweight='bold', pad=20)
    
    # Add statistics box
    textstr = f'n = {len(pt_ids)}\nMean diff = {mean_diff:.4f}\nSD = {sd_diff:.4f}'
    props = dict(boxstyle='round', facecolor='lightblue', alpha=0.8)
    ax.text(0.05, 0.95, textstr, transform=ax.transAxes, fontsize=10,
            verticalalignment='top', bbox=props, family='monospace')
    
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.legend(loc='best', fontsize=9)
    
    plt.tight_layout()
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    plt.close()

def interpret_correlation(r):
    """Provide interpretation of correlation coefficient."""
    abs_r = abs(r)
    if abs_r >= 0.80:
        strength = "Very strong"
    elif abs_r >= 0.60:
        strength = "Strong"
    elif abs_r >= 0.40:
        strength = "Moderate"
    elif abs_r >= 0.20:
        strength = "Weak"
    else:
        strength = "Very weak"
    
    direction = "positive" if r > 0 else "negative"
    return f"{strength} {direction}"

def run_analysis(input_file, output_dir='correlation_results'):
    """Run complete correlation analysis for all measurements."""
    
    # Create output directory
    Path(output_dir).mkdir(exist_ok=True)
    
    # Read the data
    print("=" * 70)
    print("Pearson Correlation & Bland-Altman Analysis")
    print("=" * 70)
    print(f"\nReading data from: {input_file}")
    
    df = pd.read_excel(input_file)
    
    # Skip header row (row 0) for data
    pt_ids = df['Pt ID'].iloc[1:]
    
    # Define the measurements to analyze
    # Format: (original_col, sctoolbox_col, short_name, long_name)
    measurements = [
        ('C5 whole cord', 'Unnamed: 2', 'C5_Whole', 'C5 Whole Cord'),
        ('C5 R hemicord', 'Unnamed: 4', 'C5_Right', 'C5 Right Hemicord'),
        ('C5 L hemicord', 'Unnamed: 6', 'C5_Left', 'C5 Left Hemicord'),
        ('C6 whole cord', 'Unnamed: 8', 'C6_Whole', 'C6 Whole Cord'),
        ('C6 R hemicord', 'Unnamed: 10', 'C6_Right', 'C6 Right Hemicord'),
        ('C6 L hemicord', 'Unnamed: 12', 'C6_Left', 'C6 Left Hemicord'),
        ('C7 whole cord', 'Unnamed: 14', 'C7_Whole', 'C7 Whole Cord'),
        ('C7 R hemicord', 'Unnamed: 16', 'C7_Right', 'C7 Right Hemicord'),
        ('C7 L hemicord', 'Unnamed: 18', 'C7_Left', 'C7 Left Hemicord'),
        ('C8 whole cord', 'Unnamed: 20', 'C8_Whole', 'C8 Whole Cord'),
        ('C8 R hemicord', 'Unnamed: 22', 'C8_Right', 'C8 Right Hemicord'),
        ('C8 L hemicord', 'Unnamed: 24', 'C8_Left', 'C8 Left Hemicord'),
        ('T1 whole cord', 'Unnamed: 26', 'T1_Whole', 'T1 Whole Cord'),
        ('T1 R hemicord', 'Unnamed: 28', 'T1_Right', 'T1 Right Hemicord'),
        ('T1 L hemicord', 'Unnamed: 30', 'T1_Left', 'T1 Left Hemicord'),
    ]
    
    # Store results
    results = []
    
    print(f"\nAnalyzing {len(measurements)} measurements...")
    print("=" * 70)
    
    for orig_col, sct_col, short_name, long_name in measurements:
        print(f"\nProcessing: {long_name}")
        
        # Get data
        x_series = df[orig_col].iloc[1:]
        y_series = df[sct_col].iloc[1:]
        
        # Get valid pairs
        x_values, y_values, valid_pt_ids = get_valid_pairs(x_series, y_series, pt_ids)
        
        if x_values is None or len(x_values) < 3:
            print(f"  ⚠ Skipped: Not enough valid data points (need at least 3)")
            continue
        
        # Calculate Pearson correlation
        r, p_value = pearsonr(x_values, y_values)
        r_squared = r ** 2
        
        # Store results
        results.append({
            'Measurement': long_name,
            'n': len(valid_pt_ids),
            'r': r,
            'p_value': p_value,
            'r_squared': r_squared,
            'interpretation': interpret_correlation(r),
            'mean_diff': np.mean([y - x for x, y in zip(x_values, y_values)]),
            'sd_diff': np.std([y - x for x, y in zip(x_values, y_values)], ddof=1)
        })
        
        # Print results
        print(f"  n = {len(valid_pt_ids)}")
        print(f"  r = {r:.4f} ({interpret_correlation(r)} correlation)")
        print(f"  p = {p_value:.6f} {'***' if p_value < 0.001 else '**' if p_value < 0.01 else '*' if p_value < 0.05 else 'ns'}")
        print(f"  R² = {r_squared:.4f}")
        
        # Create plots
        scatter_file = os.path.join(output_dir, f'{short_name}_scatter.png')
        ba_file = os.path.join(output_dir, f'{short_name}_bland_altman.png')
        
        create_scatter_plot(
            x_values, y_values, valid_pt_ids, r, p_value,
            f'Pearson Correlation: {long_name} FA',
            'Values.xls',
            'SCToolbox (combined_output)',
            scatter_file
        )
        
        create_bland_altman_plot(
            x_values, y_values, valid_pt_ids,
            f'Bland-Altman Plot: {long_name} FA',
            ba_file
        )
        
        print(f"  ✓ Plots saved")
    
    # Create summary report
    print("\n" + "=" * 70)
    print("Creating summary report...")
    
    results_df = pd.DataFrame(results)
    
    # Save to CSV
    csv_file = os.path.join(output_dir, 'correlation_summary.csv')
    results_df.to_csv(csv_file, index=False)
    print(f"✓ Summary CSV saved: {csv_file}")
    
    # Create detailed text report
    report_file = os.path.join(output_dir, 'correlation_report.txt')
    with open(report_file, 'w') as f:
        f.write("=" * 70 + "\n")
        f.write("PEARSON CORRELATION & BLAND-ALTMAN ANALYSIS\n")
        f.write("Comparison: Values.xls vs SCToolbox (combined_output.csv)\n")
        f.write("=" * 70 + "\n\n")
        
        f.write("SUMMARY TABLE\n")
        f.write("-" * 70 + "\n")
        f.write(f"{'Measurement':<25} {'n':>4} {'r':>8} {'p-value':>10} {'R²':>8} {'Interpretation':<20}\n")
        f.write("-" * 70 + "\n")
        
        for _, row in results_df.iterrows():
            sig = '***' if row['p_value'] < 0.001 else '**' if row['p_value'] < 0.01 else '*' if row['p_value'] < 0.05 else 'ns'
            f.write(f"{row['Measurement']:<25} {row['n']:>4} {row['r']:>8.4f} "
                   f"{row['p_value']:>10.6f} {row['r_squared']:>8.4f} {row['interpretation']:<20}\n")
        
        f.write("\n" + "=" * 70 + "\n")
        f.write("INTERPRETATION GUIDE\n")
        f.write("=" * 70 + "\n")
        f.write("Correlation strength (|r|):\n")
        f.write("  0.00 - 0.19: Very weak\n")
        f.write("  0.20 - 0.39: Weak\n")
        f.write("  0.40 - 0.59: Moderate\n")
        f.write("  0.60 - 0.79: Strong\n")
        f.write("  0.80 - 1.00: Very strong\n\n")
        f.write("Significance:\n")
        f.write("  *** p < 0.001 (highly significant)\n")
        f.write("  **  p < 0.01  (very significant)\n")
        f.write("  *   p < 0.05  (significant)\n")
        f.write("  ns  p ≥ 0.05  (not significant)\n\n")
        f.write("R² = Proportion of variance explained (0 to 1)\n")
        
        f.write("\n" + "=" * 70 + "\n")
        f.write("DETAILED RESULTS\n")
        f.write("=" * 70 + "\n\n")
        
        for _, row in results_df.iterrows():
            f.write(f"{row['Measurement']}\n")
            f.write(f"  Sample size: {row['n']}\n")
            f.write(f"  Pearson's r: {row['r']:.4f}\n")
            f.write(f"  P-value: {row['p_value']:.6f}\n")
            f.write(f"  R-squared: {row['r_squared']:.4f}\n")
            f.write(f"  Interpretation: {row['interpretation']} correlation\n")
            f.write(f"  Mean difference: {row['mean_diff']:.4f}\n")
            f.write(f"  SD of differences: {row['sd_diff']:.4f}\n")
            f.write(f"  95% Limits of agreement: {row['mean_diff'] - 1.96*row['sd_diff']:.4f} to {row['mean_diff'] + 1.96*row['sd_diff']:.4f}\n")
            f.write("\n")
    
    print(f"✓ Detailed report saved: {report_file}")
    
    # Print summary statistics
    print("\n" + "=" * 70)
    print("ANALYSIS COMPLETE")
    print("=" * 70)
    print(f"\nTotal measurements analyzed: {len(results)}")
    print(f"Significant correlations (p < 0.05): {sum(results_df['p_value'] < 0.05)}")
    print(f"Strong correlations (|r| ≥ 0.60): {sum(abs(results_df['r']) >= 0.60)}")
    print(f"\nMean correlation coefficient: {results_df['r'].mean():.4f}")
    print(f"Range: {results_df['r'].min():.4f} to {results_df['r'].max():.4f}")
    
    print(f"\nAll results saved to: {output_dir}/")
    print(f"  • Scatter plots: *_scatter.png")
    print(f"  • Bland-Altman plots: *_bland_altman.png")
    print(f"  • Summary table: correlation_summary.csv")
    print(f"  • Detailed report: correlation_report.txt")
    
    return results_df

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\nUsage: python correlation_analysis.py <Final_Fa_Vals_file.xlsx> [output_directory]")
        print("\nExample:")
        print("  python correlation_analysis.py Final_Fa_Vals_Complete.xlsx")
        print("  python correlation_analysis.py Final_Fa_Vals_Complete.xlsx my_results")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else 'correlation_results'
    
    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)
    
    try:
        results = run_analysis(input_file, output_dir)
        print("\n✓ Analysis completed successfully!")
    except Exception as e:
        print(f"\n✗ Error during analysis: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
