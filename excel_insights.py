
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime

class ExcelInsights:
    """
    A class to analyze Excel files and generate insights.
    """
    
    def __init__(self, file_path=None):
        """
        Initialize the ExcelInsights class.
        
        Args:
            file_path (str, optional): Path to the Excel file to analyze.
        """
        self.file_path = file_path
        self.data = None
        self.sheets = {}
        self.insights = {}
        
    def load_excel(self, file_path=None):
        """
        Load an Excel file into pandas DataFrames.
        
        Args:
            file_path (str, optional): Path to the Excel file to load.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        if file_path:
            self.file_path = file_path
            
        if not self.file_path:
            print("Error: No file path provided.")
            return False
            
        try:
            # Load all sheets from the Excel file
            excel_file = pd.ExcelFile(self.file_path)
            sheet_names = excel_file.sheet_names
            
            # Store each sheet as a DataFrame in the sheets dictionary
            for sheet in sheet_names:
                self.sheets[sheet] = pd.read_excel(self.file_path, sheet_name=sheet)
                
            # Set the first sheet as the default data
            if sheet_names:
                self.data = self.sheets[sheet_names[0]]
                
            print(f"Successfully loaded {len(sheet_names)} sheets from {os.path.basename(self.file_path)}")
            return True
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def get_basic_info(self):
        """
        Get basic information about the loaded data.
        
        Returns:
            dict: Dictionary containing basic information about the data.
        """
        if self.data is None:
            print("No data loaded. Please load an Excel file first.")
            return {}
            
        info = {
            "rows": len(self.data),
            "columns": len(self.data.columns),
            "column_names": list(self.data.columns),
            "data_types": {col: str(dtype) for col, dtype in self.data.dtypes.items()},
            "missing_values": self.data.isnull().sum().to_dict(),
            "sheets": list(self.sheets.keys())
        }
        
        self.insights["basic_info"] = info
        return info
    
    def generate_summary_statistics(self, sheet_name=None):
        """
        Generate summary statistics for numerical columns.
        
        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.
            
        Returns:
            dict: Dictionary containing summary statistics.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        # Get numerical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        
        if not numerical_cols:
            print("No numerical columns found in the data.")
            return {}
        
        # Calculate summary statistics
        summary = {
            "numerical_columns": numerical_cols,
            "statistics": data[numerical_cols].describe().to_dict()
        }
        
        # Add to insights
        if "summary_statistics" not in self.insights:
            self.insights["summary_statistics"] = {}
            
        if sheet_name:
            self.insights["summary_statistics"][sheet_name] = summary
        else:
            self.insights["summary_statistics"]["default"] = summary
            
        return summary
    
    def find_correlations(self, sheet_name=None, threshold=0.5):
        """
        Find correlations between numerical columns.
        
        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.
            threshold (float, optional): Correlation threshold to report.
            
        Returns:
            dict: Dictionary containing correlations above the threshold.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        # Get numerical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        
        if len(numerical_cols) < 2:
            print("Need at least 2 numerical columns to calculate correlations.")
            return {}
        
        # Calculate correlations
        corr_matrix = data[numerical_cols].corr()
        
        # Find correlations above threshold
        high_correlations = {}
        for i in range(len(numerical_cols)):
            for j in range(i+1, len(numerical_cols)):
                col1 = numerical_cols[i]
                col2 = numerical_cols[j]
                correlation = corr_matrix.loc[col1, col2]
                
                if abs(correlation) >= threshold:
                    high_correlations[f"{col1} - {col2}"] = correlation
        
        # Add to insights
        if "correlations" not in self.insights:
            self.insights["correlations"] = {}
            
        if sheet_name:
            self.insights["correlations"][sheet_name] = high_correlations
        else:
            self.insights["correlations"]["default"] = high_correlations
            
        return high_correlations
    
    def identify_outliers(self, sheet_name=None, method="iqr", threshold=1.5):
        """
        Identify outliers in numerical columns.
        
        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.
            method (str, optional): Method to use for outlier detection ('iqr' or 'zscore').
            threshold (float, optional): Threshold for outlier detection.
            
        Returns:
            dict: Dictionary containing outliers for each numerical column.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        # Get numerical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        
        if not numerical_cols:
            print("No numerical columns found in the data.")
            return {}
        
        outliers = {}
        
        for col in numerical_cols:
            col_data = data[col].dropna()
            
            if method == "iqr":
                # IQR method
                Q1 = col_data.quantile(0.25)
                Q3 = col_data.quantile(0.75)
                IQR = Q3 - Q1
                
                lower_bound = Q1 - threshold * IQR
                upper_bound = Q3 + threshold * IQR
                
                col_outliers = col_data[(col_data < lower_bound) | (col_data > upper_bound)]
                
            elif method == "zscore":
                # Z-score method
                mean = col_data.mean()
                std = col_data.std()
                
                if std == 0:  # Skip columns with zero standard deviation
                    continue
                    
                z_scores = abs((col_data - mean) / std)
                col_outliers = col_data[z_scores > threshold]
                
            else:
                print(f"Unknown method: {method}. Using IQR method instead.")
                Q1 = col_data.quantile(0.25)
                Q3 = col_data.quantile(0.75)
                IQR = Q3 - Q1
                
                lower_bound = Q1 - threshold * IQR
                upper_bound = Q3 + threshold * IQR
                
                col_outliers = col_data[(col_data < lower_bound) | (col_data > upper_bound)]
            
            if not col_outliers.empty:
                outliers[col] = {
                    "count": len(col_outliers),
                    "percentage": (len(col_outliers) / len(col_data)) * 100,
                    "values": col_outliers.tolist() if len(col_outliers) <= 10 else col_outliers.tolist()[:10]  # Limit to 10 values
                }
        
        # Add to insights
        if "outliers" not in self.insights:
            self.insights["outliers"] = {}
            
        if sheet_name:
            self.insights["outliers"][sheet_name] = outliers
        else:
            self.insights["outliers"]["default"] = outliers
            
        return outliers
    
    def analyze_categorical_data(self, sheet_name=None, top_n=5):
        """
        Analyze categorical columns in the data.
        
        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.
            top_n (int, optional): Number of top categories to include in the analysis.
            
        Returns:
            dict: Dictionary containing analysis of categorical columns.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        # Get categorical columns (object, string, or category dtype)
        categorical_cols = data.select_dtypes(include=["object", "string", "category"]).columns.tolist()
        
        if not categorical_cols:
            print("No categorical columns found in the data.")
            return {}
        
        categorical_analysis = {}
        
        for col in categorical_cols:
            # Count value frequencies
            value_counts = data[col].value_counts()
            
            # Get top N categories
            top_categories = value_counts.head(top_n).to_dict()
            
            # Calculate percentage of total for each category
            total_count = len(data[col].dropna())
            top_categories_pct = {k: (v / total_count) * 100 for k, v in top_categories.items()}
            
            # Count unique values and missing values
            unique_count = data[col].nunique()
            missing_count = data[col].isnull().sum()
            
            categorical_analysis[col] = {
                "unique_values": unique_count,
                "missing_values": missing_count,
                "missing_percentage": (missing_count / len(data)) * 100,
                "top_categories": top_categories,
                "top_categories_pct": top_categories_pct
            }
        
        # Add to insights
        if "categorical_analysis" not in self.insights:
            self.insights["categorical_analysis"] = {}
            
        if sheet_name:
            self.insights["categorical_analysis"][sheet_name] = categorical_analysis
        else:
            self.insights["categorical_analysis"]["default"] = categorical_analysis
            
        return categorical_analysis
    
    def analyze_date_columns(self, sheet_name=None):
        """
        Analyze date columns in the data.
        
        Args:
            sheet_name (str, optional): Name of the sheet to analyze. If None, uses the default data.
            
        Returns:
            dict: Dictionary containing analysis of date columns.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return {}
        
        date_analysis = {}
        
        # Try to identify date columns
        for col in data.columns:
            # Check if column is already a datetime type
            if pd.api.types.is_datetime64_any_dtype(data[col]):
                date_col = data[col]
            else:
                # Try to convert to datetime
                try:
                    date_col = pd.to_datetime(data[col], errors='coerce')
                    # If more than 70% of values could be converted to dates, consider it a date column
                    if date_col.notnull().sum() / len(date_col) < 0.7:
                        continue
                except:
                    continue
            
            # Analyze the date column
            date_analysis[col] = {
                "min_date": date_col.min().strftime('%Y-%m-%d') if not pd.isna(date_col.min()) else None,
                "max_date": date_col.max().strftime('%Y-%m-%d') if not pd.isna(date_col.max()) else None,
                "range_days": (date_col.max() - date_col.min()).days if not pd.isna(date_col.min()) and not pd.isna(date_col.max()) else None,
                "missing_values": date_col.isnull().sum(),
                "missing_percentage": (date_col.isnull().sum() / len(date_col)) * 100
            }
            
            # Add day of week distribution if there are enough dates
            if date_col.notnull().sum() > 10:
                day_of_week = date_col.dt.day_name().value_counts().to_dict()
                date_analysis[col]["day_of_week_distribution"] = day_of_week
                
                # Add month distribution
                month_dist = date_col.dt.month_name().value_counts().to_dict()
                date_analysis[col]["month_distribution"] = month_dist
                
                # Add year distribution
                year_dist = date_col.dt.year.value_counts().to_dict()
                date_analysis[col]["year_distribution"] = year_dist
        
        # Add to insights
        if "date_analysis" not in self.insights:
            self.insights["date_analysis"] = {}
            
        if sheet_name:
            self.insights["date_analysis"][sheet_name] = date_analysis
        else:
            self.insights["date_analysis"]["default"] = date_analysis
            
        return date_analysis
    
    def generate_insights_report(self, output_file=None):
        """
        Generate a comprehensive insights report.
        
        Args:
            output_file (str, optional): Path to save the report. If None, returns the report as a string.
            
        Returns:
            str: Report as a string if output_file is None, otherwise None.
        """
        if not self.insights:
            print("No insights generated. Please analyze the data first.")
            return None
        
        report = []
        report.append("# Excel Insights Report")
        report.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # Add file information
        if self.file_path:
            report.append(f"## File Information")
            report.append(f"- File: {os.path.basename(self.file_path)}")
            report.append(f"- Path: {self.file_path}")
            if "basic_info" in self.insights:
                info = self.insights["basic_info"]
                report.append(f"- Sheets: {', '.join(info['sheets'])}")
                report.append(f"- Rows: {info['rows']}")
                report.append(f"- Columns: {info['columns']}")
                report.append("")
        
        # Add summary statistics
        if "summary_statistics" in self.insights:
            report.append("## Summary Statistics")
            for sheet, stats in self.insights["summary_statistics"].items():
                if sheet != "default":
                    report.append(f"\n### Sheet: {sheet}")
                
                if not stats.get("numerical_columns"):
                    report.append("No numerical columns found for analysis.")
                    continue
                
                report.append("The following numerical columns were analyzed:")
                report.append(f"- {', '.join(stats['numerical_columns'])}\n")
                
                for col, col_stats in stats["statistics"].items():
                    report.append(f"### {col}")
                    report.append(f"- Count: {col_stats.get('count', 'N/A')}")
                    report.append(f"- Mean: {col_stats.get('mean', 'N/A'):.2f}")
                    report.append(f"- Std Dev: {col_stats.get('std', 'N/A'):.2f}")
                    report.append(f"- Min: {col_stats.get('min', 'N/A'):.2f}")
                    report.append(f"- 25%: {col_stats.get('25%', 'N/A'):.2f}")
                    report.append(f"- Median: {col_stats.get('50%', 'N/A'):.2f}")
                    report.append(f"- 75%: {col_stats.get('75%', 'N/A'):.2f}")
                    report.append(f"- Max: {col_stats.get('max', 'N/A'):.2f}")
                    report.append("")
        
        # Add correlations
        if "correlations" in self.insights:
            report.append("## Correlations")
            for sheet, corrs in self.insights["correlations"].items():
                if sheet != "default":
                    report.append(f"\n### Sheet: {sheet}")
                
                if not corrs:
                    report.append("No significant correlations found.")
                    continue
                
                report.append("The following pairs of columns show significant correlation:")
                for pair, corr in corrs.items():
                    report.append(f"- {pair}: {corr:.2f}")
                report.append("")
        
        # Add outliers
        if "outliers" in self.insights:
            report.append("## Outliers")
            for sheet, outs in self.insights["outliers"].items():
                if sheet != "default":
                    report.append(f"\n### Sheet: {sheet}")
                
                if not outs:
                    report.append("No outliers detected.")
                    continue
                
                for col, out_info in outs.items():
                    report.append(f"### {col}")
                    report.append(f"- Outlier count: {out_info['count']}")
                    report.append(f"- Percentage of data: {out_info['percentage']:.2f}%")
                    if out_info['values']:
                        report.append(f"- Sample outliers: {out_info['values']}")
                    report.append("")
        
        # Add categorical analysis
        if "categorical_analysis" in self.insights:
            report.append("## Categorical Data Analysis")
            for sheet, cat_analysis in self.insights["categorical_analysis"].items():
                if sheet != "default":
                    report.append(f"\n### Sheet: {sheet}")
                
                if not cat_analysis:
                    report.append("No categorical columns found for analysis.")
                    continue
                
                for col, col_analysis in cat_analysis.items():
                    report.append(f"### {col}")
                    report.append(f"- Unique values: {col_analysis['unique_values']}")
                    report.append(f"- Missing values: {col_analysis['missing_values']} ({col_analysis['missing_percentage']:.2f}%)")
                    
                    report.append("\nTop categories:")
                    for category, count in col_analysis['top_categories'].items():
                        category_str = str(category) if category is not None else "NULL"
                        pct = col_analysis['top_categories_pct'].get(category, 0)
                        report.append(f"- {category_str}: {count} ({pct:.2f}%)")
                    report.append("")
        
        # Add date analysis
        if "date_analysis" in self.insights:
            report.append("## Date Analysis")
            for sheet, date_analysis in self.insights["date_analysis"].items():
                if sheet != "default":
                    report.append(f"\n### Sheet: {sheet}")
                
                if not date_analysis:
                    report.append("No date columns found for analysis.")
                    continue
                
                for col, col_analysis in date_analysis.items():
                    report.append(f"### {col}")
                    report.append(f"- Date range: {col_analysis['min_date']} to {col_analysis['max_date']}")
                    report.append(f"- Range in days: {col_analysis['range_days']}")
                    report.append(f"- Missing values: {col_analysis['missing_values']} ({col_analysis['missing_percentage']:.2f}%)")
                    
                    if "day_of_week_distribution" in col_analysis:
                        report.append("\nDay of week distribution:")
                        for day, count in col_analysis['day_of_week_distribution'].items():
                            report.append(f"- {day}: {count}")
                    
                    if "month_distribution" in col_analysis:
                        report.append("\nMonth distribution:")
                        for month, count in col_analysis['month_distribution'].items():
                            report.append(f"- {month}: {count}")
                    
                    if "year_distribution" in col_analysis:
                        report.append("\nYear distribution:")
                        for year, count in col_analysis['year_distribution'].items():
                            report.append(f"- {year}: {count}")
                    report.append("")
        
        # Compile the report
        report_text = "\n".join(report)
        
        # Save to file if specified
        if output_file:
            try:
                with open(output_file, 'w') as f:
                    f.write(report_text)
                print(f"Report saved to {output_file}")
                return None
            except Exception as e:
                print(f"Error saving report: {e}")
                return report_text
        
        return report_text
    
    def visualize_data(self, output_dir=None, sheet_name=None):
        """
        Generate visualizations for the data.
        
        Args:
            output_dir (str, optional): Directory to save visualizations. If None, displays plots.
            sheet_name (str, optional): Name of the sheet to visualize. If None, uses the default data.
            
        Returns:
            list: List of paths to saved visualizations if output_dir is provided, otherwise None.
        """
        if sheet_name and sheet_name in self.sheets:
            data = self.sheets[sheet_name]
        elif self.data is not None:
            data = self.data
        else:
            print("No data loaded. Please load an Excel file first.")
            return []
        
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        saved_files = []
        
        # Get numerical and categorical columns
        numerical_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        categorical_cols = data.select_dtypes(include=["object", "string", "category"]).columns.tolist()
        
        # 1. Histograms for numerical columns
        for col in numerical_cols[:5]:  # Limit to first 5 columns to avoid too many plots
            plt.figure(figsize=(10, 6))
            plt.hist(data[col].dropna(), bins=20, alpha=0.7, color='skyblue', edgecolor='black')
            plt.title(f'Distribution of {col}')
            plt.xlabel(col)
            plt.ylabel('Frequency')
            plt.grid(True, alpha=0.3)
            
            if output_dir:
                file_path = os.path.join(output_dir, f"histogram_{col}.png")
                plt.savefig(file_path)
                saved_files.append(file_path)
                plt.close()
            else:
                plt.show()
        
        # 2. Bar charts for categorical columns
        for col in categorical_cols[:5]:  # Limit to first 5 columns
            # Get top 10 categories
            value_counts = data[col].value_counts().head(10)
            
            plt.figure(figsize=(12, 6))
            bars = plt.bar(value_counts.index.astype(str), value_counts.values, color='lightgreen', edgecolor='black')
            plt.title(f'Top 10 Categories in {col}')
            plt.xlabel(col)
            plt.ylabel('Count')
            plt.xticks(rotation=45, ha='right')
            plt.grid(True, alpha=0.3, axis='y')
            
            # Add count labels on top of bars
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{height}', ha='center', va='bottom')
            
            plt.tight_layout()
            
            if output_dir:
                file_path = os.path.join(output_dir, f"barchart_{col}.png")
                plt.savefig(file_path)
                saved_files.append(file_path)
                plt.close()
            else:
                plt.show()
        
        # 3. Correlation heatmap for numerical columns
        if len(numerical_cols) > 1:
            plt.figure(figsize=(12, 10))
            corr_matrix = data[numerical_cols].corr()
            plt.imshow(corr_matrix, cmap='coolwarm', interpolation='none', aspect='auto')
            plt.colorbar(label='Correlation Coefficient')
            plt.title('Correlation Heatmap')
            
            # Add correlation values
            for i in range(len(numerical_cols)):
                for j in range(len(numerical_cols)):
                    plt.text(j, i, f'{corr_matrix.iloc[i, j]:.2f}', 
                             ha='center', va='center', 
                             color='white' if abs(corr_matrix.iloc[i, j]) > 0.5 else 'black')
            
            plt.xticks(range(len(numerical_cols)), numerical_cols, rotation=45, ha='right')
            plt.yticks(range(len(numerical_cols)), numerical_cols)
            plt.tight_layout()
            
            if output_dir:
                file_path = os.path.join(output_dir, "correlation_heatmap.png")
                plt.savefig(file_path)
                saved_files.append(file_path)
                plt.close()
            else:
                plt.show()
        
        # 4. Box plots for numerical columns to visualize outliers
        if numerical_cols:
            plt.figure(figsize=(14, 8))
            data[numerical_cols[:10]].boxplot()  # Limit to first 10 columns
            plt.title('Box Plots for Numerical Columns')
            plt.xticks(rotation=45, ha='right')
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            
            if output_dir:
                file_path = os.path.join(output_dir, "boxplots.png")
                plt.savefig(file_path)
                saved_files.append(file_path)
                plt.close()
            else:
                plt.show()
        
        # 5. Pie charts for categorical columns with few unique values
        for col in categorical_cols[:3]:  # Limit to first 3 columns
            value_counts = data[col].value_counts()
            
            # Only create pie chart if there are 10 or fewer unique values
            if len(value_counts) <= 10:
                plt.figure(figsize=(10, 8))
                plt.pie(value_counts.values, labels=value_counts.index.astype(str), 
                        autopct='%1.1f%%', startangle=90, shadow=True)
                plt.title(f'Distribution of {col}')
                plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
                
                if output_dir:
                    file_path = os.path.join(output_dir, f"piechart_{col}.png")
                    plt.savefig(file_path)
                    saved_files.append(file_path)
                    plt.close()
                else:
                    plt.show()
        
        return saved_files

def main():
    """
    Main function to demonstrate the Excel Insights Engine.
    """
    print("Excel Insights Engine")
    print("--------------------")
    
    # Check if a file path was provided as an argument
    import sys
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Use a sample file for demonstration
        file_path = "samples/sample_data.xlsx"
        print(f"No file specified. Using sample file: {file_path}")
    
    # Create an instance of ExcelInsights
    insights = ExcelInsights(file_path)
    
    # Try to load the Excel file
    if not insights.load_excel():
        print("Failed to load Excel file. Exiting.")
        return
    
    # Get basic information
    print("\nGetting basic information...")
    info = insights.get_basic_info()
    print(f"Rows: {info.get('rows', 'N/A')}")
    print(f"Columns: {info.get('columns', 'N/A')}")
    print(f"Column names: {', '.join(info.get('column_names', []))}")
    
    # Generate summary statistics
    print("\nGenerating summary statistics...")
    insights.generate_summary_statistics()
    
    # Find correlations
    print("\nFinding correlations...")
    correlations = insights.find_correlations(threshold=0.7)
    if correlations:
        print("Found correlations:")
        for pair, corr in correlations.items():
            print(f"  {pair}: {corr:.2f}")
    else:
        print("No significant correlations found.")
    
    # Identify outliers
    print("\nIdentifying outliers...")
    outliers = insights.identify_outliers()
    if outliers:
        print("Found outliers in the following columns:")
        for col, out_info in outliers.items():
            print(f"  {col}: {out_info['count']} outliers ({out_info['percentage']:.2f}%)")
    else:
        print("No outliers found.")
    
    # Analyze categorical data
    print("\nAnalyzing categorical data...")
    insights.analyze_categorical_data()
    
    # Analyze date columns
    print("\nAnalyzing date columns...")
    insights.analyze_date_columns()
    
    # Generate visualizations
    print("\nGenerating visualizations...")
    vis_dir = "visualizations"
    saved_files = insights.visualize_data(output_dir=vis_dir)
    if saved_files:
        print(f"Saved {len(saved_files)} visualizations to {vis_dir} directory.")
    
    # Generate insights report
    print("\nGenerating insights report...")
    report_file = "insights_report.md"
    insights.generate_insights_report(output_file=report_file)
    print(f"Report saved to {report_file}")
    
    print("\nAnalysis complete!")

if __name__ == "__main__":
    main()
