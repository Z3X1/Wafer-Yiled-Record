"""Example usage of the Wafer Yield Analyzer.

This script demonstrates how to use the wafer_yield_analyzer module
with custom configurations.
"""

from pathlib import Path
import logging
from wafer_yield_analyzer import (
    find_wafer_summary_files,
    create_yield_dataframe,
    create_beautiful_plot,
    save_to_excel_with_chart
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def analyze_wafer_yield_custom(
    source_dir: str,
    output_excel: str = "custom_wafer_report.xlsx",
    output_image: str = "custom_wafer_chart.png"
) -> None:
    """Analyze wafer yield with custom parameters.
    
    Args:
        source_dir: Directory containing Wafer_Summary Excel files.
        output_excel: Name of the output Excel file.
        output_image: Name of the output chart image.
    """
    try:
        logger.info(f"Analyzing wafer data from: {source_dir}")
        
        # Step 1: Find files
        wafer_files = find_wafer_summary_files(source_dir)
        
        if not wafer_files:
            logger.warning(f"No Wafer_Summary files found in {source_dir}")
            return
        
        logger.info(f"Found {len(wafer_files)} files to process")
        
        # Step 2: Extract and organize data
        df = create_yield_dataframe(wafer_files)
        
        # Step 3: Display statistics
        logger.info("\n" + "=" * 60)
        logger.info("Wafer Yield Statistics:")
        logger.info("=" * 60)
        logger.info(f"Total Wafers: {len(df)}")
        logger.info(f"Average Yield: {df['Yield'].mean():.2f}%")
        logger.info(f"Max Yield: {df['Yield'].max():.2f}%")
        logger.info(f"Min Yield: {df['Yield'].min():.2f}%")
        logger.info(f"Std Deviation: {df['Yield'].std():.2f}%")
        logger.info("=" * 60 + "\n")
        
        # Step 4: Create visualization
        create_beautiful_plot(df, output_image)
        
        # Step 5: Save to Excel
        save_to_excel_with_chart(df, output_excel, output_image)
        
        logger.info("✓ Analysis completed successfully!")
        
    except Exception as e:
        logger.error(f"Analysis failed: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    # Example 1: Use default directory
    analyze_wafer_yield_custom(
        source_dir=r"C:\Users\andrel52",
        output_excel="my_wafer_report.xlsx",
        output_image="my_wafer_chart.png"
    )
    
    # Example 2: Use a different directory
    # analyze_wafer_yield_custom(
    #     source_dir=r"D:\WaferData\2024",
    #     output_excel="wafer_2024_report.xlsx",
    #     output_image="wafer_2024_chart.png"
    # )

