"""
Main script to process residential energy certificates
"""

import sys
from pathlib import Path
from src.matrix_generator import MatrixGenerator


def main():
    if len(sys.argv) < 2:
        print("Usage: python -m src.main <input_folder> [output_excel]")
        print("\nExample:")
        print("  python -m src.main data/sample_1")
        print("  python -m src.main 'data/DE PAZ FRANCO QUINTILIANA'")
        sys.exit(1)
    
    input_folder = sys.argv[1]
    
    # Get project name from folder
    project_name = Path(input_folder).name
    
    # Default output path
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = f"data/output/{project_name}_Checks.xlsx"
    
    print("="*80)
    print("🏠 RESIDENTIAL ENERGY CERTIFICATES PARSER")
    print("="*80)
    print(f"Input:  {input_folder}")
    print(f"Output: {output_path}")
    print("="*80)
    
    # Create matrix generator
    generator = MatrixGenerator(input_folder)
    
    # Parse all documents
    generator.parse_all_documents()
    
    # Generate Excel
    generator.generate_excel(output_path, project_name)
    
    print("\n✅ DONE!")


if __name__ == "__main__":
    main()
