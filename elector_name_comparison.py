#!/usr/bin/env python3
"""
Elector Name Duplicate Finder
Compares elector names between two Excel sheets (2025_LIST and 2002_LIST) 
using fuzzy matching to handle Telugu/English transliteration differences.

Author: Auto-generated
Date: 2025
"""

import os
import sys
import pandas as pd
import logging
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from datetime import datetime

try:
    from rapidfuzz import fuzz  # type: ignore
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    try:
        from fuzzywuzzy import fuzz  # type: ignore
        RAPIDFUZZ_AVAILABLE = False
    except ImportError:
        print("ERROR: Neither 'rapidfuzz' nor 'fuzzywuzzy' is installed.")
        print("Please install one of them:")
        print("  pip install rapidfuzz  (recommended)")
        print("  OR")
        print("  pip install fuzzywuzzy python-Levenshtein")
        sys.exit(1)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('elector_comparison.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ElectorNameComparator:
    """
    Compares elector names between two Excel sheets using fuzzy matching
    to handle Telugu/English transliteration differences.
    """
    
    def __init__(self, excel_path: str, similarity_threshold: int = 85):
        """
        Initialize the comparator.
        
        Args:
            excel_path: Path to the Excel file
            similarity_threshold: Minimum similarity score (0-100) to consider names as duplicates
        """
        self.excel_path = excel_path
        self.similarity_threshold = similarity_threshold
        self.df_2025 = None
        self.df_2002 = None
        self.duplicates = []
        self.primary_key_column = None  # Store detected primary key column name
        self.stats = {
            'total_2025': 0,
            'total_2002': 0,
            'exact_matches': 0,
            'fuzzy_matches': 0,
            'no_matches': 0
        }
        
    def load_excel_sheets(self) -> bool:
        """
        Load the two sheets from the Excel file.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Loading Excel file: {self.excel_path}")
            
            # Read both sheets
            excel_file = pd.ExcelFile(self.excel_path)
            sheet_names = excel_file.sheet_names
            logger.info(f"Available sheets: {sheet_names}")
            
            # Check if required sheets exist
            if '2025_LIST' not in sheet_names:
                logger.error("Sheet '2025_LIST' not found in Excel file")
                return False
            if '2002_LIST' not in sheet_names:
                logger.error("Sheet '2002_LIST' not found in Excel file")
                return False
            
            # Load sheets
            self.df_2025 = pd.read_excel(self.excel_path, sheet_name='2025_LIST')
            self.df_2002 = pd.read_excel(self.excel_path, sheet_name='2002_LIST')
            
            logger.info(f"Loaded 2025_LIST: {len(self.df_2025)} rows")
            logger.info(f"Loaded 2002_LIST: {len(self.df_2002)} rows")
            
            # Validate required columns
            required_cols = ["Elector's Name", "Elector's Name(Vernacular)"]
            
            for col in required_cols:
                if col not in self.df_2025.columns:
                    logger.error(f"Column '{col}' not found in 2025_LIST sheet")
                    return False
                if col not in self.df_2002.columns:
                    logger.error(f"Column '{col}' not found in 2002_LIST sheet")
                    return False
            
            # Clean data - remove rows with all NaN values
            self.df_2025 = self.df_2025.dropna(subset=required_cols, how='all')
            self.df_2002 = self.df_2002.dropna(subset=required_cols, how='all')
            
            logger.info(f"After cleaning - 2025_LIST: {len(self.df_2025)} rows")
            logger.info(f"After cleaning - 2002_LIST: {len(self.df_2002)} rows")
            
            # Detect primary key column
            self.primary_key_column = self._detect_primary_key_column()
            if self.primary_key_column:
                logger.info(f"Detected primary key column: '{self.primary_key_column}'")
            else:
                logger.warning("No primary key column detected. Will use sequential numbering for duplicate_id.")
            
            self.stats['total_2025'] = len(self.df_2025)
            self.stats['total_2002'] = len(self.df_2002)
            
            return True
            
        except FileNotFoundError:
            logger.error(f"Excel file not found: {self.excel_path}")
            return False
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            return False
    
    def _detect_primary_key_column(self) -> Optional[str]:
        """
        Detect primary key column from common naming patterns.
        
        Returns:
            Primary key column name if found, None otherwise
        """
        # Common primary key column names (case-insensitive)
        primary_key_patterns = [
            'id', 'ID', 'Id',
            'serial number', 'Serial Number', 'SERIAL NUMBER',
            's.no', 'S.No', 'S.NO', 'S.No.',
            'sno', 'SNO', 'SNo',
            'sl no', 'SL No', 'SL NO',
            'slno', 'SLNO',
            'elector id', 'Elector ID', 'ELECTOR ID',
            'voter id', 'Voter ID', 'VOTER ID',
            'epic no', 'EPIC No', 'EPIC NO',
            'epic', 'EPIC'
        ]
        
        # Check columns in 2025_LIST (both sheets should have same structure)
        for col in self.df_2025.columns:
            col_lower = str(col).strip().lower()
            # Check if column name matches any pattern
            for pattern in primary_key_patterns:
                if col_lower == pattern.lower() or col_lower.endswith(pattern.lower()):
                    # Verify column exists in both sheets and has unique values
                    if col in self.df_2002.columns:
                        # Check if column has mostly non-null values
                        non_null_2025 = self.df_2025[col].notna().sum()
                        non_null_2002 = self.df_2002[col].notna().sum()
                        if non_null_2025 > len(self.df_2025) * 0.8:  # At least 80% non-null
                            return col
        
        return None
    
    def normalize_name(self, name: str) -> str:
        """
        Normalize a name for comparison.
        
        Args:
            name: Name string to normalize
            
        Returns:
            Normalized name string
        """
        if pd.isna(name) or name is None:
            return ""
        
        # Convert to string and strip whitespace
        name = str(name).strip()
        
        # Convert to lowercase for case-insensitive comparison
        name = name.lower()
        
        # Remove extra spaces
        name = ' '.join(name.split())
        
        return name
    
    def calculate_similarity(self, name1: str, name2: str) -> int:
        """
        Calculate similarity score between two names.
        
        Args:
            name1: First name
            name2: Second name
            
        Returns:
            Similarity score (0-100)
        """
        if not name1 or not name2:
            return 0
        
        # Use rapidfuzz or fuzzywuzzy
        if RAPIDFUZZ_AVAILABLE:
            # rapidfuzz uses token_sort_ratio by default for better matching
            score = fuzz.token_sort_ratio(name1, name2)
        else:
            # fuzzywuzzy
            score = fuzz.token_sort_ratio(name1, name2)
        
        return score
    
    def find_best_match(self, target_name: str, candidate_names: List[str], 
                       candidate_indices: List[int]) -> Tuple[Optional[int], int]:
        """
        Find the best matching name from a list of candidates.
        
        Args:
            target_name: Name to match
            candidate_names: List of candidate names
            candidate_indices: List of indices corresponding to candidates
            
        Returns:
            Tuple of (best_match_index, similarity_score)
        """
        if not target_name or not candidate_names:
            return None, 0
        
        best_score = 0
        best_index = None
        
        for idx, candidate in zip(candidate_indices, candidate_names):
            score = self.calculate_similarity(target_name, candidate)
            if score > best_score:
                best_score = score
                best_index = idx
        
        return best_index, best_score
    
    def compare_names(self) -> List[Dict]:
        """
        Compare names between the two sheets and find duplicates.
        
        Returns:
            List of duplicate matches with details
        """
        logger.info("Starting name comparison...")
        
        duplicates = []
        
        # Prepare name lists from both sheets
        names_2025_english = []
        names_2025_vernacular = []
        names_2002_english = []
        names_2002_vernacular = []
        
        # Extract names from 2025_LIST
        for idx, row in self.df_2025.iterrows():
            eng_name = self.normalize_name(row.get("Elector's Name", ""))
            vern_name = self.normalize_name(row.get("Elector's Name(Vernacular)", ""))
            names_2025_english.append(eng_name)
            names_2025_vernacular.append(vern_name)
        
        # Extract names from 2002_LIST
        for idx, row in self.df_2002.iterrows():
            eng_name = self.normalize_name(row.get("Elector's Name", ""))
            vern_name = self.normalize_name(row.get("Elector's Name(Vernacular)", ""))
            names_2002_english.append(eng_name)
            names_2002_vernacular.append(vern_name)
        
        logger.info("Comparing names using fuzzy matching...")
        
        # Track sequential ID as fallback if no primary key
        sequential_id = 1
        
        # Compare each name from 2025_LIST with all names in 2002_LIST
        for idx_2025, row_2025 in self.df_2025.iterrows():
            eng_2025 = self.normalize_name(row_2025.get("Elector's Name", ""))
            vern_2025 = self.normalize_name(row_2025.get("Elector's Name(Vernacular)", ""))
            
            if not eng_2025 and not vern_2025:
                # Count empty records as no_matches for accurate statistics
                self.stats['no_matches'] += 1
                continue
            
            best_match_idx = None
            best_match_score = 0
            match_type = None
            
            # Try matching English name
            if eng_2025:
                # Match against English names in 2002
                eng_idx, eng_score = self.find_best_match(
                    eng_2025, names_2002_english, list(range(len(names_2002_english)))
                )
                if eng_score > best_match_score:
                    best_match_score = eng_score
                    best_match_idx = eng_idx
                    match_type = "English-English"
                
                # Match against Vernacular names in 2002 (transliteration case)
                vern_idx, vern_score = self.find_best_match(
                    eng_2025, names_2002_vernacular, list(range(len(names_2002_vernacular)))
                )
                if vern_score > best_match_score:
                    best_match_score = vern_score
                    best_match_idx = vern_idx
                    match_type = "English-Vernacular"
            
            # Try matching Vernacular name
            if vern_2025:
                # Match against Vernacular names in 2002
                vern_idx, vern_score = self.find_best_match(
                    vern_2025, names_2002_vernacular, list(range(len(names_2002_vernacular)))
                )
                if vern_score > best_match_score:
                    best_match_score = vern_score
                    best_match_idx = vern_idx
                    match_type = "Vernacular-Vernacular"
                
                # Match against English names in 2002 (transliteration case)
                eng_idx, eng_score = self.find_best_match(
                    vern_2025, names_2002_english, list(range(len(names_2002_english)))
                )
                if eng_score > best_match_score:
                    best_match_score = eng_score
                    best_match_idx = eng_idx
                    match_type = "Vernacular-English"
            
            # If match found above threshold, record it
            if best_match_score >= self.similarity_threshold and best_match_idx is not None:
                # best_match_idx is a positional index, convert to actual DataFrame index
                actual_2002_index = self.df_2002.index[best_match_idx]
                match_row_2002 = self.df_2002.loc[actual_2002_index]
                
                # Use primary key value if available, otherwise use sequential ID
                if self.primary_key_column and self.primary_key_column in self.df_2025.columns:
                    duplicate_id_value = row_2025.get(self.primary_key_column, sequential_id)
                    # Convert to string if not already, handle NaN
                    if pd.isna(duplicate_id_value):
                        duplicate_id_value = sequential_id
                    else:
                        duplicate_id_value = str(duplicate_id_value).strip()
                else:
                    duplicate_id_value = sequential_id
                
                duplicate_info = {
                    'duplicate_id': duplicate_id_value,  # Primary key value or sequential ID
                    '2025_index': idx_2025,  # Already the actual DataFrame index from iterrows()
                    '2025_english': row_2025.get("Elector's Name", ""),
                    '2025_vernacular': row_2025.get("Elector's Name(Vernacular)", ""),
                    '2002_index': actual_2002_index,  # Use actual DataFrame index, not positional
                    '2002_english': match_row_2002.get("Elector's Name", ""),
                    '2002_vernacular': match_row_2002.get("Elector's Name(Vernacular)", ""),
                    'similarity_score': best_match_score,
                    'match_type': match_type,
                    'is_exact_match': best_match_score == 100
                }
                
                duplicates.append(duplicate_info)
                sequential_id += 1  # Increment for next duplicate (fallback)
                
                if best_match_score == 100:
                    self.stats['exact_matches'] += 1
                else:
                    self.stats['fuzzy_matches'] += 1
            else:
                self.stats['no_matches'] += 1
        
        self.duplicates = duplicates
        logger.info(f"Found {len(duplicates)} potential duplicates")
        logger.info(f"Exact matches: {self.stats['exact_matches']}")
        logger.info(f"Fuzzy matches: {self.stats['fuzzy_matches']}")
        logger.info(f"No matches: {self.stats['no_matches']}")
        
        return duplicates
    
    def export_results(self, output_path: Optional[str] = None) -> str:
        """
        Export comparison results to Excel file.
        
        Args:
            output_path: Optional output file path
            
        Returns:
            Path to the output file
        """
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = Path(self.excel_path).stem
            output_path = f"{base_name}_duplicates_{timestamp}.xlsx"
        
        logger.info(f"Exporting results to: {output_path}")
        
        # Create DataFrame from duplicates
        if self.duplicates:
            df_results = pd.DataFrame(self.duplicates)
            
            # Reorder columns for better readability (duplicate_id first for easy reference)
            column_order = [
                'duplicate_id', 'similarity_score', 'match_type', 'is_exact_match',
                '2025_english', '2025_vernacular', '2025_index',
                '2002_english', '2002_vernacular', '2002_index'
            ]
            df_results = df_results[column_order]
            
            # Sort by similarity score (descending)
            df_results = df_results.sort_values('similarity_score', ascending=False)
            
            # Only reassign sequential IDs if no primary key was used (all IDs are numeric sequential)
            # If primary key was used, keep the original primary key values
            if not self.primary_key_column:
                # Reassign sequential duplicate_id after sorting for clean sequential numbering
                df_results['duplicate_id'] = list(range(1, len(df_results) + 1))
        else:
            df_results = pd.DataFrame()
        
        # Create summary sheet
        summary_data = {
            'Metric': [
                'Total records in 2025_LIST',
                'Total records in 2002_LIST',
                'Total duplicates found',
                'Exact matches (100% similarity)',
                'Fuzzy matches (85-99% similarity)',
                'No matches found',
                'Similarity threshold used',
                'Analysis date'
            ],
            'Value': [
                self.stats['total_2025'],
                self.stats['total_2002'],
                len(self.duplicates),
                self.stats['exact_matches'],
                self.stats['fuzzy_matches'],
                self.stats['no_matches'],
                f"{self.similarity_threshold}%",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        
        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            if not df_results.empty:
                df_results.to_excel(writer, sheet_name='Duplicates', index=False)
        
        logger.info(f"Results exported successfully to: {output_path}")
        return output_path
    
    def print_summary(self):
        """Print a summary of the comparison results."""
        print("\n" + "="*80)
        print("ELECTOR NAME COMPARISON SUMMARY")
        print("="*80)
        print(f"Total records in 2025_LIST: {self.stats['total_2025']}")
        print(f"Total records in 2002_LIST: {self.stats['total_2002']}")
        print(f"\nDuplicates found: {len(self.duplicates)}")
        print(f"  - Exact matches (100%): {self.stats['exact_matches']}")
        print(f"  - Fuzzy matches (≥{self.similarity_threshold}%): {self.stats['fuzzy_matches']}")
        print(f"  - No matches: {self.stats['no_matches']}")
        print(f"\nSimilarity threshold: {self.similarity_threshold}%")
        print("="*80 + "\n")


def get_excel_file_path() -> str:
    """
    Prompt user to provide Excel file path.
    
    Returns:
        Path to the Excel file
    """
    print("\n" + "="*80)
    print("ELECTOR NAME DUPLICATE FINDER")
    print("="*80)
    print("\nThis tool compares elector names between two Excel sheets:")
    print("  - 2025_LIST")
    print("  - 2002_LIST")
    print("\nRequired columns in each sheet:")
    print("  - Elector's Name")
    print("  - Elector's Name(Vernacular)")
    print("\n" + "-"*80)
    
    while True:
        file_path = input("\nPlease enter the path to your Excel file (or drag and drop the file here): ").strip()
        
        # Remove quotes if user copied path with quotes
        file_path = file_path.strip('"').strip("'")
        
        if not file_path:
            print("Error: Please provide a file path.")
            continue
        
        if not os.path.exists(file_path):
            print(f"Error: File not found: {file_path}")
            print("Please check the path and try again.")
            continue
        
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            print("Error: Please provide an Excel file (.xlsx or .xls)")
            continue
        
        return file_path


def main():
    """Main execution function."""
    try:
        # Get Excel file path from user
        excel_path = get_excel_file_path()
        
        # Get similarity threshold (optional)
        print("\n" + "-"*80)
        threshold_input = input("Enter similarity threshold (0-100, default 85): ").strip()
        threshold = 85
        if threshold_input:
            try:
                threshold = int(threshold_input)
                if threshold < 0 or threshold > 100:
                    print("Invalid threshold. Using default: 85")
                    threshold = 85
            except ValueError:
                print("Invalid threshold. Using default: 85")
        
        # Initialize comparator
        comparator = ElectorNameComparator(excel_path, similarity_threshold=threshold)
        
        # Load Excel sheets
        if not comparator.load_excel_sheets():
            print("\nError: Failed to load Excel sheets. Please check the file and try again.")
            print("See 'elector_comparison.log' for detailed error information.")
            sys.exit(1)
        
        # Compare names
        print("\nComparing names... This may take a few moments...")
        duplicates = comparator.compare_names()
        
        # Print summary
        comparator.print_summary()
        
        # Export results
        output_file = comparator.export_results()
        print(f"\nResults have been exported to: {output_file}")
        
        # Show sample duplicates
        if duplicates:
            print("\n" + "-"*80)
            print("SAMPLE DUPLICATES (Top 5):")
            print("-"*80)
            for i, dup in enumerate(duplicates[:5], 1):
                print(f"\n{i}. Similarity: {dup['similarity_score']:.1f}% ({dup['match_type']})")
                print(f"   2025 - English: {dup['2025_english']}")
                print(f"          Vernacular: {dup['2025_vernacular']}")
                print(f"   2002 - English: {dup['2002_english']}")
                print(f"          Vernacular: {dup['2002_vernacular']}")
        
        print("\n" + "="*80)
        print("Analysis complete!")
        print("="*80 + "\n")
        
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        logger.exception("Unexpected error occurred")
        print(f"\nError: {str(e)}")
        print("See 'elector_comparison.log' for detailed error information.")
        sys.exit(1)


if __name__ == "__main__":
    main()

