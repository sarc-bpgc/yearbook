import pandas as pd
import re

def main():
    try:
        df = pd.read_excel('data.xlsx')
        
        # Determine the quote and name columns based on what's available
        name_col = 'Name'
        quote_col = next((c for c in df.columns if 'quote' in c.lower()), None)
        
        if not quote_col or name_col not in df.columns:
            print("Could not find Name or Quote columns.")
            return

        print(f"Searching for Unicode (non-ASCII) characters in quotes...")
        print("-" * 50)
        
        # Function to find non-ASCII characters
        def get_unicode_chars(text):
            if pd.isna(text): return []
            return list(set(re.findall(r'[^\x00-\x7F]', str(text))))

        found_unicode = False
        for idx, row in df.iterrows():
            name = row[name_col]
            quote = row[quote_col]
            
            unicode_chars = get_unicode_chars(quote)
            
            if unicode_chars:
                found_unicode = True
                print(f"Student: {name}")
                print(f"Quote:   {quote}")
                print(f"Unicode Characters Found: {' '.join(unicode_chars)}")
                print("-" * 50)
                
        if not found_unicode:
            print("No Unicode characters found in the dataset.")
            
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == '__main__':
    main()
