import pandas as pd

def main():
    try:
        import emoji
    except ImportError:
        print("[ERROR] Please run 'pip install emoji' first!")
        return

    try:
        df = pd.read_excel('data.xlsx', engine='openpyxl')
    except Exception as e:
        print(f"[ERROR] Could not read data.xlsx: {e}")
        try:
            with open('data.xlsx', 'rb') as f:
                print(f"[DEBUG] File Signature (first 4 bytes): {f.read(4)}")
        except Exception:
            pass
        return
        
    quote_col = next((c for c in df.columns if 'quote' in c.lower()), None)
    id_col = next((c for c in df.columns if 'BITS ID' in c or 'id' in c.lower()), None)
    name_col = 'Name'

    print("Analyzing Yearbook Data for Emojis...\n" + "="*50)
    
    emoji_students = []
    
    for idx, row in df.iterrows():
        quote = str(row[quote_col]) if pd.notna(row[quote_col]) else ""
        if emoji.emoji_count(quote) > 0:
            student_id = str(row[id_col]).strip()
            name = str(row.get(name_col, "Unknown")).strip()
            
            # Find exact emojis they used
            emojis_used = [e['emoji'] for e in emoji.emoji_list(quote)]
            emojis_str = "".join(emojis_used)
            
            emoji_students.append((name, student_id, emojis_str))
            
    if not emoji_students:
        print("No students found with emojis in their quotes! Everything is up to date.")
        return
        
    print(f"\nFound {len(emoji_students)} students containing emojis!\n")
    
    for name, sid, emojis in emoji_students:
        print(f"Student: {name} ({sid})")
        print(f"Emojis : {emojis}")
        print("-" * 30)

    # Generate the auto-updater bash script
    with open("update_emojis.sh", "w") as f:
        f.write("#!/bin/bash\n\n")
        f.write("# Auto-generated targeted update commands for students with emojis\n")
        for _, sid, _ in emoji_students:
            f.write(f"python script_4x4.py --update {sid}\n")
            
    print(f"\n[DONE] I have created an 'update_emojis.sh' file.")
    print("Just run 'bash update_emojis.sh' to batch update ONLY these specific people!")

if __name__ == '__main__':
    main()
