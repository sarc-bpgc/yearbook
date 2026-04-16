import pandas as pd
COL_PHOTO  = 'Upload a clear, well-lit, decent photo (1:1 ratio or passport size). Editing is not allowed, and you can only upload once. Ensure View Permissions are set to "Anyone with the Link"'
df = pd.read_excel('data.xlsx')
print("Unique photo values:", df[COL_PHOTO].unique())
