import pandas as pd
import plotly.express as px
import fileinput

# Load Excel file
file_path = "/Users/ghazaleh/Documents/ECE88/ECE88-Reunion04.xlsx"
df = pd.read_excel(file_path)

# Clean and normalize Gender
df['Gender'] = df['Gender'].astype(str).str.strip().str.title()
df['Gender'] = df['Gender'].replace({
    'F': 'Female', 'M': 'Male',
    'Woman': 'Female', 'Man': 'Male',
    '': 'Unknown', 'Nan': 'Unknown', 'None': 'Unknown', 'Na': 'Unknown'
})
df['Gender'] = df['Gender'].where(df['Gender'].isin(['Female', 'Male']), 'Unknown')

# Clean and normalize Marital Status
df['Marital Status :-)'] = df['Marital Status :-)'].astype(str).str.strip().str.title()
df['Marital Status :-)'] = df['Marital Status :-)'].replace({
    '': 'Unknown', 'Nan': 'Unknown', 'None': 'Unknown', 'Na': 'Unknown'
})
df['Marital Status :-)'] = df['Marital Status :-)'].fillna('Unknown')

# Add emojis to marital status
marital_emoji_map = {
    'Married': '💍 Married',
    'Single': '💫 Single',
    'Unknown': '❓ Unknown'
}
df['Marital Status :-)'] = df['Marital Status :-)'].apply(lambda x: marital_emoji_map.get(x, '❓ Unknown'))

# Clean Degree
df['Degree'] = df['Degree'].astype(str).str.strip()
df['Degree'] = df['Degree'].replace({
    '': 'Unknown', 'Nan': 'Unknown', 'None': 'Unknown', 'Na': 'Unknown'
})
df['Degree'] = df['Degree'].fillna('Unknown')

# Rename column for clarity
df = df.rename(columns={'Marital Status :-)': 'Marital Status'})

# Custom color mapping for Gender
color_map = {
    'Female': '#F8766D',   # Coral/Pink
    'Male': '#00BFC4',     # Teal/Cyan
    'Unknown': '#C0C0C0'   # Light gray
}

# Create Sunburst chart: Gender → Degree → Marital Status
fig = px.sunburst(
    df,
    path=['Gender', 'Degree', 'Marital Status'],
    title="🎉 ECE88 Reunion: Gender → Degree → Marital Status",
    color='Gender',
    color_discrete_map=color_map
)

# Improve visibility
fig.update_traces(
    textinfo='label+percent entry',
    insidetextorientation='radial',
    textfont_size=14
)

# Save HTML
output_path = "/Users/ghazaleh/Documents/ECE88/ece88_sunburst.html"
fig.write_html(output_path)

# Inject note at the top of the HTML body
note_html = """
<div style="background:#f9f9f9; border-left: 6px solid #00BFC4; padding: 12px; margin: 20px; font-family: sans-serif;">
  <h3>🧠 How to use this chart</h3>
  <ul>
    <li>Click on any colored section (e.g., <strong>Female</strong>, <strong>PhD</strong>) to zoom in</li>
    <li>Click the center circle to zoom out</li>
    <li>Hover over segments to see detailed info and percentages</li>
  </ul>
</div>
"""

# Inject right after <body>
for line in fileinput.input(output_path, inplace=True):
    if "<body>" in line:
        print(line, end='')
        print(note_html)  # inject note right after <body>
    else:
        print(line, end='')

print(f"✅ Chart saved with instruction note: {output_path}")
