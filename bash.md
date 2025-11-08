# Create project folder
mkdir academic-results-dashboard
cd academic-results-dashboard

# Create all files (copy the content from above)
# Save each file with its respective name

# Initialize git
git init
git add .
git commit -m "Initial commit: Academic Results Dashboard Generator"

# Connect to GitHub (replace with your repo URL)
git remote add origin https://github.com/rhmunna143/academic-results-dashboard.git
git branch -M main
git push -u origin main






# Install dependencies
pip install -r requirements.txt

# Run the script
python generate_excel.py