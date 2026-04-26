# Top5Counties

## Description

Top5Counties is a web application that analyzes job growth data for counties in Pennsylvania. Users select an industry and view the top 5 counties with the highest job growth rates, visualized with a bar graph.

## Features

- Interactive industry selection
- Top 5 counties ranking by job growth
- Visual chart
- Data processing with machine learning models

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/tkg-create/Top5Counties.git
   cd Top5Counties
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure you have the data file `qcew_full_report.xlsx` in the root directory (this file contains the QCEW data for Pennsylvania counties).

## Usage

Run the Flask application:
```bash
python app.py
```

Open your web browser and navigate to `http://127.0.0.1:5000/` to access the application.

Select an industry from the dropdown, and the app will display the top 5 counties with the highest job growth for that industry, along with a chart.

## Technologies Used

- **Flask**: Web framework for Python
- **Pandas**: Data manipulation and analysis
- **Scikit-learn**: Machine learning library for data processing
- **Chart.js**: JavaScript library for interactive charts
- **OpenPyXL**: For reading Excel files

## Data Source

The application uses data from the Quarterly Census of Employment and Wages (QCEW) program by the U.S. Bureau of Labor Statistics.
