from flask import Flask, render_template_string, request

app = Flask(__name__)

# Placeholder
def get_top_counties(industry):
    dummy_results = {
        "Technology": ["Allegheny County", "Montgomery County", "Chester County", "Bucks County", "Centre County"],
        "Healthcare": ["Philadelphia County", "Allegheny County", "Dauphin County", "Lancaster County", "Lehigh County"],
        "Manufacturing": ["York County", "Berks County", "Erie County", "Westmoreland County", "Luzerne County"]
    }
    return dummy_results.get(industry, ["No data available"])


HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <title>Best Counties in PA for Job Growth</title>
    <style>
        body {
            font-family: Arial;
            background-color: #f4f4f4;
            text-align: center;
            padding: 40px;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0px 0px 10px rgba(0,0,0,0.1);
            display: inline-block;
        }
        select, button {
            padding: 10px;
            margin: 10px;
            font-size: 16px;
        }
        .results {
            margin-top: 20px;
            text-align: left;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Best PA Counties for Job Growth</h1>
        <form method="POST">
            <label>Select Industry:</label><br>
            <select name="industry">
                <option value="Technology">Technology</option>
                <option value="Healthcare">Healthcare</option>
                <option value="Manufacturing">Manufacturing</option>
            </select><br>
            <button type="submit">Find Top Counties</button>
        </form>

        {% if results %}
        <div class="results">
            <h2>Top 5 Counties:</h2>
            <ol>
                {% for county in results %}
                <li>{{ county }}</li>
                {% endfor %}
            </ol>
        </div>
        {% endif %}
    </div>
</body>
</html>
"""


@app.route('/', methods=['GET', 'POST'])
def home():
    results = None
    if request.method == 'POST':
        industry = request.form.get('industry')
        results = get_top_counties(industry)
    return render_template_string(HTML_PAGE, results=results)


if __name__ == '__main__':
    app.run(debug=True)
