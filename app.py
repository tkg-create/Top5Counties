from flask import Flask, render_template_string, request
import main

app = Flask(__name__)

# Load data once
data = main.load_and_process_data()
industries = main.get_industries(data)

# Placeholder replaced with actual function
def get_top_counties(industry):
    return main.get_top_counties(data, industry)


HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <title>Best Counties in PA for Job Growth</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
        canvas {
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Best PA Counties for Job Growth</h1>
        <form method="POST">
            <label>Select Industry:</label><br>
            <input type="text" name="industry" list="industries" placeholder="Type to search industries"><br>
            <datalist id="industries">
                {% for industry in industries %}
                <option value="{{ industry }}">
                {% endfor %}
            </datalist><br>
            <button type="submit">Find Top Counties</button>
        </form>

        {% if results %}
        <div class="results">
            <h2>Top 5 Counties:</h2>
            <ol>
                {% for county, score in results %}
                <li>{{ county }} (Score: {{ "%.2f"|format(score) }})</li>
                {% endfor %}
            </ol>
        </div>
        <canvas id="myChart" width="400" height="200"></canvas>
        <script>
            var ctx = document.getElementById('myChart').getContext('2d');
            var labels = [{% for county, score in results %}"{{ county }}",{% endfor %}];
            var data = [{% for county, score in results %}{{ score }},{% endfor %}];
            var myChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Scores',
                        data: data,
                        backgroundColor: 'rgba(75, 192, 192, 0.2)',
                        borderColor: 'rgba(75, 192, 192, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        </script>
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
    return render_template_string(HTML_PAGE, results=results, industries=industries)


if __name__ == '__main__':
    app.run(debug=True)
