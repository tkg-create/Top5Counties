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
            font-family: "Georgia", serif;
            background: url('static/bg.png');
            background-size: cover;
            background-position: center 40px;
            color: #eae0c8;
            margin: 0;
            padding: 0;
            text-align: center;
        }

        header {
            background: linear-gradient(to right, #3b2f2f, #1b1b1b);
            color: #d4af37;
            padding: 25px;
            border-bottom: 2px solid #d4af37;
        }

        .container {
            background-color: #2a2a2a;
            margin: 40px auto;
            padding: 30px;
            width: 80%;
            max-width: 700px;
            border-radius: 10px;
            border: 1px solid #444;
            box-shadow: 0px 4px 12px rgba(0,0,0,0.6);
        }

        h1, h2 {
            color: #d4af37;
        }

        input, select {
            padding: 10px;
            margin: 10px;
            width: 90%;
            border-radius: 6px;
            border: 1px solid #555;
            background-color: #1b1b1b;
            color: #eae0c8;
        }

        button {
            padding: 10px 20px;
            margin-top: 10px;
            font-size: 16px;
            background-color: #d4af37;
            border: none;
            border-radius: 6px;
            color: #1b1b1b;
            cursor: pointer;
            transition: 0.3s;
        }

        button:hover {
            background-color: #b8962e;
        }

        .results {
            margin-top: 20px;
            text-align: left;
        }

        ol {
            padding-left: 20px;
        }

        canvas {
            margin-top: 20px;
            background-color: #1b1b1b;
            padding: 10px;
            border-radius: 8px;
        }
    </style>
</head>

<body>

<header>
    <h1>Best Counties in Pennsylvania for Job Growth</h1>
</header>

<div class="container">

    <form method="POST">
        <label>Select Industry:</label><br>
        <input type="text" name="industry" list="industries" placeholder="Type to search industries"><br>

        <datalist id="industries">
            {% for industry in industries %}
            <option value="{{ industry }}">
            {% endfor %}
        </datalist>

        <button type="submit">Find Top Counties</button>
    </form>

    {% if results %}
    <div class="results">
        <h2>Top 5 Counties</h2>
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
                    label: 'Growth Score',
                    data: data,
                    backgroundColor: 'rgba(212, 175, 55, 0.3)',
                    borderColor: '#d4af37',
                    borderWidth: 2
                }]
            },
            options: {
                plugins: {
                    legend: {
                        labels: {
                            color: '#eae0c8'
                        }
                    }
                },
                scales: {
                    x: {
                        ticks: {
                            color: '#eae0c8'
                        }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            color: '#eae0c8'
                        }
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
