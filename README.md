# Narrative Companion for Time‑Series Forecasts

A Streamlit app that lets you:

* **Upload** time‑series CSV data
* **Fit** a Prophet forecasting model
* **Generate** concise, GPT‑powered bullet‑point insights
* **Visualize** an interactive, styled time‑series plot
* **Export** a multi‑slide PowerPoint deck with chart and individual insight slides

---

## 🚀 Features

1. **Prophet Forecasting**: Fit ARIMA‑based models to your historical data, with customizable forecast horizon.
2. **LLM Insights**: Leverage GPT‑4o for narrative analysis—trends, seasonality, and anomalies in bullet points.
3. **Interactive Visualization**: Clean, transparent, white‑text time‑series charts with dynamic date formatting.
4. **Presentation Export**: One‑click download of a professional PPTX deck, including:

   * Title slide
   * Forecast plot slide
   * One slide per insight
5. **Configurable UI**: Sidebar controls for data columns, forecast periods, and OpenAI credentials.

---

## 📥 Installation

1. **Clone the repo**

   ```bash
   git clone https://github.com/mishazahid/narrative-forecast-app.git
   cd narrative-forecast-app
   ```

2. **Create & activate** a Python environment (recommended)

   ```bash
   python -m venv .venv
   source .venv/bin/activate      # macOS/Linux
   .\.venv\\Scripts\\activate   # Windows
   ```

3. **Install dependencies**

   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

   > **Note:** `requirements.txt` includes:
   >
   > ```
   > streamlit
   > prophet
   > openai
   > matplotlib
   > python-pptx
   > pandas
   > ```

---

## 🛠 Usage

1. **Obtain an OpenAI API key** and keep it handy.
2. **Run the app**

   ```bash
   streamlit run app.py
   ```
3. **In your browser**:

   * Upload a CSV with a date column and a numeric value column.
   * Select the appropriate columns and forecast horizon.
   * Enter your OpenAI API key.
   * Click **Run Forecast**.
   * View the styled chart and GPT‑generated insights.
   * Click **Download Presentation** to get the PPTX.

---

## 🗂 Sample CSV

```csv
Date,Value
2023-01-01,100
2023-02-01,115
2023-03-01,130
2023-04-01,125
2023-05-01,140
2023-06-01,155
2023-07-01,160
2023-08-01,170
2023-09-01,185
2023-10-01,200
2023-11-01,210
2023-12-01,230
```

---

## 📂 Project Structure

```
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
├── sample.csv          # Example data for quick start
└── README.md           # Project documentation
```

---

## 💡 Customization

* **Forecast model**: Swap Prophet for ARIMA or other libraries.
* **Insight prompts**: Tweak the system/user messages in `generate_insights()`.
* **Styling**: Adjust Matplotlib and PPTX theme colors and fonts.

---

## 📄 License

This project is open‑source under the MIT License. See [LICENSE](LICENSE) for details.
