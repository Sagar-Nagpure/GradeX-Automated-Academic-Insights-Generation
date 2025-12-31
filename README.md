# 📊 GradeX – Automated Academic Insights & Reports Generation Platform

GradeX is a powerful and intuitive Streamlit web application designed to automate the analysis of student academic performance. By uploading standardized Excel marksheets for different examination types (Class Tests, Mid-Sems, End-Sems), educators can instantly generate comprehensive class-level and individual student reports, complete with interactive visualizations and exportable PDFs.

## ✨ Features

-   **Dual Exam Type Support**: Analyze both "CT/Mid" and "End Sem" exam data with dedicated processing logic.
-   **Template-Driven**: Provides downloadable Excel templates to ensure data consistency and accurate parsing.
-   **Comprehensive Class Analysis**:
    -   Identify Top and Bottom 3 performers.
    -   Calculate and visualize class averages vs. highest scores.
    -   Subject-wise performance breakdown and difficulty ranking.
    -   Pass/Fail analysis for the entire class and individual subjects.
    -   K-Means clustering to group students into performance tiers.
-   **In-Depth Individual Analysis**:
    -   View individual student marks compared against the class average.
    -   Track student performance trends across different assessment types (CA, MSE, ESE).
    -   Detailed metrics including total marks, percentage, class rank, and percentile.
-   **Interactive Visualizations**: Leverages Plotly for a rich, interactive user experience with bar charts, pie charts, scatter plots, and box plots.
-   **PDF Report Generation**:
    -   Export detailed, multi-page class-wise analysis reports.
    -   Generate and download individual student report cards in PDF format.
    -   Batch export all individual student reports into a single PDF document.

## 🛠️ Technologies Used

-   **Frontend**: Streamlit
-   **Data Manipulation**: Pandas, NumPy
-   **Data Visualization**: Plotly
-   **Machine Learning**: Scikit-learn (for clustering)
-   **PDF Generation**: ReportLab
-   **Excel Handling**: openpyxl
-   **Plotly Image Export**: kaleido

## ✅ Getting Started

To run GradeX on your local machine, follow these steps.

### Prerequisites

-   Python 3.8+
-   `pip` package manager

### Installation & Setup

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/Sagar-Nagpure/GradeX.git
    cd GradeX
    ```

2.  **Create and activate a virtual environment (recommended):**
    ```bash
    # For Unix/macOS
    python3 -m venv venv
    source venv/bin/activate

    # For Windows
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Run the Streamlit application:**
    ```bash
    streamlit run app.py
    ```
    Your web browser should automatically open to the application's local URL.

## 📋 How to Use

1.  **Download a Template**: From the sidebar, download either the "Blank CT/Mid Template" or the "Blank End Sem Template" depending on your needs. You can also download a filled sample for reference.
2.  **Fill the Data**: Open the downloaded Excel template and fill in the student data according to the predefined columns. **Do not alter the structure or headers of the template.**
3.  **Select Exam Type**: In the main application window, select the correct exam type ("CT/Mid" or "End Sem") that matches your data file.
4.  **Upload File(s)**: Drag and drop or browse to upload one or more of your filled Excel files. The application will process each file and generate a separate analysis tab.
5.  **Explore Insights**:
    -   Navigate through the tabs for each uploaded file.
    -   Examine the class-level dashboards for an overview of performance.
    -   Select individual students from the dropdown menu to view their detailed reports and charts.
6.  **Download Reports**:
    -   Click the "Download Classwise Report" button to get a comprehensive PDF summary of the entire class.
    -   For individual reports, select a student and click "Download Individual Student PDF".
    -   Use the "Download All Student Reports" button to generate a single PDF containing reports for every student in the uploaded file.


---
