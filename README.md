[![Releases](https://img.shields.io/badge/Releases-Download%20Workbook-blue?logo=github)](https://github.com/zhaa-kun/analytical-models-in-excel/releases)

# Spreadsheet Machine Learning: Excel Models for Core Analytics Demo 📊🔧

Short description
- A curated Excel workbook that shows core data analysis techniques. It uses spreadsheet formulas, structured sheets, and clear formatting to show how regression, classification, dimensionality reduction, and validation work without code.

Badges
- Topics: analytics · cross-validation · data-analysis · data-analytics · data-science-portfolio · data-visualization · decision-trees · excel · excel-models · knn · linear-regression · logistic-regression · machine-learning · no-code-machine-learning · pca · predictive-modeling · spreadsheet-models · statistical-analysis

Hero image
![Excel Icon](https://upload.wikimedia.org/wikipedia/commons/7/73/Microsoft_Office_Excel_%282018%E2%80%93present%29.svg)

Why this repo
- Use it to teach model logic inside a spreadsheet.
- Show model mechanics step-by-step for presentations or classes.
- Audit model math in cells, not in black-box code.
- Share a portfolio piece that highlights Excel modeling skills.

What you will find
- A single, well-structured Excel workbook (.xlsx) with multiple sheets.
- Worked examples and small datasets embedded in the workbook.
- Clean layout for inputs, calculations, and outputs.
- Visuals: charts and conditional formatting for model insight.
- Reusable templates for experiments.

Get the workbook
- Download the release file from the Releases page and open it in Excel. Execute the workbook by opening the downloaded file and stepping through the sheets.
- Releases: https://github.com/zhaa-kun/analytical-models-in-excel/releases

Structure of the workbook (sheet-by-sheet)
- README (sheet): Quick navigation and short guide.
- Data: Small sample datasets for each demo. Columns include features and targets.
- Preprocess: Missing value handling, scaling, and simple encoding done with formulas.
- Linear Regression (OLS): Full OLS via matrix algebra with formulas. Includes residual plots and diagnostics.
- Logistic Regression: Logit link implemented with iterative update (Newton-Raphson) and log-likelihood tracking.
- k-NN: Distance matrix, neighbor selection, tie rules, and performance table.
- Decision Tree (simple): Split metrics (Gini, entropy), split selection process, and manual tree diagram using cells.
- PCA: Covariance matrix, eigen decomposition via characteristic polynomial approximation, variance explained table.
- Cross-Validation: k-fold split by index formulas, aggregate metrics, and bias-variance illustration.
- Model Comparison: Side-by-side metrics and a chart for AUC, RMSE, accuracy, and explained variance.
- Notes (calc): Key formula references and aliases to cells that hold hyperparameters.
- Visuals: Chart examples and interactive controls (data validation lists, slider-style cells).
- Tests: Small suites of formula checks that validate expected numeric results.

Models and methods covered
- Linear Regression: Ordinary Least Squares; matrix solution with X'X inversion using spreadsheet functions. Diagnostics: R², adjusted R², standard errors, t-stats, residual plots.
- Logistic Regression: Iterative parameter update and probability output. Model fit via log-likelihood. Metrics: accuracy, precision, recall, ROC data points.
- k-Nearest Neighbors (k-NN): Euclidean distance matrix, vote aggregation, weighted k versions.
- Decision Tree (tiny): Manual split search with impurity reduction and depth control.
- Principal Component Analysis (PCA): Center data, compute covariance, and derive principal components and explained variance.
- Cross-Validation: k-fold, stratified split for classification demos, and metric aggregation.

How the sheets show the work
- Inputs live at the top of each sheet.
- Calculations sit in a dedicated block with named ranges.
- Outputs and charts sit in a right-hand column for quick review.
- Key steps use simple formulas so any user can trace a number from input to output.
- Comments and short cell notes explain the formula intent.

Sample workflows
- Fit a linear model
  1. Open the workbook and go to the Linear Regression sheet.
  2. Review the Inputs block and change the design matrix if needed.
  3. Watch the matrix algebra section compute coefficients.
  4. Check residual diagnostics and scatter plots.
- Train and test a classifier
  1. Use the Cross-Validation sheet to set k (folds).
  2. The sheet splits data using index math and validation formulas.
  3. Go to Logistic Regression or k-NN sheet and view metrics per fold.
  4. Inspect aggregated metrics on the Model Comparison sheet.
- Run PCA
  1. Center and scale features in Preprocess.
  2. Open PCA sheet to see covariance and component scores.
  3. Use the Visuals sheet to plot explained variance.

Key formulas and Excel features used
- INDEX, MATCH, OFFSET, INDIRECT for structured references.
- MMULT, MINVERSE, TRANSPOSE for matrix math.
- SUMPRODUCT for dot products and weighted sums.
- IF, COUNTIFS, SUMIFS for logic and grouped aggregates.
- conditional formatting for residuals and outlier flags.
- chart types: scatter, line, bar for metrics and diagnostics.
- data validation lists for toggles between model variants.

Teaching tips
- Freeze panes on calculation blocks so learners track formula flow.
- Use cell colors for input, calc, and output to set expectations.
- Step through iterations (logistic Newton steps) by copying intermediate columns.
- Use the Tests sheet to run quick checks after changes.
- Replace sample data with your dataset. The workbook uses named ranges for easy swap.

Performance notes
- The workbook handles small to medium datasets that fit in one sheet.
- Large datasets may slow spreadsheet calculations. Use smaller samples for teaching.
- Matrix inversion via MINVERSE can show numerical issues; the sheet has a small demo of conditioning.

Examples and visuals
- Residual plot: shows predicted vs actual and highlights heteroskedastic patterns.
- ROC pseudo-curve: sorted thresholds and TPR/FPR computed in-sheet.
- PCA scree plot: bar chart of explained variance per component.
- Decision split table: shows candidate thresholds with impurity metrics.

Contributing
- Pull requests welcome. Use clear issue descriptions and attach small test data.
- Suggested contributions:
  - Add new model sheet (SVM, Lasso) using cell formulas.
  - Improve numeric stability for matrix math.
  - Add workbook macros for automation (kept separate from core formulas).
  - Improve visuals or add tutorial steps in separate sheets.

Releases and download
- The workbook is packaged in Releases. Download the .xlsx from the releases page and open it in Excel to run the demos.
- Releases page: https://github.com/zhaa-kun/analytical-models-in-excel/releases
- If the link does not work in your environment, check the Releases section on the repository page.

License
- MIT License. See LICENSE file in the repo for terms.

Contact
- File issues on GitHub for bugs or feature requests.
- Use pull requests for changes to workbook structure or new model sheets.

Examples of classroom exercises
- Exercise 1 — OLS step trace
  - Goal: Reproduce coefficient for a single predictor using cell-by-cell math.
  - Task: Follow matrix build, compute X'X and invert, obtain beta.
  - Deliverable: A screenshot of cell ranges with matching coefficient values.
- Exercise 2 — Evaluate k-NN sensitivity
  - Goal: Compare accuracy across k = 1,3,5.
  - Task: Use the k-NN sheet and the cross-validation split to record fold metrics.
  - Deliverable: Table and chart of accuracy vs k.
- Exercise 3 — Bias-variance sketch
  - Goal: Visualize training vs validation error as model complexity changes.
  - Task: Use Decision Tree sheet, vary depth, record RMSE or error rates.
  - Deliverable: Line chart of errors across depths.

Advanced notes for readers who want depth
- Numerical conditioning: The Linear Regression sheet has an example that shows how near-collinearity affects MINVERSE results. It uses a small ridge stabilizer formula to show shrinkage.
- Logistic convergence: The sheet logs Newton-Raphson steps and shows change in log-likelihood. This helps teach convergence and step damping.
- PCA eigen approximation: Full eigen decomposition is hard in raw Excel. The sheet uses power-iteration style approximations and orthogonal deflation for component extraction.

Screenshots
- Add your own screenshots to issues or PRs. The workbook includes a Visuals sheet with sample charts that you can capture.

Repository topics (for search)
analytics, cross-validation, data-analysis, data-analytics, data-science-portfolio, data-visualization, decision-trees, excel, excel-models, knn, linear-regression, logistic-regression, machine-learning, manual-calculations, no-code-machine-learning, pca, predictive-modeling, spreadsheet-models, statistical-analysis

Credits
- Built to show in-sheet model mechanics and clear workbook design.
- Inspired by classroom needs for transparent model math.

Legal and data hygiene
- Remove sensitive data before uploading new workbook versions.
- Use sample data sheets for public demonstrations.

Quick start checklist
- Download the release from the Releases page and open the .xlsx file in Excel.
- Read the README sheet inside the workbook to find the demo you want.
- Run tests on the Tests sheet to validate formulas.
- Modify inputs and watch outputs update.

Additional resources
- Link to Excel function docs and matrix formula examples: use official Microsoft documentation for advanced formula behavior.
- For numeric stability and linear algebra background, refer to standard statistics and linear algebra texts.

Thank you for exploring the workbook.