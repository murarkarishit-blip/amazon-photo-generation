# Listing Creative Lab

A concept-to-test workspace for Amazon listing creative. The Home screen explains the flow; open the hamburger menu to work with shopper signals, briefs, image options, and example test results.

The app uses preset rules, has no Amazon connection, and does not generate images itself. Copy a brief into ChatGPT, make image options there, and upload the options to Image studio.

The Brief builder groups repeated topics and narrows workbook data to a prioritized shortlist. It starts at five ideas, with a control to choose between three and seven. Search opportunities are ranked by exposure and click-through rate; shopper questions and competitor gaps are ranked by recurrence. Uploading another workbook refreshes the shortlist. The suggestions are rule-based and do not require an AI API.

The Weekly archive page combines dated uploads, compares two selected weeks, and charts search CTR, search activity, shopper-question frequency, competitor gaps, and test-variant CTR/order-rate progression. The example workbook contains four dated weeks. Search Data, Customer Questions, Competitors, and Test Results use a week_start column. If an uploaded week has no date column, choose its start date in Weekly archive.

When you add a weekly workbook, the app saves an archive file in its local app folder when available. Download the archive backup after each update. Streamlit Community Cloud may clear files created while the app is running, so restore your latest downloaded archive if the app loses it. People with access to the same app share its archive; keep the app private before using business data.

## Run locally

1. Install Python 3.10 or newer.
2. In a terminal in this folder, run `pip install -r requirements.txt`.
3. Run `streamlit run app.py`.
4. Open the local address Streamlit prints.

## Deploy with Streamlit Community Cloud

Upload `app.py`, `requirements.txt`, `README.md`, and `demo_inputs.xlsx` to the root of your GitHub repository. In Streamlit Community Cloud, select **Create app**, choose your repository and `main` branch, set the app file to `app.py`, then deploy.

## Use

Include the .streamlit/config.toml file in the GitHub repository so the charts and controls use the app's palette.

- Open the menu and download `demo_inputs.xlsx` or upload a workbook of your own.
- Review the evidence and select a hypothesis.
- Copy the creative brief/prompt into ChatGPT to generate draft images manually.
- Upload image variants to Image studio for side-by-side review.
- In Weekly archive, add or replace the uploaded week, compare weeks, and download an archive backup.
- Review variant results in Experiment review and their weekly progression in Weekly archive.

## Workbook tabs

`Product`, `Search Data`, `Customer Questions`, `Competitors`, and `Test Results`.

Do not add API keys, credentials, customer personal data, or confidential brand data to this demo repository.
