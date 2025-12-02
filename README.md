# Kingspan Sales Report Generator

This is a Streamlit web application for generating the Kingspan Store Sales + Ops HTML report.

## How to Run Locally

1.  **Install Python** (if not already installed).
2.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Run the App**:
    ```bash
    streamlit run app.py
    ```
4.  The app will open in your browser at `http://localhost:8501`.

## How to Deploy (Free)

1.  **Create a GitHub Repository**:
    *   Upload `app.py` and `requirements.txt` to a new GitHub repo.
2.  **Sign up for Streamlit Community Cloud**:
    *   Go to [share.streamlit.io](https://share.streamlit.io/).
    *   Connect your GitHub account.
3.  **Deploy**:
    *   Click "New App".
    *   Select your repository, branch (main), and file path (`app.py`).
    *   Click "Deploy".

## Features

*   **Multi-language Support**: NL (Default), FR, EN.
*   **Data Processing**: Upload standard Shopify exports (Orders & Customers).
*   **Filtering**: Automatically filters out test accounts and maps Account Managers.
*   **AI Integration**: Connects to Make.com via Webhook to generate insights text.
