# Make.com Scenario Setup Guide

This guide explains how to set up the automation in Make.com (formerly Integromat) to generate the "Challenges & Opportunities" text using AI.

## 1. Create a New Scenario

1.  Log in to Make.com.
2.  Click **"Create a new scenario"**.

## 2. Add the Trigger: Custom Webhook

1.  Click the big **+** button and search for **"Webhooks"**.
2.  Select **"Custom webhook"**.
3.  Click **"Add"** to create a new webhook.
    *   **Name:** `Kingspan Report Generator`
    *   **IP restrictions:** Leave empty (or restrict to Streamlit Cloud IPs if known).
4.  Click **"Save"**.
5.  **Copy the URL** (e.g., `https://hook.eu1.make.com/...`).
    *   *Paste this URL into the Streamlit App's sidebar.*
6.  **Important:** The webhook needs to learn the data structure.
    *   Click **"Re-determine data structure"**.
    *   Go to your Streamlit App (running locally or deployed).
    *   Upload files, select a period, paste the Webhook URL, and click **"✨ Generate Insights with AI"**.
    *   Make.com should say "Successfully determined".

## 3. Add the AI: OpenAI (ChatGPT)

1.  Add a new module next to the Webhook.
2.  Search for **"OpenAI"** and select **"Create a completion (Chat Completions)"**.
3.  **Connection:** Connect your OpenAI API Key.
4.  **Model:** Select `gpt-4o` or `gpt-3.5-turbo`.
5.  **Messages:**
    *   **Role:** `System`
    *   **Content:**
        ```text
        You are a helpful sales assistant for Kingspan Light + Air.
        Write a short, professional paragraph (max 50 words) summarizing the sales performance and suggesting a focus for next month.
        Write in the language: {{language}}
        ```
    *   **Role:** `User`
    *   **Content:**
        ```text
        Period: {{period}}
        Revenue: €{{revenue}}
        Orders: {{orders}}
        New Customers: {{new_customers}}
        Top Product: {{top_product}}
        
        Write a "Challenges & Opportunities" insight paragraph.
        ```
        *(Map the purple variables from the Webhook module to these fields)*.

## 4. Add the Response: Webhook Response

1.  Add a new module next to OpenAI.
2.  Search for **"Webhooks"** and select **"Webhook Response"**.
3.  **Status:** `200`
4.  **Body:** Map the output from OpenAI: `Choices[] -> Message -> Content`.
    *   It should look like: `{{1.choices[].message.content}}`

## 5. Turn It On

1.  Click **"Run once"** to test it again from the App.
2.  If it works, toggle the **"Scheduling"** switch to **ON** (at the bottom left).
