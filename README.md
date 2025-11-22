# App-Script

A curated collection of Google Apps Script utilities for enhanced productivity and automation within Google Workspace.

---

## 🚀 Overview

This repository serves as a personal collection of Google Apps Script projects developed by Ayo-G. It encompasses a variety of scripts designed to automate repetitive tasks, synchronize data between different platforms, and extend the functionality of Google Workspace applications like Google Sheets.

Whether you're looking to streamline your data management, automate reporting, or create custom workflows, these scripts offer practical solutions to common challenges faced by users and developers working within the Google ecosystem. Each script is self-contained and focuses on a specific functionality, making them easy to adapt and integrate into your own Google Workspace environment.

## ✨ Features

The scripts in this repository provide a range of functionalities, including:

*   **Google Sheets Automation**: Advanced formatting, data refreshing, and manipulation within Google Sheets.
*   **Data Synchronisation**: Seamlessly sync data between Google Sheets and external services like Airtable.
*   **Custom Workflows**: Implement specific business logic, such as linking users to stories or other data points.
*   **API Integration**: Interact with third-party APIs (e.g., Airtable) directly from Google Apps Script.
*   **Time-Saving Utilities**: Automate tasks that would otherwise require manual intervention, saving time and reducing errors.

## 🛠️ Tech Stack

*   **Main Language**: JavaScript
*   **Platform**: Google Apps Script
*   **Integrated Services**:
    *   Google Sheets API
    *   Airtable API

## 📂 Repository Structure

The repository contains several independent Google Apps Script files, each designed for a specific purpose:

```
.
├── LICENSE
├── format_refresh_google_sheets.js
├── linkUserstoStories.js
├── sync_from_airtable.js
└── sync_to_airtable.js
```

*   `LICENSE`: Contains the licensing information for the project.
*   `format_refresh_google_sheets.js`: Script for automating formatting and refreshing data within Google Sheets.
*   `linkUserstoStories.js`: A utility script likely for linking user data to specific "stories" or tasks, potentially within a project management context.
*   `sync_from_airtable.js`: Script to pull data from Airtable into a Google Sheet or another Google Workspace application.
*   `sync_to_airtable.js`: Script to push data from a Google Sheet or another Google Workspace application to Airtable.

## 🚀 Installation & Setup

Google Apps Script projects are typically hosted and run within the Google Cloud environment. There's no traditional "installation" process like `npm install`. Instead, you'll import these scripts into your Google Apps Script projects.

### Steps:

1.  **Create a New Google Apps Script Project**:
    *   Go to [script.google.com/home](https://script.google.com/home) and click `New project`.
    *   Alternatively, from a Google Sheet, go to `Extensions > Apps Script`.

2.  **Copy the Script Content**:
    *   Open the `.js` file you wish to use from this repository (e.g., `format_refresh_google_sheets.js`).
    *   Copy its entire content.
    *   Paste the content into the `Code.gs` file (or a new `.gs` file) in your Apps Script project, replacing any existing code.

3.  **Configure Project Properties (if applicable)**:
    *   Some scripts (e.g., Airtable sync scripts) may require API keys or other configuration values.
    *   Go to `Project settings` (the gear icon on the left sidebar).
    *   Under `Script properties`, add new properties for any required keys (e.g., `AIRTABLE_API_KEY`, `AIRTABLE_BASE_ID`, `AIRTABLE_TABLE_NAME`).
    *   **Important**: Never hardcode sensitive information directly into your script files. Use Script Properties or the Properties Service.

4.  **Authorize the Script**:
    *   When you run a script for the first time, Google will prompt you to authorise it to access your Google services (e.g., Google Sheets, external services).
    *   Click `Review permissions`, select your Google account, and grant the necessary permissions.

5.  **Set Up Triggers (Optional)**:
    *   If you want a script to run automatically (e.g., every hour, on form submission), you'll need to set up a trigger.
    *   In the Apps Script editor, click the `Triggers` icon (the clock icon on the left sidebar).
    *   Click `Add Trigger`, select the function to run, choose the event source (e.g., `Time-driven`), and configure the time interval.

## 💡 Usage

Once installed and configured, you can run the scripts manually or via triggers.

### Example: `sync_from_airtable.js`

This script is designed to pull data from Airtable into a Google Sheet.

1.  **Preparation**:
    *   Ensure you have an active Google Sheet where the data will be imported.
    *   Obtain your Airtable API Key, Base ID, and Table Name.
    *   Add these as `Script properties` in your Apps Script project:
        *   `AIRTABLE_API_KEY`: `YOUR_AIRTABLE_API_KEY`
        *   `AIRTABLE_BASE_ID`: `YOUR_AIRTABLE_BASE_ID`
        *   `AIRTABLE_TABLE_NAME`: `YOUR_AIRTABLE_TABLE_NAME`
        *   `SHEET_NAME`: `The name of your Google Sheet tab` (e.g., "Airtable Data")

2.  **Running the Script**:
    *   In the Apps Script editor, select the `syncFromAirtable` function from the dropdown menu at the top.
    *   Click the `Run` button (play icon).
    *   The script will fetch data from your specified Airtable table and populate it into the designated Google Sheet tab.

3.  **Automating with Triggers**:
    *   Set up a time-driven trigger to run `syncFromAirtable` every hour, day, or at a custom interval to keep your Google Sheet updated.

---

## 🤝 Contributing

Contributions are welcome! If you have improvements, bug fixes, or new script ideas, please follow these steps:

1.  **Fork** the repository.
2.  **Clone** your forked repository:
    ```bash
    git clone https://github.com/YOUR_USERNAME/App-Script.git
    ```
3.  **Create a new branch** for your feature or bug fix:
    ```bash
    git checkout -b feature/your-feature-name
    ```
    or
    ```bash
    git checkout -b bugfix/issue-description
    ```
4.  **Make your changes**.
5.  **Commit your changes** with a clear and descriptive message:
    ```bash
    git commit -m "feat: Add new script for X functionality"
    ```
6.  **Push your changes** to your forked repository:
    ```bash
    git push origin feature/your-feature-name
    ```
7.  **Open a Pull Request** to the original `Ayo-G/App-Script` repository, describing your changes in detail.

## 📄 License

This project is licensed under the [MIT License](LICENSE).
