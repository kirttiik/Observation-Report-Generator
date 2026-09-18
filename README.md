# Observation Report Generator

A full-stack application designed to automate the process of generating observation reports from Excel data. It takes in multiple term reports, calculates target observations, aggregates the data for each school, and outputs a formatted Excel workbook along with individual PNG image snapshots for each school.

## Features

- **Data Aggregation**: Processes an entire year's base report and optional term reports.
- **Excel Formatting**: Generates clean, bordered, and colored Excel sheets for each school.
- **Image Generation**: Automatically creates a styled PNG image of each school's observation table.
- **Batch Export**: Zips the final Excel file and all school images into a single downloadable archive.
- **Custom Mapping**: Allows custom 1-based index mapping for columns directly from the UI.

## Tech Stack

- **Frontend**: React, Vite
- **Backend**: FastAPI, Pandas, openpyxl, Pillow

## Local Development

### Prerequisites
- Node.js
- Python 3.8+

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/kirttiik/Observation-Report-Generator.git
   cd Observation-Report-Generator
   ```

2. **Backend Setup**
   ```bash
   cd api
   pip install -r requirements.txt
   python -m uvicorn index:app --reload
   ```
   The backend will run on `http://localhost:8000`.

3. **Frontend Setup**
   Open a new terminal window:
   ```bash
   cd frontend
   npm install
   npm run dev
   ```
   The frontend will run on `http://localhost:5173`. API requests are proxied automatically.

## Deployment

This project is configured as a monorepo optimized for deployment on Vercel. 

1. Import the project into Vercel.
2. Keep the Framework Preset as **Other**.
3. Keep the Root Directory as `./`.
4. Set the **Output Directory** to `frontend/dist`.
5. Deploy. Vercel will automatically build the frontend and serve the FastAPI backend as serverless functions.
