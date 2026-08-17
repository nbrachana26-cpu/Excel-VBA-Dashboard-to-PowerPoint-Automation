# Excel VBA Dashboard to PowerPoint Automation

## Project Overview

An Excel VBA automation project that converts multiple Excel dashboards into a PowerPoint presentation, with two dashboards automatically placed on each slide.

## Objective

The objective of this project is to reduce the manual effort involved in copying Excel dashboards into PowerPoint presentations and arranging them for reporting.

## Key Features

- Dynamically identifies worksheets with names starting with `Dashboard_`
- Automates PowerPoint presentation creation
- Creates PowerPoint slides automatically
- Places two dashboards on each slide
- Copies dashboard visuals from Excel to PowerPoint
- Automatically positions and resizes dashboard objects
- Uses VBA loops and conditional logic for automation

## Technologies Used

- Microsoft Excel
- Excel VBA
- Microsoft PowerPoint
- PowerPoint Object Library

## Project Workflow

Excel Dashboard Worksheets  
↓  
VBA identifies `Dashboard_` sheets  
↓  
PowerPoint Presentation Created  
↓  
Dashboards Copied from Excel  
↓  
Two Dashboards per Slide  
↓  
PowerPoint Presentation Generated

## Dashboard Structure

The project currently contains:

- Sales Dashboard
- Channel & Payment Dashboard
- Product Performance Dashboard
- Category & Trend Dashboard

## VBA Concepts Used

- Variables and data types
- Excel Worksheet objects
- PowerPoint Application object
- PowerPoint Presentation and Slide objects
- Shape objects
- `For Each` loop
- `If...Then` conditions
- `Mod` operator
- Worksheet name handling
- Copy and Paste automation
- Shape positioning and resizing

## How to Run

1. Download the `.xlsm` file.
2. Open the workbook in Microsoft Excel.
3. Enable macros if prompted.
4. Open the VBA Editor using `Alt + F11`.
5. Run the `DashboardToPowerPoint` macro.
6. The macro creates a PowerPoint presentation containing two dashboards per slide.

## Output

The macro automatically generates a PowerPoint presentation containing the Excel dashboards arranged for presentation and reporting.

## Future Enhancements

- Automatic dashboard sizing based on slide dimensions
- Automatic PowerPoint file naming and saving
- UserForm for selecting dashboards
- Error handling and validation
- Support for different dashboard layouts
