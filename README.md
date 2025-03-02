# ExamLSGen

## Overview
ExamLSGen is a VBA script designed to generate multiple-choice assessments using a predefined question bank. It automates the process of selecting, shuffling, and formatting questions, ensuring variability and efficiency in test creation.

## Features
Category Selection – Users can specify the number of categories to choose from.
Customizable Question Count – Define the number of multiple-choice examples per question.
Question Prompt Selection – Customize the text prompt for each generated MCA question.
Randomization – Option to shuffle examples within each category for varied question sets.
Formatted Output – Automatically structures the output for easy copy-pasting into Word or other documents.

## How It Works
1. User inputs the number of categories to select from.
2. User inputs how many examples should appear per question.
3. User inputs question prompt.
4. User optionally shuffles the order of examples within each category.
5. The script compiles the questions into a formatted list, outputting them in a new worksheet.
6. Each question and its corresponding multiple-choice options are structured for easy integration into exam documents.

## Usage
1. Open the Excel workbook containing the question bank.
2. Run the VBA script.
3. Follow the on-screen prompts to configure the exam settings.
4. Copy and paste the formatted output into your desired document.

## Notes
- Ensure that the question bank is structured correctly with categories in separate columns.
- The script will generate a new worksheet named "Combinations" to store the output.
- Users can modify the script to further customize output formatting if needed.

## Requirements
- Microsoft Excel with VBA enabled.
- A structured question bank with categorized content.
