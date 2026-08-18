# Export Document Tracker & Automated Reminder System

An automated document tracking and reminder system developed using Google Sheets and Google Apps Script.

## Project Overview

This project helps track export-related documents and automatically sends email reminders when required documents are pending.

The system is designed to reduce manual follow-ups and improve document management.

## Key Features

- Automated tracking of export documents
- Invoice, Packing List, Customer Bill, E-Way Bill and Airway Bill status tracking
- Automatic email reminders for pending documents
- Conditional reminder logic based on invoice amount
- INR value calculation based on currency and exchange rate
- E-Way Bill reminder logic for invoices above ₹50,000
- Next reminder date tracking
- Reminder count tracking
- Automatic status management
- Prevents further reminders after documents are received/completed

## Technologies Used

- Google Sheets
- Google Apps Script
- Google Apps Script Triggers
- Email Automation

## Workflow

1. Invoice details are entered into the Document Tracker.
2. Currency and invoice amount are recorded.
3. The system calculates the INR value.
4. Required document statuses are updated.
5. The automation checks pending documents and applicable conditions.
6. If a reminder is required, an automated email is sent to the responsible senior.
7. Once the required document is received/completed, further reminders are stopped.

## Documents Tracked

- Invoice
- Packing List
- Customer Bill
- E-Way Bill
- Airway Bill

## Project Objective

The main objective of this project is to automate document follow-ups, reduce manual tracking and ensure that pending export documents are followed up on time.

## Author

Utkrsta
