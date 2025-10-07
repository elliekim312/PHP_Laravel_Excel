# PHP Laravel Excel

A PHP Laravel service for bulk uploading RMA (Return Merchandise Authorization) order items via Excel file.

## Overview

This service processes Excel files containing RMA order data, validates the information, and inserts it into the database while preventing duplicates.

## Requirements

- PHP 7.4 or higher
- Laravel Framework
- Maatwebsite Excel package
- MySQL Database

## Installation

Install the required Excel package:

```bash
composer require maatwebsite/excel
```

## Features

- **Bulk Upload**: Process multiple RMA orders from Excel files
- **Duplicate Prevention**: Automatically skips duplicate entries based on RMA order number, line number, and serial number
- **Customer Management**: Automatically creates new customers if they don't exist
- **Validation**: Validates required fields and column count
- **Multi-sheet Support**: Reads the first two sheets of the Excel file

## Excel File Format

### Required Columns (Minimum 22 columns)

| Column | Field | Required | Description |
|--------|-------|----------|-------------|
| 0 | Transaction Date | Yes | Date of transaction |
| 1 | Customer Code | Yes | Unique customer identifier |
| 2 | Customer Name | Yes | Name of customer |
| 3 | Return Type | Yes | Type of return |
| 4 | Source Header No | Yes | RMA order number |
| 5 | Source Line No | Yes | Line number in order |
| 6 | Source Header ID | No | Header ID |
| 10 | Item Code | No | Product item code |
| 11 | Order Quantity | No | Quantity ordered |
| 12 | Serial No | No | Product serial number |
| 13 | Organization Code | No | Organization identifier |
| 14 | Subinventory Code | No | Subinventory location |
| 20 | Ship to City | No | Shipping city |
| 21 | Ship to State | No | Shipping state |

### File Requirements

- Maximum 2 sheets will be processed
- Skip top 2 rows (headers start at row 3)
- No extra columns beyond column 22
- All required fields must be filled

## How It Works

1. **File Loading**: Loads Excel file and reads the first two sheets
2. **Header Validation**: Checks for extra columns (max 22 allowed)
3. **Data Validation**: Validates required fields for each row
4. **Duplicate Check**: Queries database to check if RMA order item already exists
5. **Data Insertion**: Inserts new records into `receive_contents` and `receive_details` tables
6. **Customer Management**: Creates new customer records if they don't exist
7. **Response**: Returns count of uploaded and skipped records

## Database Tables

### receive_contents
Stores RMA order header information:
- Transaction date
- Customer information
- Return type
- Source header details
- Shipping information
- Status

### receive_details
Stores RMA order line item details:
- Item code
- Quantity
- Serial number
- Organization code
- Subinventory code
- Status

### customer
Stores customer information:
- Customer code
- Customer name

## Error Handling

- **Transaction Rollback**: All database operations are wrapped in transactions
- **Exception Handling**: Catches and returns user-friendly error messages
- **Memory Management**: Increased memory limit for large files

## Usage

```php
use Modules\OMM\Service\ReportServiceImpl;

$reportService = new ReportServiceImpl();

$dataArray = [
    'system_id' => 'your_system_id'
];

$result = $reportService->uploadExcelForRmaOrderItemBulkInsert($dataArray, $uploadFile);

if (isset($result['status_code']) && $result['status_code'] == 200) {
    echo $result['success_data'];
} else {
    echo $result['error'];
}
```

## Response Format

### Success Response
```php
[
    'status_code' => 200,
    'success_data' => 'Total Uploaded Count: 50, Total Skipped Count: 5'
]
```

### Error Responses

**Missing Required Fields:**
```php
['error' => '1 Sheet: Row 7 is missing required fields.']
```

**Extra Columns:**
```php
['error' => '1 Sheet: has extra column(s). Please check again.']
```

**Empty Data:**
```php
['error' => 'Excel Template Data Empty.']
```
