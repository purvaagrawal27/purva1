# MPCC Connect Backend - Excel Upload for Office Bearers

This backend API allows uploading Excel files to create office bearers in the database with comprehensive validation.

## Features

1. ✅ Accepts Excel documents (.xlsx, .xls)
2. ✅ Validates required columns (Name, Email)
3. ✅ Validates that required fields are not empty
4. ✅ Validates email format and checks for duplicates
5. ✅ Uploads data to database and creates office bearers on success
6. ✅ Returns specific error messages for validation failures

## Installation

1. Install dependencies:
```bash
npm install
```

2. Create a `.env` file in the root directory with the following variables:
```env
DB_HOST=localhost
DB_PORT=5432
DB_NAME=your_database_name
DB_USER=your_database_user
DB_PASSWORD=your_database_password
PORT=3000
NODE_ENV=development
```

3. Make sure PostgreSQL is running and the database exists.

## API Endpoints

### Upload Excel File
**POST** `/api/office-bearers/upload`

Upload an Excel file to create office bearers.

**Request:**
- Method: POST
- Content-Type: multipart/form-data
- Body: Form data with field name `excelFile` containing the Excel file

**Required Excel Columns:**
- `Name` (required, cannot be empty)
- `Email` (required, cannot be empty, must be valid email format)

**Optional Excel Columns:**
- `Phone`
- `Position`
- `Department`
- `Address`

**Success Response (201):**
```json
{
  "success": true,
  "message": "Successfully created 5 office bearer(s)",
  "count": 5,
  "data": [
    {
      "id": 1,
      "name": "John Doe",
      "email": "john@example.com",
      "phone": "1234567890",
      "position": "President",
      "department": "Executive",
      "address": "123 Main St"
    }
  ]
}
```

**Error Response (400):**
```json
{
  "success": false,
  "error": "Validation failed",
  "message": "Excel file validation failed",
  "errors": [
    {
      "type": "missing_column",
      "column": "Email",
      "message": "Required column \"Email\" is missing from the Excel file"
    }
  ]
}
```

### Get All Office Bearers
**GET** `/api/office-bearers`

Retrieve all office bearers from the database.

**Success Response (200):**
```json
{
  "success": true,
  "count": 10,
  "data": [...]
}
```

## Excel File Format

Your Excel file should have the following structure:

| Name       | Email              | Phone      | Position   | Department | Address        |
|------------|--------------------|------------|------------|------------|----------------|
| John Doe   | john@example.com   | 1234567890 | President  | Executive  | 123 Main St    |
| Jane Smith | jane@example.com   | 0987654321 | Secretary  | Admin      | 456 Oak Ave    |

**Notes:**
- Column names are case-insensitive
- First row should contain column headers
- Name and Email are required
- Email must be in valid format
- Duplicate emails (both in file and database) will be rejected

## Validation Rules

1. **File Type**: Only Excel files (.xlsx, .xls) are accepted
2. **File Size**: Maximum 10MB
3. **Required Columns**: Name and Email must be present
4. **Required Fields**: Name and Email cannot be empty
5. **Email Format**: Must be a valid email address
6. **Unique Emails**: No duplicate emails within the file or in the database

## Error Types

- `missing_column`: Required column is missing
- `empty_field`: Required field is empty
- `invalid_format`: Email format is invalid
- `duplicate_emails_in_file`: Same email appears multiple times in the file
- `duplicate_emails_in_database`: Email already exists in database
- `parsing_error`: Error parsing the Excel file

## Running the Server

Development mode (with auto-reload):
```bash
npm run dev
```

Production mode:
```bash
npm start
```

The server will start on the port specified in `.env` (default: 3000).

## Database Schema

The `office_bearers` table has the following structure:

- `id`: Primary key (auto-increment)
- `name`: VARCHAR (required)
- `email`: VARCHAR (required, unique)
- `phone`: VARCHAR (optional)
- `position`: VARCHAR (optional)
- `department`: VARCHAR (optional)
- `address`: TEXT (optional)
- `createdAt`: TIMESTAMP
- `updatedAt`: TIMESTAMP
