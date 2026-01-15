const XLSX = require('xlsx');
const { body, validationResult } = require('express-validator');

// Required columns for office bearers Excel file
const REQUIRED_COLUMNS = ['Name', 'Email'];
const OPTIONAL_COLUMNS = ['Phone', 'Position', 'Department', 'Address'];

// Email validation regex
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * Parse Excel file buffer and extract data
 */
function parseExcelFile(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { raw: false });
    return data;
  } catch (error) {
    throw new Error('Failed to parse Excel file: ' + error.message);
  }
}

/**
 * Normalize column names (remove spaces, convert to lowercase for comparison)
 */
function normalizeColumnName(name) {
  return name ? name.toString().trim() : '';
}

/**
 * Validate Excel structure - check if required columns are present
 */
function validateColumns(row, rowIndex) {
  const errors = [];
  const foundColumns = {};
  const normalizedColumns = {};

  // Normalize all column names for comparison
  Object.keys(row).forEach(key => {
    const normalized = normalizeColumnName(key);
    foundColumns[normalized] = key; // Store original key
    normalizedColumns[normalized] = true;
  });

  // Check for required columns (case-insensitive)
  const requiredColumnsLower = REQUIRED_COLUMNS.map(col => col.toLowerCase());
  
  for (const requiredCol of REQUIRED_COLUMNS) {
    const normalizedRequired = requiredCol.toLowerCase();
    let found = false;
    
    // Check if any column matches (case-insensitive)
    for (const foundCol of Object.keys(normalizedColumns)) {
      if (foundCol.toLowerCase() === normalizedRequired) {
        found = true;
        break;
      }
    }
    
    if (!found) {
      errors.push({
        type: 'missing_column',
        column: requiredCol,
        message: `Required column "${requiredCol}" is missing from the Excel file`,
        row: rowIndex
      });
    }
  }

  return { errors, foundColumns };
}

/**
 * Validate a single row of data
 */
function validateRow(row, rowIndex, columnMap) {
  const errors = [];
  const normalizedRow = {};

  // Create normalized row with correct column mapping
  Object.keys(columnMap).forEach(normalizedKey => {
    const originalKey = columnMap[normalizedKey];
    normalizedRow[normalizedKey] = row[originalKey] ? row[originalKey].toString().trim() : '';
  });

  // Find Name and Email columns (case-insensitive)
  let nameColumn = null;
  let emailColumn = null;

  Object.keys(row).forEach(key => {
    const normalized = normalizeColumnName(key).toLowerCase();
    if (normalized === 'name') nameColumn = key;
    if (normalized === 'email') emailColumn = key;
  });

  // Validate Name (required)
  if (!nameColumn || !row[nameColumn] || !row[nameColumn].toString().trim()) {
    errors.push({
      type: 'empty_field',
      field: 'Name',
      message: `Row ${rowIndex + 2}: Name field is required and cannot be empty`,
      row: rowIndex + 2
    });
  }

  // Validate Email (required and format)
  if (!emailColumn || !row[emailColumn] || !row[emailColumn].toString().trim()) {
    errors.push({
      type: 'empty_field',
      field: 'Email',
      message: `Row ${rowIndex + 2}: Email field is required and cannot be empty`,
      row: rowIndex + 2
    });
  } else {
    const emailValue = row[emailColumn].toString().trim();
    if (!EMAIL_REGEX.test(emailValue)) {
      errors.push({
        type: 'invalid_format',
        field: 'Email',
        message: `Row ${rowIndex + 2}: Invalid email format: "${emailValue}"`,
        row: rowIndex + 2,
        value: emailValue
      });
    }
  }

  return { errors, normalizedRow };
}

/**
 * Main validation function
 */
function validateExcelData(buffer) {
  const validationErrors = [];
  let parsedData = [];
  let columnMap = {};

  try {
    // Parse Excel file
    parsedData = parseExcelFile(buffer);

    if (!parsedData || parsedData.length === 0) {
      validationErrors.push({
        type: 'empty_file',
        message: 'Excel file is empty or contains no data rows'
      });
      return { isValid: false, errors: validationErrors, data: [] };
    }

    // Validate columns in first row
    if (parsedData.length > 0) {
      const columnValidation = validateColumns(parsedData[0], 0);
      validationErrors.push(...columnValidation.errors);
      columnMap = columnValidation.foundColumns;

      if (validationErrors.length > 0) {
        return { isValid: false, errors: validationErrors, data: [] };
      }
    }

    // Validate each row
    for (let i = 0; i < parsedData.length; i++) {
      const rowValidation = validateRow(parsedData[i], i, columnMap);
      validationErrors.push(...rowValidation.errors);
    }

    if (validationErrors.length > 0) {
      return { isValid: false, errors: validationErrors, data: [] };
    }

    // Transform data to match database schema
    const transformedData = parsedData.map((row, index) => {
      const result = {};
      
      // Find columns (case-insensitive)
      Object.keys(row).forEach(key => {
        const normalized = normalizeColumnName(key).toLowerCase();
        const value = row[key] ? row[key].toString().trim() : '';
        
        if (normalized === 'name') result.name = value;
        else if (normalized === 'email') result.email = value;
        else if (normalized === 'phone') result.phone = value || null;
        else if (normalized === 'position') result.position = value || null;
        else if (normalized === 'department') result.department = value || null;
        else if (normalized === 'address') result.address = value || null;
      });

      return result;
    });

    return {
      isValid: true,
      errors: [],
      data: transformedData
    };
  } catch (error) {
    validationErrors.push({
      type: 'parsing_error',
      message: error.message || 'Error parsing Excel file'
    });
    return { isValid: false, errors: validationErrors, data: [] };
  }
}

module.exports = {
  validateExcelData,
  REQUIRED_COLUMNS,
  OPTIONAL_COLUMNS
};
